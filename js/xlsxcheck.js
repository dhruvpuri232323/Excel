// Reads a workbook the learner made in real Excel (.xlsx) and checks it against a rubric.
// Formulas are recalculated with Sheetah's own engine, so checks can also run on hidden data.
import { parseSection } from './m.js';
import { checkPQ } from './pqcheck.js';
import { readVbaProject } from './ovba.js';
import { checkVBA, makeHost } from './vbacheck.js';
import { Program } from './vba.js';
import { Workbook, parseAddr, parseRange, addr, colIndex, colName, shiftFormula, isErr, eachCell, compareValues, isDateFmt, rangeStr } from './engine.js';

const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '₹#,##0;-₹#,##0', 6: '₹#,##0;[Red]-₹#,##0', 7: '₹#,##0.00;-₹#,##0.00', 8: '₹#,##0.00;[Red]-₹#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'dd-mm-yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'dd-mm-yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 44: '_(₹* #,##0.00_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const parseXML = t => new DOMParser().parseFromString(t, 'application/xml');
const kids = (node, name) => node ? [...node.getElementsByTagNameNS('*', name)] : [];
const one = (node, name) => kids(node, name)[0] || null;
const attr = (n, a) => n ? n.getAttribute(a) : null;

// a chart part (xl/charts/chartN.xml, or chartExN.xml for waterfall, histogram, funnel …) → what Sheetah checks
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const runs = n => n ? kids(n, 't').map(t => t.textContent).join('').trim() : '';
const truthy = n => !!n && attr(n, 'val') !== '0' && attr(n, 'val') !== 'false';
function parseChart(cx, alt, ex) {
  if (ex) { const ss = kids(cx, 'series'); const title = one(cx, 'title'); const fOf = n => { const f = n && one(n, 'f'); return f ? f.textContent.trim() : null; };
    const data = {}; for (const d of kids(cx, 'data')) data[attr(d, 'id')] = { cat: fOf(one(d, 'strDim')), val: fOf(one(d, 'numDim')) };
    const txt = n => n ? (runs(n) || kids(n, 'v').map(v => v.textContent).join('')).trim() : '';
    const series = ss.map(q => { const d = data[attr(one(q, 'dataId'), 'val')] || {}; const tx = kids(q, 'tx').find(t => t.parentNode === q); return { name: txt(tx), layout: attr(q, 'layoutId'), cat: d.cat || null, val: d.val || null, refs: [d.cat, d.val, fOf(tx && one(tx, 'txData'))].filter(Boolean) }; });
    const bin = one(cx, 'binning'); const num = (n, a) => n && attr(n, a) !== null ? +attr(n, a) : null;
    return { ex: true, kinds: [...new Set(ss.map(q => attr(q, 'layoutId')))], groups: [], title: txt(title && kids(title, 'tx')[0]), titleDeleted: !title, legend: one(cx, 'legend') ? (attr(one(cx, 'legend'), 'pos') || 'b') : null, alt, series,
      labels: kids(cx, 'dataLabels').some(d => { const v = one(d, 'visibility'); return !v || attr(v, 'value') !== '0' || attr(v, 'categoryName') === '1'; }), totals: kids(cx, 'subtotals').flatMap(t => kids(t, 'idx').map(i => +attr(i, 'val'))),
      binning: bin ? { width: num(one(bin, 'binSize'), 'val'), count: num(one(bin, 'binCount'), 'val'), underflow: num(bin, 'underflow'), overflow: num(bin, 'overflow') } : null, valAxes: [], points: 0 }; }
  const pa = one(cx, 'plotArea'); if (!pa) return null; const chartEl = one(cx, 'chart');
  const valAx = kids(pa, 'valAx').map(v => ({ id: attr(one(v, 'axId'), 'val'), min: one(one(v, 'scaling'), 'min') ? +attr(one(one(v, 'scaling'), 'min'), 'val') : null, max: one(one(v, 'scaling'), 'max') ? +attr(one(one(v, 'scaling'), 'max'), 'val') : null, grid: !!one(v, 'majorGridlines'), deleted: truthy(one(v, 'delete')), pos: attr(one(v, 'axPos'), 'val') }));
  const catAx = [...kids(pa, 'catAx'), ...kids(pa, 'dateAx')].map(v => ({ id: attr(one(v, 'axId'), 'val'), reverse: attr(one(one(v, 'scaling'), 'orientation'), 'val') === 'maxMin', deleted: truthy(one(v, 'delete')) }));
  const KIND = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', radarChart: 'radar', bubbleChart: 'bubble', stockChart: 'stock', surfaceChart: 'surface' };
  const groups = []; let points = 0;
  for (const el of [...pa.children]) { const k = KIND[el.localName]; if (!k) continue; const dir = attr(one(el, 'barDir'), 'val'); const ax = kids(el, 'axId').map(a => attr(a, 'val'));
    const ser = kids(el, 'ser').filter(q => q.parentNode === el).map(q => { const tx = one(q, 'tx'); const ref = n => { const e = one(q, n); const f = e ? one(e, 'f') : null; return f ? f.textContent : null; }; const pts = kids(q, 'dPt').filter(d => one(d, 'spPr')).length; points += pts; const dl = one(q, 'dLbls');
      return { name: tx ? (runs(tx) || kids(tx, 'v').map(v => v.textContent).join('') || (one(tx, 'f') || {}).textContent || '') : '', nameRef: tx && one(tx, 'f') ? one(tx, 'f').textContent : null, cat: ref('cat') || ref('xVal'), val: ref('val') || ref('yVal'), points: pts, labels: !!dl && !truthy(one(dl, 'delete')) && (truthy(one(dl, 'showVal')) || truthy(one(dl, 'showPercent')) || truthy(one(dl, 'showCatName'))) }; });
    const gdl = kids(el, 'dLbls').find(d => d.parentNode === el);
    groups.push({ kind: k === 'bar' ? (dir === 'bar' ? 'bar' : 'column') : k, grouping: attr(one(el, 'grouping'), 'val') || 'standard', axes: ax, secondary: k !== 'scatter' && k !== 'bubble' && valAx.length > 1 && ax.includes(valAx[1].id), series: ser, labels: ser.some(q => q.labels) || (!!gdl && (truthy(one(gdl, 'showVal')) || truthy(one(gdl, 'showPercent')))) }); }
  const t = one(chartEl, 'title'); const tdel = truthy(one(chartEl, 'autoTitleDeleted')); const lg = one(chartEl, 'legend');
  const series = groups.flatMap(g => g.series);
  let title = t ? (runs(one(t, 'tx')) || kids(t, 'v').map(v => v.textContent).join('')) : '';
  if (!t && !tdel && series.length === 1) title = series[0].name; // Excel shows the series name as the title
  const pvs = (cx.getElementsByTagName('c:pivotSource')[0] || cx.getElementsByTagName('pivotSource')[0]);
  return { ex: false, pivot: !!pvs, kinds: groups.map(g => g.kind), groups, title, titleDeleted: !t && (tdel || series.length !== 1), titleRef: t && one(t, 'f') ? one(t, 'f').textContent : null, legend: lg && !truthy(one(lg, 'delete')) ? (attr(one(lg, 'legendPos'), 'val') || 'r') : null, valAxes: valAx, catAxes: catAx, series, labels: groups.some(g => g.labels), points, alt };
}
export async function readXlsx(file) {
  const name = file.name || 'workbook.xlsx';
  if (/\.csv$/i.test(name)) return { error: 'This is a CSV file. CSV keeps only values, so formats, formulas and sheets are lost. Save as Excel Workbook (.xlsx) and upload again.', ext: 'csv' };
  if (/\.xls$/i.test(name)) return { error: 'This is an old .xls file. In Excel choose File › Save As › Excel Workbook (.xlsx), then upload that.', ext: 'xls' };
  if (/\.numbers$/i.test(name)) return { error: 'This is an Apple Numbers file. In Numbers choose File › Export To › Excel, then upload the .xlsx file.', ext: 'numbers' };
  if (/\.xlsb$/i.test(name)) return { error: 'This is a binary workbook (.xlsb). In Excel choose File › Save As › Excel Macro-Enabled Workbook (.xlsm), then upload that.', ext: 'xlsb' };
  if (/\.ods$/i.test(name)) return { error: 'This is an OpenDocument (LibreOffice) file. Choose File › Save As, pick “Excel 2007-365 (.xlsx)”, and upload that.', ext: 'ods' };
  if (file.size === 0) return { error: 'This file is empty. Save your workbook in Excel first, then upload it again.' };
  if (file.size > 25e6) return { error: 'This file is very large (over 25 MB). Upload the workbook from the task, without extra data or pictures.' };
  const head = new Uint8Array(await file.slice(0, 8).arrayBuffer());
  if (head[0] === 0xD0 && head[1] === 0xCF && head[2] === 0x11 && head[3] === 0xE0) return { error: 'Sheetah can’t open this workbook: it’s either an old .xls file or protected with a password to open. In Excel, remove the password (File › Info › Protect Workbook › Encrypt with Password, clear it), then File › Save As › Excel Workbook (.xlsx).', ext: 'cfb' };
  if (!(head[0] === 0x50 && head[1] === 0x4B)) return { error: 'This doesn’t look like an Excel workbook. In Excel choose File › Save As › Excel Workbook (.xlsx), then upload that file.' };
  let zip; try { zip = await JSZip.loadAsync(file); } catch (e) { return { error: 'This file couldn’t be opened as an Excel workbook. It may be damaged. Open it in Excel, choose File › Save As › Excel Workbook (.xlsx), and try again.' }; }
  if (!zip.file('xl/workbook.xml')) return { error: zip.file('content.xml') ? 'This is an OpenDocument (LibreOffice) file. Save it as Excel 2007-365 (.xlsx) and upload that.' : zip.file('word/document.xml') ? 'This is a Word document, not an Excel workbook.' : 'This zip file isn’t an Excel workbook. Upload the .xlsx file you saved from Excel.' };
  const text = async p => { const f = zip.file(p.replace(/^\//, '')); return f ? f.async('text') : null; };
  const rels = async p => { const dir = p.replace(/[^/]+$/, ''); const t = await text(dir + '_rels/' + p.split('/').pop() + '.rels'); const out = {}; if (t) for (const r of kids(parseXML(t), 'Relationship')) { let tgt = attr(r, 'Target'); if (!tgt.startsWith('/')) tgt = norm(dir + tgt); else tgt = tgt.slice(1); out[attr(r, 'Id')] = { target: tgt, type: attr(r, 'Type') }; } return out; };
  const wbXml = parseXML(await text('xl/workbook.xml')); const wbRels = await rels('xl/workbook.xml');
  const ss = []; const sst = await text('xl/sharedStrings.xml'); if (sst) for (const si of kids(parseXML(sst), 'si')) ss.push(kids(si, 't').map(t => t.textContent).join(''));
  // styles
  const styles = { xfs: [], fonts: [], fills: [], fmts: { ...BUILTIN_FMT } }; const stx = await text('xl/styles.xml');
  if (stx) { const d = parseXML(stx); for (const nf of kids(d, 'numFmt')) styles.fmts[+attr(nf, 'numFmtId')] = attr(nf, 'formatCode');
    styles.fonts = kids(one(d, 'fonts'), 'font').map(f => ({ bold: !!one(f, 'b') && attr(one(f, 'b'), 'val') !== '0', italic: !!one(f, 'i') && attr(one(f, 'i'), 'val') !== '0' }));
    styles.fills = kids(one(d, 'fills'), 'fill').map(f => { const p = one(f, 'patternFill'); const t = attr(p, 'patternType'); return { solid: !!t && t !== 'none' && t !== 'gray125', color: attr(one(p, 'fgColor'), 'rgb') || attr(one(p, 'fgColor'), 'theme') }; });
    styles.xfs = kids(one(d, 'cellXfs'), 'xf').map(x => { const al = one(x, 'alignment'), pr = one(x, 'protection'); return { fmt: styles.fmts[+attr(x, 'numFmtId') || 0] || 'General', font: +attr(x, 'fontId') || 0, fill: +attr(x, 'fillId') || 0, align: attr(al, 'horizontal'), wrap: attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true', locked: !(pr && (attr(pr, 'locked') === '0' || attr(pr, 'locked') === 'false')) }; }); }
  const names = {}; for (const dn of kids(wbXml, 'definedName')) names[attr(dn, 'name')] = { ref: dn.textContent, local: attr(dn, 'localSheetId') };
  const sheets = [];
  for (const sh of kids(wbXml, 'sheet')) {
    const rid = attr(sh, 'r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'); const path = wbRels[rid] && wbRels[rid].target; if (!path) continue;
    const x = parseXML(await text(path)); const srels = await rels(path);
    const s = { name: attr(sh, 'name'), state: attr(sh, 'state') || 'visible', dataTables: [], scenarios: kids(x, 'scenario').map(q => attr(q, 'name')), cells: {}, merges: kids(x, 'mergeCell').map(m => attr(m, 'ref')), cols: {}, validations: [], tables: [], comments: [], protect: null };
    for (const col of kids(x, 'col')) { const w = +attr(col, 'width'); for (let c = +attr(col, 'min'); c <= +attr(col, 'max') && c < 200; c++) s.cols[c - 1] = w; }
    const pane = one(x, 'pane'); if (pane && /frozen/.test(attr(pane, 'state') || '')) s.freeze = { rows: +(attr(pane, 'ySplit') || 0), cols: +(attr(pane, 'xSplit') || 0) };
    const shared = {}; const arrays = []; // spilled answers: Excel saves the formula once, with the cells it fills
    for (const c of kids(x, 'c')) {
      const r = attr(c, 'r'); const t = attr(c, 't'); const sIdx = +(attr(c, 's') || 0); const vEl = one(c, 'v'), fEl = one(c, 'f'); let v = vEl ? vEl.textContent : null; let f = fEl ? fEl.textContent : null;
      if (t === 's') v = ss[+v]; else if (t === 'b') v = v === '1'; else if (t === 'inlineStr') v = kids(c, 't').map(n => n.textContent).join(''); else if (t === 'e') v = { err: v }; else if (t === 'str') v = v; else if (v !== null) v = +v;
      if (fEl && attr(fEl, 't') === 'array' && /:/.test(attr(fEl, 'ref') || '')) arrays.push({ at: r, ref: attr(fEl, 'ref') });
      if (fEl && attr(fEl, 't') === 'dataTable') { s.dataTables.push({ ref: attr(fEl, 'ref'), twoWay: attr(fEl, 'dt2D') === '1' || attr(fEl, 'dt2D') === 'true' }); f = null; } // What-If › Data Table: Excel keeps the results as values
      if (fEl && attr(fEl, 't') === 'shared') { const si = attr(fEl, 'si'); if (f) shared[si] = { f, at: parseAddr(r) }; else if (shared[si]) { const p = parseAddr(r); f = shiftFormula('=' + shared[si].f, p.r - shared[si].at.r, p.c - shared[si].at.c).slice(1); } }
      const xf = styles.xfs[sIdx] || { fmt: 'General', font: 0, fill: 0, locked: true };
      s.cells[r] = { v, f: f ? '=' + f : null, fmt: xf.fmt, bold: (styles.fonts[xf.font] || {}).bold, italic: (styles.fonts[xf.font] || {}).italic, fill: (styles.fills[xf.fill] || {}).solid, align: xf.align, wrap: xf.wrap, locked: xf.locked };
    }
    for (const a of arrays) { const g = parseRange(a.ref); if (g) eachCell(g, (rr, cc) => { const k = addr(rr, cc); if (k !== a.at && s.cells[k] && !s.cells[k].f) s.cells[k].v = null; }); } // Sheetah re-spills them itself
    for (const dv of kids(x, 'dataValidation')) { const f1 = one(dv, 'formula1'); s.validations.push({ type: attr(dv, 'type'), sqref: attr(dv, 'sqref'), formula: f1 ? f1.textContent : '' }); }
    // newer Excel writes list validations in an extension block (x14:dataValidation)
    const sp = one(x, 'sheetProtection'); if (sp) s.protect = { sheet: attr(sp, 'sheet') === '1' || attr(sp, 'sheet') === 'true' };
    const ps = one(x, 'pageSetup'); s.page = { orientation: attr(ps, 'orientation') || 'portrait', fitW: attr(ps, 'fitToWidth'), fitH: attr(ps, 'fitToHeight'), fitToPage: !!(one(x, 'pageSetUpPr') && attr(one(x, 'pageSetUpPr'), 'fitToPage') === '1'), scale: attr(ps, 'scale') };
    const hf = one(x, 'headerFooter'); s.headerFooter = hf ? hf.textContent : '';
    const af = one(x, 'autoFilter'); s.autoFilter = af ? attr(af, 'ref') : null;
    const tc = one(x, 'tabColor'); s.tabColor = tc ? (attr(tc, 'rgb') || attr(tc, 'theme') || 'set') : null;
    for (const tp of kids(x, 'tablePart')) { const rel = srels[attr(tp, 'r:id') || tp.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')]; if (!rel) continue; const tx = parseXML(await text(rel.target)); const tb = tx.documentElement; let conn = null; if (attr(tb, 'tableType') === 'queryTable') { const qr = Object.values(await rels(rel.target)).find(r => /queryTable/.test(r.type || '')); const qx = qr && await text(qr.target); conn = qx ? (/connectionId="(\d+)"/.exec(qx) || [])[1] || null : null; }
      s.tables.push({ name: attr(tb, 'displayName') || attr(tb, 'name'), ref: attr(tb, 'ref'), query: attr(tb, 'tableType') === 'queryTable', conn, totals: +(attr(tb, 'totalsRowCount') || 0) > 0 || attr(tb, 'totalsRowShown') === '1' && +(attr(tb, 'totalsRowCount') || 0) > 0, columns: kids(tx, 'tableColumn').map(c => attr(c, 'name')) }); }
    s.pivots = []; // PivotTables: which fields sit in Rows, Columns, Filters and Values, how values are shown, calculated fields and groupings
    for (const rel of Object.values(srels)) if (/\/pivotTable$/.test(rel.type || '')) {
      const pt = await text(rel.target); if (!pt) continue; const px = parseXML(pt); const pr = await rels(rel.target); const cdRel = Object.values(pr).find(r => /pivotCacheDefinition$/.test(r.type || ''));
      const cdText = cdRel ? await text(cdRel.target) : null; const cx = cdText ? parseXML(cdText) : null;
      const cf = cx ? kids(cx, 'cacheField').map(c => { const fg = one(c, 'fieldGroup'); const rp = fg ? one(fg, 'rangePr') : null; return { name: attr(c, 'name') || '', formula: attr(c, 'formula'), group: fg && (one(fg, 'rangePr') || one(fg, 'groupItems') || one(fg, 'discretePr')) ? (rp ? (attr(rp, 'groupBy') || 'range') : 'items') : null }; }) : [];
      const nm = i => i === -2 ? 'Values' : (cf[i] || {}).name || ''; const area = tag => { const e = one(px, tag); return e ? kids(e, 'field').map(fl => nm(+attr(fl, 'x'))) : []; };
      const pf = one(px, 'pageFields'), df = one(px, 'dataFields');
      s.pivots.push({ name: attr(px.documentElement, 'name'), location: attr(one(px, 'location'), 'ref'), rows: area('rowFields'), cols: area('colFields'), filters: pf ? kids(pf, 'pageField').map(q => nm(+attr(q, 'fld'))) : [],
        values: df ? kids(df, 'dataField').map(d => { const x14 = d.getElementsByTagName('x14:dataField')[0]; return { name: attr(d, 'name') || '', field: nm(+attr(d, 'fld')), fn: attr(d, 'subtotal') || 'sum', showAs: (x14 && attr(x14, 'pivotShowAs')) || attr(d, 'showDataAs') || 'normal' }; }) : [],
        calc: cf.filter(c => c.formula).map(c => c.name), groups: cf.filter(c => c.group).map(c => ({ field: c.name, by: c.group })) });
    }
    s.charts = []; // charts on the sheet's drawing, with their alt text
    for (const rel of Object.values(srels)) if (/\/drawing$/.test(rel.type || '')) {
      const dt = await text(rel.target); if (!dt) continue; const dx = parseXML(dt); const drels = await rels(rel.target);
      for (const gf of kids(dx, 'graphicFrame')) { const nv = one(gf, 'cNvPr'); const ch = kids(gf, 'chart')[0]; if (!ch) continue; const rid = attr(ch, 'r:id') || ch.getAttributeNS(RNS, 'id'); const cr = drels[rid]; if (!cr) continue;
        const ct = await text(cr.target); if (!ct) continue; const c = parseChart(parseXML(ct), (attr(nv, 'descr') || attr(nv, 'title') || '').trim(), /chartEx/i.test((cr.type || '') + cr.target)); if (c) { c.name = attr(nv, 'name') || ''; if (c.ex) { const res = r => { if (!r || !/^_xlchart\./i.test(r)) return r; const k = Object.keys(names).find(n => n.toLowerCase() === r.toLowerCase()); return k ? names[k].ref : r; }; c.series.forEach(q => { q.cat = res(q.cat); q.val = res(q.val); q.refs = q.refs.map(res); }); } s.charts.push(c); } }
    }
    for (const rel of Object.values(srels)) if (/comments/.test(rel.type)) { const cx = await text(rel.target); if (cx) for (const cm of kids(parseXML(cx), 'comment')) s.comments.push(attr(cm, 'ref')); }
    for (const rel of Object.values(srels)) if (/threadedComment/.test(rel.type)) { const cx = await text(rel.target); if (cx) for (const cm of kids(parseXML(cx), 'threadedComment')) s.comments.push(attr(cm, 'ref')); }
    s.cf = []; for (const cfm of kids(x, 'conditionalFormatting')) for (const r of kids(cfm, 'cfRule')) { const is = one(r, 'iconSet'); s.cf.push({ sqref: attr(cfm, 'sqref'), type: attr(r, 'type'), operator: attr(r, 'operator'), bottom: attr(r, 'bottom') === '1', aboveAverage: attr(r, 'aboveAverage') !== '0', rank: +attr(r, 'rank') || null, text: attr(r, 'text'), iconSet: is ? attr(is, 'iconSet') || '3TrafficLights1' : null, colors: kids(r, 'color').map(c => attr(c, 'rgb') || ''), formulas: kids(r, 'formula').map(f => f.textContent) }); }
    s.sparklines = [...x.getElementsByTagName('x14:sparklineGroup')].map(gp => ({ type: { column: 'column', stacked: 'winloss' }[attr(gp, 'type')] || 'line', high: attr(gp, 'high') === '1', low: attr(gp, 'low') === '1', negative: attr(gp, 'negative') === '1', markers: attr(gp, 'markers') === '1',
      items: [...gp.getElementsByTagName('x14:sparkline')].map(q => ({ f: (q.getElementsByTagName('xm:f')[0] || {}).textContent || '', sqref: (q.getElementsByTagName('xm:sqref')[0] || {}).textContent || '' })) }));
    const ext = x.getElementsByTagName('x14:dataValidation'); for (const dv of [...ext]) { const f1 = dv.getElementsByTagName('xm:f')[0]; const sq = dv.getElementsByTagName('xm:sqref')[0]; s.validations.push({ type: attr(dv, 'type'), sqref: sq ? sq.textContent : '', formula: f1 ? f1.textContent : '' }); }
    sheets.push(s);
  }
  for (const sh of sheets) for (const c of sh.charts || []) for (const q of c.series) if (q.nameRef && (!q.name || q.name === q.nameRef)) { // a series named by a cell, saved without its cached text
    const m = /^(?:'?(.+?)'?!)?\$?([A-Z]+)\$?(\d+)$/.exec(q.nameRef.replace(/^=/, '')); const src = m ? (m[1] ? sheets.find(z => z.name === m[1]) : sh) : null; const cl = src && src.cells[m[2] + m[3]]; if (cl && cl.v !== null) q.name = String(cl.v); }
  for (const sh of sheets) for (const c of sh.charts || []) if (!c.title && !c.titleDeleted && c.series.length === 1) c.title = c.series[0].name;
  const props = await text('docProps/core.xml');
  const slicerLinks = []; for (const f of Object.keys(zip.files).filter(f => /^xl\/slicerCaches\/[^/]+\.xml$/.test(f))) { const t = await zip.file(f).async('string'); slicerLinks.push({ field: (/sourceName="([^"]*)"/.exec(t) || [])[1] || '', pivots: (t.match(/<(?:\w+:)?pivotTable /g) || []).length }); }
  const files = Object.keys(zip.files); const slicers = files.filter(f => /^xl\/slicers\/[^/]+\.xml$/.test(f)).length, timelines = files.filter(f => /^xl\/timelines\/[^/]+\.xml$/.test(f)).length;
  // Power Query: the queries live in a DataMashup package (customXml), and connections.xml says how each one loads
  const queries = {}, queryLoads = {}; try { for (const f of Object.keys(zip.files).filter(p => /^customXml\/item\d*\.xml$/i.test(p))) { const raw = await zip.file(f).async('uint8array'); const t = raw[0] === 0xFF && raw[1] === 0xFE ? new TextDecoder('utf-16le').decode(raw) : new TextDecoder('utf-8').decode(raw);
      const m = /<DataMashup[^>]*>([\s\S]*?)<\/DataMashup>/.exec(t); if (!m) continue; const b64 = m[1].replace(/\s+/g, ''); const bin = typeof atob === 'function' ? Uint8Array.from(atob(b64), c => c.charCodeAt(0)) : new Uint8Array(Buffer.from(b64, 'base64'));
      const n = new DataView(bin.buffer, bin.byteOffset).getUint32(4, true); const pkg = await JSZip.loadAsync(bin.slice(8, 8 + n)); const sec = pkg.file('Formulas/Section1.m'); if (sec) Object.assign(queries, parseSection(await sec.async('text'))); }
    const cx = await text('xl/connections.xml'); const allTables = sheets.flatMap(s => s.tables); const qtables = allTables.filter(t => t.query).map(t => t.name.toLowerCase()); const anyTables = allTables.map(t => t.name.toLowerCase());
    const blockOf = q => { if (!cx) return ''; const at = cx.indexOf(`name="Query - ${q.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`); if (at < 0) return ''; const st = cx.lastIndexOf('<connection', at); const en = cx.indexOf('</connection>', at); return cx.slice(st, en < 0 ? at + 200 : en); };
    // loaded to a table: a query table linked to the query's connection, else a table named after the query (Excel names it so; a renamed one still counts by its link)
    for (const q of Object.keys(queries)) { const blk = blockOf(q); const id = (/<connection[^>]*\sid="(\d+)"/.exec(blk) || [])[1]; const tn = q.replace(/[^A-Za-z0-9_.]/g, '_').toLowerCase(); const named = x => x === tn || new RegExp('^' + tn.replace(/\./g, '\\.') + '_\\d+$').test(x);
      const tbl = (id && allTables.some(t => t.query && t.conn === id)) || qtables.some(named) || anyTables.some(x => x === tn);
      if (blk || tbl) queryLoads[q] = { to: tbl ? 'table' : 'connection', model: /model="1"/.test(blk) }; } } catch (e) { /* not a Power Query workbook */ }
  // the Data Model's tables and relationships, which Excel lists in workbook.xml (x15:dataModel)
  const byLocal = (doc, n) => [...doc.getElementsByTagName('*')].filter(e => e.localName === n);
  const modelTables = byLocal(wbXml, 'modelTable').map(e => attr(e, 'name')); const modelRels = byLocal(wbXml, 'modelRelationship').map(e => ({ fromTable: attr(e, 'fromTable'), fromColumn: attr(e, 'fromColumn'), toTable: attr(e, 'toTable'), toColumn: attr(e, 'toColumn') }));
  let vba = null, vbaError = null; const vf = zip.file('xl/vbaProject.bin'); if (vf) { try { vba = readVbaProject(await vf.async('uint8array')); } catch (e) { vbaError = e.message; } } // the macros in an .xlsm: [{name, kind, code}]
  return { name, sheets, names, slicers, timelines, slicerLinks, queries, queryLoads, modelTables, modelRels, author: props ? ((/<dc:creator>([^<]*)/.exec(props) || [])[1] || '') : '', hasMacros: !!vf, vba, vbaError, ext: (/\.([a-z]+)$/i.exec(name) || [])[1] ? (/\.([a-z]+)$/i.exec(name) || [])[1].toLowerCase() : 'xlsx' };
}
function norm(p) { const out = []; for (const part of p.split('/')) { if (part === '..') out.pop(); else if (part && part !== '.') out.push(part); } return out.join('/'); }

// Replace structured references (Table1[Amount], Table1[[#Totals],[Amount]], [@Amount]) with A1 ranges
function resolveStructured(f, tables, cellAddr) {
  if (!f || !f.includes('[')) return f;
  const here = parseAddr(cellAddr);
  const find = n => tables.find(t => t.name.toLowerCase() === n.toLowerCase());
  const own = tables.find(t => { const g = parseRange(t.ref); return here && here.r >= g.r1 && here.r <= g.r2 && here.c >= g.c1 && here.c <= g.c2; });
  const colRange = (t, col, part) => { const g = parseRange(t.ref); const i = t.columns.findIndex(c => c.toLowerCase() === col.toLowerCase().replace(/'(.)/g, '$1')); if (i < 0) return null; const c = colName(g.c1 + i); const dataStart = g.r1 + 2, dataEnd = g.r2 + 1 - (t.totals ? 1 : 0); if (part === '#totals') return c + (g.r2 + 1); if (part === '#headers') return c + (g.r1 + 1); if (part === '#all') return `${c}${g.r1 + 1}:${c}${g.r2 + 1}`; if (part === '@') return c + (here.r + 1); return `${c}${dataStart}:${c}${dataEnd}`; };
  return f.replace(/([A-Za-z_][A-Za-z0-9_.]*)?\[(\[?#?@?[^\]]*\]?(?:,\s*\[[^\]]*\])*)\]/g, (m, tname, inner) => {
    const t = tname ? find(tname) : own; if (!t) return m;
    let part = '', col = inner;
    const parts = inner.split(/\],\s*\[/).map(x => x.replace(/^\[|\]$/g, ''));
    if (parts.length > 1) { part = parts[0].toLowerCase(); col = parts[1]; }
    else if (inner.startsWith('@')) { part = '@'; col = inner.slice(1).replace(/^\[|\]$/g, ''); }
    else if (/^#/.test(inner)) { const g = parseRange(t.ref); if (/#data/i.test(inner) || /#all/i.test(inner)) return `${addr(g.r1 + (/#all/i.test(inner) ? 0 : 1), g.c1)}:${addr(g.r2 - (t.totals && !/#all/i.test(inner) ? 1 : 0), g.c2)}`; return m; }
    if (part === '#this row') part = '@';
    const r = colRange(t, col, part); return r || m;
  });
}
export function toWorkbook(x) {
  const spec = { sheets: x.sheets.map(s => { const cells = {}; const allTables = x.sheets.flatMap(q => q.tables); for (const [a, c] of Object.entries(s.cells)) { if (c.f) cells[a] = { f: resolveStructured(c.f, allTables, a).replace(/_xlfn\./g, '').replace(/_xlws\./g, ''), fmt: c.fmt !== 'General' ? c.fmt : undefined }; else if (c.v !== null && c.v !== undefined) cells[a] = { v: c.v && c.v.err ? null : c.v, fmt: c.fmt !== 'General' ? c.fmt : undefined }; } return { name: s.name, cells, tables: (s.tables || []).map(t => ({ name: t.name, ref: t.ref, totals: t.totals })) }; }), names: {} };
  for (const [n, d] of Object.entries(x.names)) if (!n.startsWith('_xlnm')) spec.names[n] = d.ref;
  const wb = new Workbook(spec);
  // remember Excel's own cached results to fall back on for functions Sheetah doesn't know
  x.sheets.forEach((s, si) => { for (const [a, c] of Object.entries(s.cells)) if (c.f) { const p = parseAddr(a); const o = wb.cellObj(wb.sheets[si], p.r, p.c); if (o) o.cached = c.v; } });
  const orig = wb.value.bind(wb); wb.value = (si, r, c) => { const v = orig(si, r, c); if (isErr(v) && (v.code === '#NAME?' || v.code === '#VALUE!')) { const o = wb.cellObj(wb.sheets[si], r, c); if (o && o.cached !== undefined && o.cached !== null && !(o.cached && o.cached.err)) return o.cached; } return v; };
  return wb;
}

const near = (a, b, tol = 0.005) => typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= tol * Math.max(1, Math.abs(b)) : String(a ?? '').trim().toLowerCase() === String(b ?? '').trim().toLowerCase();
const overlaps = (sqref, range) => { const g = parseRange(range); return String(sqref || '').split(/\s+/).some(part => { const h = parseRange(part); return h && h.r1 <= g.r2 && h.r2 >= g.r1 && h.c1 <= g.c2 && h.c2 >= g.c1; }); };

// rubric: [{label, type, points, ...}]
// vbarun: run the learner's macro and the model answer on fresh copies of the uploaded workbook, with Sheetah's own copy of the files where its Settings sheet says they are
export function vbaRun(item, spec, mods, name) { const W = new Workbook(JSON.parse(JSON.stringify(spec)));
  const set = a => { const q = parseRange(a.replace(/^.*!/, '')); const sj = W.sheets.findIndex(z => z.name.toLowerCase() === (a.split('!')[0] || '').toLowerCase()); const val = sj < 0 ? '' : W.value(sj, q.r1, q.c1); return String(val ?? '').trim().replace(/\\+$/, ''); };
  const places = {}; for (const [k, a] of Object.entries(item.places || {})) places[k] = set(a) || item.defaults[k];
  const fill = (p, P) => Object.entries(P).reduce((t, [k, v]) => t.split('{' + k + '}').join(v), p); const put = p => fill(p, places); const dflt = p => fill(p, item.defaults || {});
  // the files go where Settings says, and also next to the workbook (ThisWorkbook.Path), so code that looks there finds them too
  const mount = files => Object.fromEntries(Object.entries(files || {}).flatMap(([p, f]) => put(p).toLowerCase() === dflt(p).toLowerCase() ? [[put(p), f]] : [[put(p), f], [dflt(p), f]]));
  const mountDirs = ds => [...new Set(ds.flatMap(d => [put(d), dflt(d)]))];
  const lab = { sheet: spec, vba: { modules: [], files: mount(item.files), folders: mountDirs(item.folders || []), book: name || 'Monthly report.xlsm', path: item.path || 'C:\\Users\\you\\Documents' } };
  const task = { kind: 'vba', id: 'project', macro: item.macro, chain: item.chain, compare: item.compare, edits: item.edits, msol: item.msol, alsoWith: (item.alsoWith || []).map(w => ({ ...w, files: w.files ? mount(w.files) : undefined, folders: w.folders ? mountDirs(w.folders) : undefined })), whyMain: item.whyMain };
  const res = checkVBA(task, { vba: { lab, state: { modules: mods.map(m => ({ ...m })) } } });
  return res.ok ? { ok: true, msg: '' } : { ok: false, msg: /^\(Sheetah’s model answer/.test(res.msg) ? `Sheetah couldn’t run its own model answer on your workbook, so it can’t compare: keep the starter’s sheets and headings (Settings, Products, Orders) as they were. ${res.msg}` : res.msg }; }
// the macros in an uploaded workbook: standard modules run as they are; code in sheet, ThisWorkbook, class and form windows is listed but not run
function vbaOf(x) { if (!x.vba) return { msg: x.vbaError ? `The macros in this file couldn’t be read (${x.vbaError}). Open it in Excel, check the code is there, and save it again as .xlsm.` : x.ext === 'xlsx' ? 'This file has no macros: an .xlsx can’t keep VBA, so Excel removed them when it saved. In Excel choose File › Save As › Excel Macro-Enabled Workbook (.xlsm), check your modules are still there, and upload the .xlsm.' : 'There are no macros in this workbook. Add your modules (Insert › Module in the Visual Basic Editor) and save it as .xlsm.' };
  const all = x.vba.map(m => ({ name: m.name, kind: m.kind, code: m.code })); const mods = all.filter(m => m.kind === 'std').map(m => ({ name: m.name, kind: 'std', code: m.code }));
  if (!mods.length) return { msg: 'There’s no standard module in this workbook. In the Visual Basic Editor choose Insert › Module and put the macros there.' }; return { all, mods }; }
export function grade(x, rubric, opts = {}) {
  if (x.error) return { score: 0, total: 1, pct: 0, results: [{ label: 'Open the file', ok: false, msg: x.error, points: 1 }] };
  const wb = toWorkbook(x); const results = [];
  const S = n => { if (n === undefined) return 0; const i = x.sheets.findIndex(s => s.name.toLowerCase() === String(n).toLowerCase()); return i; };
  const cellInfo = (si, a) => (x.sheets[si] || { cells: {} }).cells[a.replace(/\$/g, '')] || null;
  for (const item of rubric) {
    const si = S(item.sheet); let ok = false, msg = '';
    const missingSheet = si < 0;
    try {
      if (missingSheet && item.type !== 'sheetName') { ok = false; msg = `I couldn’t find a sheet called “${item.sheet}”. Keep the sheet names from the starter file.`; }
      else switch (item.type) {
        case 'vbamods': case 'vbacode': case 'vbarun': { const v = vbaOf(x); if (v.msg) { ok = false; msg = v.msg; break; }
          if (item.type === 'vbamods') { const p = new Program(v.mods.map(m => ({ ...m })), makeHost({ vba: {} }, toWorkbook(x))); const ce = p.firstError(); if (ce) { ok = false; msg = `The code doesn’t compile: “${ce.message}”${ce.mod ? ` in ${ce.mod}` : ''}${ce.line ? `, line ${ce.line}` : ''}. In the Visual Basic Editor, Debug › Compile VBAProject takes you to it.`; break; }
            const have = new Map(); for (const m of v.mods) for (const mm of m.code.matchAll(/^[ \t]*(?:public[ \t]+|private[ \t]+)?sub[ \t]+([A-Za-z_]\w*)/gim)) have.set(mm[1].toLowerCase(), m.name);
            const other = new Map(); for (const m of v.all.filter(m => m.kind !== 'std')) for (const mm of m.code.matchAll(/^[ \t]*(?:public[ \t]+|private[ \t]+)?sub[ \t]+([A-Za-z_]\w*)/gim)) other.set(mm[1].toLowerCase(), m.name);
            const miss = item.subs.filter(n => !have.has(n.toLowerCase())); ok = !miss.length;
            msg = ok ? '' : other.has(miss[0].toLowerCase()) ? `${miss[0]} is in ${other.get(miss[0].toLowerCase())}’s code window. Macros that buttons and other macros call belong in a standard module: Insert › Module, and move it there.` : `There’s no Sub called ${miss.join(', ')} in a standard module yet.${miss.length < item.subs.length ? ' Check the spelling: the names must match the brief.' : ''}`; break; }
          if (item.type === 'vbacode') { const bad = v.mods.find(m => (item.each || []).some(n => !new RegExp(n.re, n.flags || 'im').test(m.code))); if (bad) { const n = item.each.find(n => !new RegExp(n.re, n.flags || 'im').test(bad.code)); ok = false; msg = `${bad.name}: ${n.msg}`; break; }
            for (const n of item.not || []) { const m = v.mods.find(m => new RegExp(n.re, n.flags || 'im').test(m.code.split(/\r?\n/).map(l => l.replace(/'.*$/, '')).join('\n'))); if (m) { ok = false; msg = `${m.name}: ${n.msg}`; break; } } if (msg) break; ok = true; break; }
          const res = vbaRun(item, toWorkbook(x).toSpec(), v.mods, x.name); ok = res.ok; msg = res.msg; break; }
        case 'modelrel': { // a relationship in the Data Model: rels [['Sales[Item Code]', 'Items[Item Code]'], …]
          const R = x.modelRels || []; const P = f => { const m = /^([^[]+)\[([^\]]+)\]$/.exec(f); return m ? [m[1].toLowerCase(), m[2].toLowerCase()] : ['', '']; }; const miss = (item.rels || []).find(([a, b]) => { const [ta, ca] = P(a), [tb, cb] = P(b); return !R.some(r => { const f = [String(r.fromTable).toLowerCase(), String(r.fromColumn).toLowerCase()], t = [String(r.toTable).toLowerCase(), String(r.toColumn).toLowerCase()]; return (f[0] === ta && f[1] === ca && t[0] === tb && t[1] === cb) || (f[0] === tb && f[1] === cb && t[0] === ta && t[1] === ca); }); });
          ok = !miss; msg = ok ? '' : (x.modelTables || []).length ? `The Data Model has no relationship between ${miss[0]} and ${miss[1]}. Data › Relationships › New, or drag the columns together in Power Pivot’s Diagram View.` : 'This workbook has no Data Model. Add the tables to it: Power Pivot › Add to Data Model, or tick “Add this data to the Data Model” when you load.'; break; }
        case 'pivotvalues': { // a PivotTable's numbers on the sheet: pairs [label, value]; each label's cell must have the value within 3 cells to its right
          const all = item.sheet ? [x.sheets[si]].filter(Boolean) : x.sheets; const near = (a, b) => typeof a === 'number' && Math.abs(a - b) <= Math.max(item.tol ?? 0.5, Math.abs(b) * 1e-6); let bad = null;
          for (const [label, want] of item.pairs || []) { let found = false; for (const sh of all) for (const [ad, c] of Object.entries(sh.cells)) { if (String(c.v ?? '').trim().toLowerCase() !== String(label).toLowerCase()) continue; const p = parseAddr(ad); for (let k = 1; k <= 3; k++) { const cc = sh.cells[addr(p.r, p.c + k)]; if (cc && near(cc.v, want)) found = true; } } if (!found) { bad = [label, want]; break; } }
          ok = !bad; msg = ok ? '' : `The PivotTable doesn’t show ${bad[1]} for ${bad[0]}. ${item.pivotMsg || 'Check its fields and the relationships between the tables.'}`; break; }
        case 'modeltables': { // tables in the Data Model (x15:modelTable names): a query loaded to the model, or a Table added with Add to Data Model
          const have = (x.modelTables || []).map(t => t.toLowerCase()); const miss = (item.tables || []).filter(t => !have.includes(t.toLowerCase()));
          ok = !miss.length; msg = ok ? '' : have.length ? `The Data Model doesn’t have ${miss.join(', ')} yet. Load ${miss.length > 1 ? 'them' : 'it'} with Close & Load To… › Only Create Connection and Add this data to the Data Model, or Power Pivot › Add to Data Model.` : 'This workbook has no Data Model. Load the queries with Close & Load To… › Add this data to the Data Model.'; break; }
        case 'pivotmeasures': { // explicit measures used in a PivotTable: a model PivotTable's value fields are [Measures].[Name]
          const vals = x.sheets.flatMap(q => (q.pivots || []).flatMap(pv => pv.values)); const has = m => vals.some(v => v.name.toLowerCase() === m.toLowerCase() || v.field.toLowerCase() === `[measures].[${m.toLowerCase()}]`);
          const miss = (item.measures || []).filter(m => !has(m)); ok = !miss.length; msg = ok ? '' : `No PivotTable uses the measure ${miss.map(m => `“${m}”`).join(', ')}. Create ${miss.length > 1 ? 'them' : 'it'} with Power Pivot › Measures › New Measure (exactly that name), then tick ${miss.length > 1 ? 'them' : 'it'} in PivotTable Fields.`; break; }
        case 'cubevalues': { // a report of CUBE formulas (OLAP Tools › Convert to Formulas): each label has a CUBEVALUE formula within 3 cells to its right showing the value
          const near = (a, b) => typeof a === 'number' && Math.abs(a - b) <= Math.max(0.5, Math.abs(b) * 1e-6); let bad = null, typed = false;
          for (const [label, want] of item.pairs || []) { let found = false; for (const sh of x.sheets) for (const [ad, c] of Object.entries(sh.cells)) { if (String(c.v ?? '').trim().toLowerCase() !== String(label).toLowerCase()) continue; const p = parseAddr(ad); for (let k = 1; k <= 3; k++) { const cc = sh.cells[addr(p.r, p.c + k)]; if (cc && near(cc.v, want)) { if (cc.f && /CUBEVALUE/i.test(cc.f)) found = true; else typed = true; } } } if (!found && !bad) bad = [label, want]; }
          ok = !bad; msg = ok ? '' : typed ? 'The numbers are right, but they’re in a PivotTable or typed in. Click in the PivotTable, then PivotTable Analyze › OLAP Tools › Convert to Formulas.' : `The report doesn’t show ${bad[1]} for ${bad[0]} with a CUBEVALUE formula. ${item.cubeMsg || ''}`; break; }
        case 'pq': case 'pqload': { const r = checkPQ({ ...item, kind: item.type }, { wb, pq: { queries: x.queries || {}, order: Object.keys(x.queries || {}), loads: x.queryLoads || {}, files: item.files, folders: item.folders, looseFolders: true, culture: item.culture, now: item.now } }); ok = r.ok; msg = ok ? '' : (Object.keys(x.queries || {}).length ? r.msg : 'This workbook has no Power Query queries. Build them with Data › Get Data, Close & Load, then save the workbook.'); break; }
        case 'value': { const p = parseAddr(item.cell); const v = wb.value(si, p.r, p.c); ok = item.approx ? near(v, item.equals, item.approx) : near(v, item.equals, 1e-6); msg = ok ? '' : `${item.cell} shows ${v ?? 'nothing'}; expected ${item.equals}.`; break; }
        case 'formula': { const c = cellInfo(si, item.cell); ok = !!(c && c.f) && (item.contains || []).every(t => c.f.toUpperCase().includes(t.toUpperCase())); if (ok && item.ref) { const p = parseAddr(item.cell); ok = near(wb.value(si, p.r, p.c), wb.evaluate(item.ref, si), 1e-6); } msg = ok ? '' : (c && !c.f ? `${item.cell} holds a typed value, not a formula.` : `${item.cell} needs a formula${item.contains ? ' using ' + item.contains.join(', ') : ''} that gives the right result.`); break; }
        case 'formulaRange': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { if (!ok) return; const a = addr(r, c); const ci = cellInfo(si, a); if (!ci || !ci.f) { ok = false; msg = `${a} needs a formula.`; return; } const ref = shiftFormula(item.ref, r - g.r1, c - g.c1); if (!near(wb.value(si, r, c), wb.evaluate(ref, si), 1e-6)) { ok = false; msg = `${a} gives the wrong result. Check the references, and whether any should be locked with $.`; } }); break; }
        case 'numFmt': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { if (!ok) return; const ci = cellInfo(si, addr(r, c)); if (!ci) return; const f = ci.fmt || 'General'; const good = (item.oneOf && item.oneOf.includes(f)) || (item.includes && item.includes.every(t => f.includes(t))) || (item.date && isDateFmt(f)) || (item.indian && /#,##,##0/.test(f)) || (item.regex && new RegExp(item.regex).test(f)); if (!good) { ok = false; msg = `${addr(r, c)} uses the format “${f}”. ${item.msg || ''}`; } }); break; }
        case 'bold': case 'fill': case 'wrap': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const ci = cellInfo(si, addr(r, c)); if (ok && (!ci || !ci[item.type])) { ok = false; msg = `${addr(r, c)} isn’t ${item.type === 'fill' ? 'filled with a colour' : item.type === 'wrap' ? 'set to wrap text' : 'bold'} yet.`; } }); break; }
        case 'align': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const ci = cellInfo(si, addr(r, c)); if (ok && (!ci || ci.align !== item.value)) { ok = false; msg = item.msg || `${addr(r, c)} isn’t aligned with ${item.value}.`; } }); break; }
        case 'noMerge': ok = !x.sheets[si].merges.some(m => !item.range || overlaps(m, item.range)); msg = ok ? '' : `Cells ${x.sheets[si].merges.join(', ')} are merged. Use Center Across Selection instead, so sorting and selecting still work.`; break;
        case 'colWidth': { const w = x.sheets[si].cols[colIndex(item.col)]; ok = !!w && w >= item.min; msg = ok ? '' : `Column ${item.col} is too narrow (${w ? w.toFixed(1) : 'default 8.43'} characters). Widen it to at least ${item.min}.`; break; }
        case 'freeze': { const f = x.sheets[si].freeze; ok = !!f && f.rows === (item.rows || 0) && f.cols === (item.cols || 0); msg = ok ? '' : (f ? `Panes are frozen at ${f.rows} row(s) and ${f.cols} column(s); the task needs ${item.rows || 0} and ${item.cols || 0}.` : 'Nothing is frozen yet. Use View › Freeze Panes.'); break; }
        case 'table': { const t = x.sheets[si].tables; ok = t.length > 0 && (!item.hasTotal || t.some(q => q.totals)) && (!item.name || t.some(q => q.name.toLowerCase() === item.name.toLowerCase())) && (!item.columns || t.some(q => item.columns.every(c => q.columns.map(z => z.toLowerCase()).includes(c.toLowerCase())))); msg = ok ? '' : (t.length ? (item.hasTotal ? 'The table has no Total Row. Tick Total Row on the Table Design tab.' : item.name ? `Name the table ${item.name} on the Table Design tab.` : 'The table is missing some columns.') : 'There is no Excel Table on this sheet yet. Click inside the data and press Ctrl+T.'); break; }
        case 'validationList': { const v = x.sheets[si].validations.filter(d => d.type === 'list' && overlaps(d.sqref, item.range)); ok = v.length > 0 && (!item.values || item.values.every(val => v.some(d => d.formula.toLowerCase().includes(val.toLowerCase()) || /^\$?[A-Z]+\$?\d+|!|^[A-Za-z_]/.test(d.formula)))); msg = ok ? '' : `${item.range} doesn’t have a drop-down list yet. Use Data › Data Validation › List.`; break; }
        case 'validation': { const v = x.sheets[si].validations.filter(d => (!item.kind || d.type === item.kind) && overlaps(d.sqref, item.range)); ok = v.length > 0; msg = ok ? '' : `${item.range} needs data validation (${item.kind || 'any type'}).`; break; }
        case 'protected': ok = !!(x.sheets[si].protect && x.sheets[si].protect.sheet); msg = ok ? '' : 'The sheet isn’t protected yet. Use Review › Protect Sheet.'; break;
        case 'unlocked': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const ci = cellInfo(si, addr(r, c)); if (ok && (!ci || ci.locked)) { ok = false; msg = `${addr(r, c)} is still locked, so nobody can type in it once the sheet is protected. Unlock input cells in Format Cells › Protection.`; } }); if (ok && !Object.keys(x.sheets[si].cells).some(a => { const p = parseAddr(a); const gg = parseRange(item.range); return p.r >= gg.r1 && p.r <= gg.r2 && p.c >= gg.c1 && p.c <= gg.c2; })) { /* empty input cells: style info missing; accept */ } break; }
        case 'orientation': ok = x.sheets[si].page.orientation === item.value; msg = ok ? '' : `The page is set to ${x.sheets[si].page.orientation}. Switch it to ${item.value} in Page Layout › Orientation.`; break;
        case 'fitWidth': ok = x.sheets[si].page.fitToPage && (x.sheets[si].page.fitW === null || x.sheets[si].page.fitW === '1'); msg = ok ? '' : 'Set Page Layout › Width to 1 page (Scale to Fit), so the columns print on one page.'; break;
        case 'fitPage': ok = x.sheets[si].page.fitToPage && (x.sheets[si].page.fitW === null || x.sheets[si].page.fitW === '1') && (x.sheets[si].page.fitH === null || x.sheets[si].page.fitH === '1'); msg = ok ? '' : 'Set Width and Height to 1 page in Page Layout › Scale to Fit, so it prints on one page.'; break;
        case 'printTitles': { const d = Object.entries(x.names).find(([n, d]) => n === '_xlnm.Print_Titles' && (d.local === null || +d.local === si)); ok = !!d && /\$1:\$1|\$\d+:\$\d+/.test(d[1].ref); msg = ok ? '' : 'Set Page Layout › Print Titles › Rows to repeat at top to row 1.'; break; }
        case 'headerFooter': ok = x.sheets[si].headerFooter.length > 0 && (!item.contains || item.contains.some(t => x.sheets[si].headerFooter.includes(t))); msg = ok ? '' : 'Add a header or footer (Insert › Header & Footer), for example page numbers.'; break;
        case 'sheetName': ok = x.sheets.some(s => s.name.toLowerCase() === item.name.toLowerCase()); msg = ok ? '' : `There’s no sheet named “${item.name}”. Double-click a sheet tab to rename it.`; break;
        case 'sheetCount': ok = x.sheets.length >= item.min; msg = ok ? '' : `The workbook has ${x.sheets.length} sheet(s); the task needs at least ${item.min}.`; break;
        case 'tabColor': ok = !!x.sheets[si].tabColor; msg = ok ? '' : 'Give the sheet tab a colour: right-click the tab › Tab Color.'; break;
        case 'sorted': { const g = parseRange(item.range); const c = colIndex(item.col); ok = true; let prev; for (let r = g.r1; r <= g.r2; r++) { const v = wb.value(si, r, c); if (r > g.r1 && v !== null && prev !== null) { const k = compareValues(prev, v); if (item.desc ? k < 0 : k > 0) { ok = false; msg = `Rows aren’t sorted by column ${item.col} (${item.desc ? 'largest first' : 'smallest first'}) — check row ${r + 1}.`; break; } } prev = v; } break; }
        case 'noDuplicates': { const g = parseRange(item.range); const seen = new Set(); ok = true; for (let r = g.r1; r <= g.r2; r++) { const key = []; for (let c = g.c1; c <= g.c2; c++) key.push(String(wb.value(si, r, c) ?? '').toLowerCase()); const k = key.join('|'); if (k.replace(/\|/g, '') === '') continue; if (seen.has(k)) { ok = false; msg = `Row ${r + 1} is a duplicate. Use Data › Remove Duplicates.`; break; } seen.add(k); } break; }
        case 'noBlankRows': { const g = parseRange(item.range); ok = true; let first = -1, last = -1; const blank = r => { for (let c = g.c1; c <= g.c2; c++) if (wb.value(si, r, c) !== null && wb.value(si, r, c) !== '') return false; return true; }; for (let r = g.r1; r <= g.r2; r++) if (!blank(r)) { if (first < 0) first = r; last = r; } for (let r = first; r <= last; r++) if (blank(r)) { ok = false; msg = `Row ${r + 1} is blank inside the data. Delete it.`; break; } break; }
        case 'rowCount': { const g = parseRange(item.range); let n = 0; for (let r = g.r1; r <= g.r2; r++) if (wb.value(si, r, g.c1) !== null && wb.value(si, r, g.c1) !== '') n++; ok = n === item.equals; msg = ok ? '' : `I counted ${n} rows of data in ${item.range}; there should be ${item.equals}.`; break; }
        case 'numbers': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const v = wb.value(si, r, c); if (ok && v !== null && typeof v !== 'number') { ok = false; msg = `${addr(r, c)} holds text (“${v}”), not a number.`; } }); break; }
        case 'textValues': { ok = Object.entries(item.cells).every(([a, want]) => { const p = parseAddr(a); const v = wb.value(si, p.r, p.c); return item.exact && typeof want === 'string' ? v === want : near(v, want); }); /* exact: capitals and spaces must match too */ msg = ok ? '' : item.msg || 'Some cells don’t hold the expected text yet.'; break; }
        case 'filled': { const p = parseAddr(item.cell); const v = wb.value(si, p.r, p.c); ok = v !== null && v !== ''; msg = ok ? '' : item.msg || `${item.cell} is still empty.`; break; }
        case 'filledRange': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const v = wb.value(si, r, c); if (ok && (v === null || v === '')) { ok = false; msg = item.msg || `${addr(r, c)} is still blank.`; } }); break; }
        case 'definedName': { const d = Object.entries(x.names).find(([n]) => n.toLowerCase() === item.name.toLowerCase()); ok = !!d && (!item.ref || d[1].ref.replace(/\$/g, '').toUpperCase().endsWith(item.ref.toUpperCase())); msg = ok ? '' : `Create a name called ${item.name}${item.ref ? ' for ' + item.ref : ''} (Formulas › Define Name, or the Name Box).`; break; }
        case 'comment': ok = x.sheets[si].comments.some(a => !item.cell || a === item.cell); msg = ok ? '' : `Add a note or comment${item.cell ? ' to ' + item.cell : ''}.`; break;
        case 'hiddenSheet': ok = x.sheets.some(s => s.state !== 'visible'); msg = ok ? '' : 'Hide the sheet named in the task (right-click its tab › Hide).'; break;
        case 'autoFilter': ok = !!x.sheets[si].autoFilter || x.sheets[si].tables.length > 0; msg = ok ? '' : 'Turn on filter buttons with Data › Filter.'; break;
        case 'variant': { // put hidden data into the inputs, recalculate, compare outputs with the model formulas
          ok = true; const sh0 = item.sheet || x.sheets[si].name; for (const dataset of item.datasets) { const w2 = toWorkbook(x); for (const [a, v] of Object.entries(dataset)) { const [sh, ad] = a.includes('!') ? a.split('!') : [sh0, a]; const s2 = w2.sheetIndexByName(sh); const p = parseAddr(ad); const o = w2.cellObj(w2.sheets[s2], p.r, p.c, true); o.v = v; o.f = null; } w2.invalidate(); for (const out of item.outputs) { const [sh, ad] = out.cell.includes('!') ? out.cell.split('!') : [sh0, out.cell]; const s2 = w2.sheetIndexByName(sh); const p = parseAddr(ad); const o = w2.cellObj(w2.sheets[s2], p.r, p.c); if (o) o.cached = undefined; const got = w2.value(s2, p.r, p.c); const want = w2.evaluate(out.ref, s2); if (!near(got, want, 1e-6)) { ok = false; msg = `${out.label || out.cell} gave ${got} on a hidden set of data, but should give ${want}. Check that it points at the whole input range.`; break; } } if (!ok) break; } break; }
        case 'spill': { // one formula whose spilled answer matches the model formula, cell by cell
          const p = parseAddr(item.cell); const ci = cellInfo(si, item.cell); ok = !!(ci && ci.f);
          if (!ok) { msg = `${item.cell} needs one formula that spills its answer.`; break; }
          const got = wb.spillArray(si, p.r, p.c), want = wb.evaluateArray(item.ref, si);
          if (got.length === 1 && got[0].length === 1 && isErr(got[0][0])) { ok = false; msg = `${item.cell} shows ${got[0][0].code}.${got[0][0].code === '#SPILL!' ? ' Clear the cells below and to the right so it can spill.' : ''}`; break; }
          if (got.length !== want.length || got[0].length !== want[0].length) { ok = false; msg = `${item.cell} spills ${got.length} × ${got[0].length}; it should be ${want.length} × ${want[0].length}.`; break; }
          for (let i = 0; i < want.length && ok; i++) for (let j = 0; j < want[0].length && ok; j++) if (!near(got[i][j], want[i][j], 1e-6)) { ok = false; msg = `${addr(p.r + i, p.c + j)} shows ${got[i][j] && got[i][j].code ? got[i][j].code : got[i][j]}; it should be ${want[i][j] && want[i][j].code ? want[i][j].code : want[i][j]}.`; }
          break; }
        case 'pivot': { // a PivotTable with these fields: rows/cols/filters ['Region' or 'OrderDate|Month'], values [{field, fn, showAs}]
          const all = item.sheet ? (x.sheets[si] ? x.sheets[si].pivots || [] : []) : x.sheets.flatMap(q => q.pivots || []); const alt = (want, have) => String(want).toLowerCase().split('|').some(w => have.some(h => h.toLowerCase() === w || h.toLowerCase().includes(w)));
          const miss = pv => { for (const a of ['rows', 'cols', 'filters']) for (const w of item[a] || []) if (!alt(w, pv[a])) return `Put ${w.split('|')[0]} in the ${{ rows: 'Rows', cols: 'Columns', filters: 'Filters' }[a]} area.`;
            for (const w of item.any || []) if (!alt(w, [...pv.rows, ...pv.cols, ...pv.filters])) return `Use ${w.split('|')[0]} in the PivotTable: in Rows, Columns or Filters.`;
            for (const v of item.values || []) { const hit = pv.values.filter(d => alt(v.field, [d.field])); if (!hit.length) return `Put ${v.field.split('|')[0]} in the Values area.`;
              if (v.fn && !hit.some(d => String(v.fn).toLowerCase().split('|').includes(d.fn.toLowerCase()))) return `Summarize ${v.field.split('|')[0]} by ${String(v.fn).split('|')[0].replace(/^./, c => c.toUpperCase())} (Value Field Settings).`;
              if (v.showAs && !hit.some(d => String(v.showAs).toLowerCase().split('|').includes(d.showAs.toLowerCase()))) return item.showAsMsg || `Change how ${v.field.split('|')[0]} is shown: Value Field Settings › Show Values As.`; }
            return ''; };
          if (!all.length) { msg = item.sheet ? `There’s no PivotTable on ${item.sheet}. Insert one with Insert › PivotTable and place it there.` : 'There’s no PivotTable in your file. Insert one with Insert › PivotTable.'; break; }
          const why = all.map(miss); ok = why.some(w => !w); msg = ok ? '' : (item.msg ? item.msg + ' ' : '') + why[0]; break; }
        case 'calcField': { const all = x.sheets.flatMap(q => q.pivots || []); ok = all.some(pv => pv.calc.length && (!item.name || pv.calc.some(c => c.toLowerCase() === item.name.toLowerCase()))); msg = ok ? '' : item.msg || `Add a calculated field${item.name ? ' called ' + item.name : ''}: PivotTable Analyze › Fields, Items & Sets › Calculated Field.`; break; }
        case 'pivotGroup': { const all = x.sheets.flatMap(q => q.pivots || []); ok = all.some(pv => pv.groups.some(g => (!item.by || String(item.by).split('|').includes(g.by)))); msg = ok ? '' : item.msg || 'Group the field in your PivotTable: right-click an item › Group.'; break; }
        case 'formulaUses': { const c = cellInfo(si, item.cell); ok = !!(c && c.f && String(c.f).toUpperCase().includes(String(item.text).toUpperCase())); msg = ok ? '' : item.msg || `${item.cell} should use ${item.text}.`; break; }
        case 'noMatch': { // no cell in the range matches a pattern: personal data removed or masked
          const g = parseRange(item.range); const re = new RegExp(item.regex); ok = true; let bad = null; eachCell(g, (r, c) => { if (bad) return; const ci = cellInfo(si, addr(r, c)); const v = ci ? (ci.v ?? '') : ''; if (re.test(String(v))) bad = addr(r, c); });
          ok = !bad; msg = ok ? '' : item.msg ? `${item.msg} (${bad} still has one.)` : `${bad} still matches what should have been removed.`; break; }
        case 'condFmt': { // conditional formatting on item.range: kind = dataBar | colorScale | iconSet | highlight (any rule with a fill/font) | expression
          const g = parseRange(item.range); const hit = r => String(r.sqref).split(/\s+/).some(p => { const q = parseRange(p.replace(/\$/g, '')); return q && !(q.r2 < g.r1 || q.r1 > g.r2 || q.c2 < g.c1 || q.c1 > g.c2); }); const rules = (x.sheets[si].cf || []).filter(hit);
          const VIS = ['dataBar', 'colorScale', 'iconSet']; const want = item.kind; const m = rules.filter(r => want === 'highlight' ? !VIS.includes(r.type) : want === 'expression' ? r.type === 'expression' : r.type === want);
          ok = m.length > 0; const NM = { dataBar: 'Data Bars', colorScale: 'a Colour Scale', iconSet: 'an Icon Set', highlight: 'a highlight rule', expression: 'a formula rule (New Rule › Use a formula)' };
          msg = ok ? '' : item.msg || `Add ${NM[want]} to ${item.range}: Home › Conditional Formatting.`;
          if (ok && item.uses) { const fs = m.flatMap(r => r.formulas).join(' ').toUpperCase().replace(/\s/g, ''); ok = item.uses.every(u => fs.includes(String(u).toUpperCase())); if (!ok) msg = item.usesMsg || `The rule’s formula should use ${item.uses.join(' and ')}.`; }
          if (ok && item.marks && want !== 'dataBar' && want !== 'colorScale' && want !== 'iconSet') { // the cells it lights up, worked out by recalculating the rule's own formula
            const w = toWorkbook(x); const exp = m.find(r => r.type === 'expression'); if (exp) { const q0 = parseRange(String(exp.sqref).split(/\s+/)[0].replace(/\$/g, '')); let bad = null;
              for (let r = g.r1; r <= g.r2 && !bad; r++) for (let c = g.c1; c <= g.c2 && !bad; c++) { const tv = v => v === true || (typeof v === 'number' && v !== 0); const got = tv(w.evaluate(shiftFormula('=' + exp.formulas[0], r - q0.r1, c - q0.c1), si)), wantV = tv(w.evaluate(shiftFormula(item.marks, r - g.r1, c - g.c1), si)); if (got !== wantV) bad = addr(r, c) + (wantV ? ' should be highlighted, but isn’t' : ' is highlighted, but shouldn’t be'); }
              if (bad) { ok = false; msg = item.marksMsg || `${bad}. Check the rule’s formula and its $ signs: it is written for the first cell of the range.`; } } }
          break; }
        case 'sparkline': { const g = parseRange(item.range); const all = (x.sheets[si].sparklines || []); const inR = q => { const a = parseRange(String(q.sqref).replace(/\$/g, '').replace(/^.*!/, '')); return a && a.r1 >= g.r1 && a.r2 <= g.r2 && a.c1 >= g.c1 && a.c2 <= g.c2; };
          const gp = all.find(q => q.items.some(inR)); const n = gp ? gp.items.filter(inR).length : 0; const need = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1); ok = !!gp && n >= need; msg = ok ? '' : item.msg || `Put sparklines in ${item.range}: select the numbers, then Insert › Sparklines.`;
          if (ok && item.kind && !item.kind.split('|').includes(gp.type)) { ok = false; msg = `Make them ${item.kind.split('|').join(' or ')} sparklines (Sparkline tab › Type).`; }
          if (ok && item.data) { const d = parseRange(item.data); const f0 = parseRange(gp.items.find(inR).f.replace(/\$/g, '').replace(/^.*!/, '')); if (!f0 || f0.r1 < d.r1 || f0.r2 > d.r2 || f0.c1 < d.c1 || f0.c2 > d.c2) { ok = false; msg = `The sparklines should draw ${item.data}.`; } }
          for (const [k, label] of [['high', 'High Point'], ['low', 'Low Point'], ['negative', 'Negative Points'], ['markers', 'Markers']]) if (ok && item[k] && !gp[k]) { ok = false; msg = `Tick ${label} on the Sparkline tab.`; }
          break; }
        case 'slicer': { ok = (x.slicers || 0) >= (item.min || 1); msg = ok ? '' : item.msg || 'Add a slicer: select the PivotTable, then PivotTable Analyze › Insert Slicer.';
          if (ok && item.field && !(x.slicerLinks || []).some(l => l.field.toLowerCase() === item.field.toLowerCase())) { ok = false; msg = `Add a slicer for ${item.field}.`; }
          if (ok && item.connects && !(x.slicerLinks || []).some(l => (!item.field || l.field.toLowerCase() === item.field.toLowerCase()) && l.pivots >= item.connects)) { ok = false; msg = item.connectsMsg || `Connect the ${item.field || ''} slicer to ${item.connects} PivotTables: right-click the slicer › Report Connections.`; }
          break; }
        case 'timeline': ok = (x.timelines || 0) >= (item.min || 1); msg = ok ? '' : item.msg || 'Add a timeline: select the PivotTable, then PivotTable Analyze › Insert Timeline.'; break;
        case 'dataTable': { const t = item.sheet ? (x.sheets[si] || {}).dataTables || [] : x.sheets.flatMap(q => q.dataTables || []); ok = t.filter(d => !item.twoWay || d.twoWay).length >= (item.min || 1); msg = ok ? '' : item.msg || (t.length ? 'Make it a two-variable data table: give both a Row input cell and a Column input cell.' : 'There’s no data table yet: select the grid, then Data › What-If Analysis › Data Table.'); break; }
        case 'scenarios': { const n = x.sheets.reduce((k, q) => k + (q.scenarios || []).length, 0); ok = n >= (item.min || 1); msg = ok ? '' : item.msg || `Add ${item.min || 1} scenarios with Data › What-If Analysis › Scenario Manager (you have ${n}).`; break; }
        case 'anyFormulaUses': { const t = String(item.text).toUpperCase().split('|'); ok = x.sheets.some(q => Object.values(q.cells).some(c => c.f && t.some(w => String(c.f).toUpperCase().includes(w)))); msg = ok ? '' : item.msg || `Use ${String(item.text).split('|')[0]} somewhere in the workbook.`; break; }
        case 'chart': { // charts the learner made in Excel: type, data, title, labels, legend, axis, highlight, alt text
          const pool = (item.sheet ? [x.sheets[si]] : x.sheets).flatMap(sh => (sh && sh.charts) || []); const norm = r => String(r || '').replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '').toUpperCase();
          const EXK = { histogram: k => k.includes('clusteredColumn') && !k.includes('paretoLine'), pareto: k => k.includes('paretoLine'), box: k => k.includes('boxWhisker'), map: k => k.includes('regionMap') };
          const typeOK = c => { if (!item.chartType) return true; return item.chartType.split('|').some(t => t === 'combo' ? !c.ex && c.kinds.length >= 2 : c.ex ? (EXK[t] ? EXK[t](c.kinds) : c.kinds.some(k => k && k.toLowerCase().includes(t.toLowerCase()))) : c.kinds.includes(t)); };
          const fails = c => { const f = []; const low = String(c.title || (c.titleRef ? 'linked title' : '')).toLowerCase();
            if (!typeOK(c)) f.push(item.typeMsg || `Make it a ${item.chartType.split('|').join(' or ')} chart (Chart Design › Change Chart Type).`);
            if (item.val && !c.series.some(q => norm(q.val) === norm(item.val) || (q.refs || []).some(r => norm(r) === norm(item.val)))) f.push(item.dataMsg || `Chart the values in ${item.val} (Chart Design › Select Data).`);
            if (item.cat && !c.series.some(q => norm(q.cat) === norm(item.cat) || (q.refs || []).some(r => norm(r) === norm(item.cat)))) f.push(item.dataMsg || `Use ${item.cat} for the categories (Chart Design › Select Data › Horizontal Axis Labels › Edit).`);
            if (item.series && c.series.length !== item.series) f.push(`The chart should have ${item.series} series; it has ${c.series.length}.`);
            if (item.title) { if (!low || c.titleDeleted && !c.title) f.push('Add a chart title that states the finding.'); else if ((item.title.not || ['chart title']).some(n => low.trim() === n.toLowerCase())) f.push(item.title.notMsg || 'Replace “Chart Title” with a sentence that says what the chart shows.'); else if (item.title.has && !item.title.has.some(h => String(h).split('+').every(p => low.includes(p.toLowerCase())))) f.push(item.title.msg || `Make the title state the finding: mention ${item.title.has[0].split('+').join(' and ')}.`); else if (item.title.min && low.split(/\s+/).length < item.title.min) f.push(item.title.msg || 'Make the title a short sentence that states the finding.'); }
            if (item.pivot && !c.pivot) f.push(item.pivotMsg || 'Make it a PivotChart: click inside the PivotTable, then PivotTable Analyze › PivotChart.');
            if (item.titleCell && norm(c.titleRef) !== norm(item.titleCell)) f.push(item.titleCellMsg || `Link the chart title to ${item.titleCell}: click the title, type = in the formula bar, click ${item.titleCell}, then press Enter.`);
            if (item.labels !== undefined && !!c.labels !== !!item.labels) f.push(item.labels ? 'Add data labels (Chart Elements + › Data Labels).' : 'Remove the data labels on this chart.');
            if (item.legend === 'none' && c.legend) f.push('Remove the legend (Chart Elements + › untick Legend).'); if (item.legend === 'shown' && !c.legend) f.push('Show the legend, so readers can tell the series apart.');
            if (item.grid === false && c.valAxes.some(v => v.grid)) f.push('Remove the gridlines (Chart Elements + › untick Gridlines).');
            if (item.yZero && !(c.valAxes[0] && (c.valAxes[0].min === 0 || (c.valAxes[0].min === null && ['column', 'bar', 'area'].some(k => c.kinds.includes(k)))))) f.push('Start the vertical axis at 0 (Format Axis › Bounds › Minimum: 0).');
            if (item.highlight && !c.points) f.push('Highlight the bar that matters: click it twice, then Format Data Point › Fill, and make the rest grey.');
            if (item.secondary && !c.groups.some(g => g.secondary)) f.push('Put the second measure on a secondary axis (Change Chart Type › Combo › Secondary Axis).');
            if (item.stacked && !c.groups.some(g => g.grouping === (item.stacked === 'percent' ? 'percentStacked' : 'stacked'))) f.push(item.stacked === 'percent' ? 'Use a 100% stacked chart.' : 'Use a stacked chart.');
            if (item.reverse && !(c.catAxes || []).some(a => a.reverse)) f.push('Put the first category at the top: Format Axis › Categories in reverse order.');
            if (item.totals !== undefined) { const tt = (c.totals || []).slice().sort((x, y) => x - y); if (Array.isArray(item.totals) ? tt.join() !== item.totals.slice().sort((x, y) => x - y).join() : tt.length < item.totals) f.push(item.totalsMsg || 'Set the total bars as totals: click the bar twice, right-click › Format Data Point › tick Set as total.'); }
            if (item.binWidth !== undefined && !(c.binning && c.binning.width === item.binWidth)) f.push(item.binMsg || `Set the bin width to ${item.binWidth} (Format Axis › Bins › Bin width).`);
            if (item.underflow !== undefined && !(c.binning && c.binning.underflow === item.underflow)) f.push(item.binMsg || `Set the underflow bin to ${item.underflow} (Format Axis › Underflow bin).`);
            if (item.overflow !== undefined && !(c.binning && c.binning.overflow === item.overflow)) f.push(item.binMsg || `Set the overflow bin to ${item.overflow} (Format Axis › Overflow bin).`);
            if (item.alt && (c.alt || '').split(/\s+/).filter(Boolean).length < (item.alt === true ? 5 : item.alt)) f.push('Add alt text that says what the chart shows: right-click the chart › View Alt Text.');
            return f; };
          if (!pool.length) { ok = false; msg = item.msg || `There’s no chart${item.sheet ? ' on ' + item.sheet : ''} yet. Select the data, then Insert › Charts.`; break; }
          const scored = pool.map(c => ({ c, f: fails(c) })).sort((a, b) => a.f.length - b.f.length); const need = item.count || 1;
          const passing = scored.filter(q => !q.f.length).length; ok = passing >= need; msg = ok ? '' : need > 1 && passing ? `${passing} of the ${need} charts are right. ${scored.find(q => q.f.length).f[0]}` : scored[0].f[0]; break; }
        case 'chartCount': { const n = (item.sheet ? [x.sheets[si]] : x.sheets).reduce((a, sh) => a + ((sh && sh.charts) || []).length, 0); ok = n >= (item.min || 1); msg = ok ? '' : item.msg || `Make ${item.min || 1} chart${(item.min || 1) > 1 ? 's' : ''}; the file has ${n}.`; break; }
        case 'allowedValues': { const g = parseRange(item.range); const norm = v => item.exact ? String(v) : String(v).trim().toLowerCase(); const set = new Set(item.values.map(norm)); ok = true; eachCell(g, (r, c) => { if (!ok) return; const v = wb.value(si, r, c); if (v === null || v === '') return; if (!set.has(norm(v))) { ok = false; msg = `${addr(r, c)} says “${v}”. ${item.msg || 'It should be one of: ' + item.values.join(', ') + '.'}`; } }); break; }
        case 'lessThan': { const p = parseAddr(item.cell), q = parseAddr(item.than); const v = wb.value(si, p.r, p.c), w = wb.value(si, q.r, q.c); ok = typeof v === 'number' && typeof w === 'number' && v < w && (item.min === undefined || v >= item.min); msg = ok ? '' : item.msg || `${item.cell} should be a number below ${item.than}.`; break; }
        case 'between': { const p = parseAddr(item.cell); const v = wb.value(si, p.r, p.c); ok = typeof v === 'number' && v >= item.min && v <= item.max; msg = ok ? '' : item.msg || `${item.cell} shows ${v ?? 'nothing'}, outside the range that fits this data (${item.min.toLocaleString('en-IN')} to ${item.max.toLocaleString('en-IN')}).`; break; }
        case 'noErrors': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { if (!ok) return; const v = wb.value(si, r, c); if (isErr(v)) { ok = false; msg = `${addr(r, c)} shows ${v.code}. ${item.msg || 'A finished report shouldn’t show error values.'}`; } }); break; }
        default: ok = false; msg = 'Unknown check ' + item.type;
      }
    } catch (e) { ok = false; msg = 'This check couldn’t run: ' + e.message; }
    results.push({ label: item.label, ok, msg: ok ? (item.okMsg || '') : (item.msg && !msg ? item.msg : msg), points: item.points || 1 });
  }
  const total = results.reduce((a, r) => a + r.points, 0), score = results.reduce((a, r) => a + (r.ok ? r.points : 0), 0);
  return { score, total, pct: total ? Math.round(score / total * 100) : 0, results };
}
