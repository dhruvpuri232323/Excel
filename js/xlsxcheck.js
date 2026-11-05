// Reads a workbook the learner made in real Excel (.xlsx) and checks it against a rubric.
// Formulas are recalculated with Sheetah's own engine, so checks can also run on hidden data.
import { Workbook, parseAddr, parseRange, addr, colIndex, colName, shiftFormula, isErr, eachCell, compareValues, isDateFmt, rangeStr } from './engine.js';

const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '₹#,##0;-₹#,##0', 6: '₹#,##0;[Red]-₹#,##0', 7: '₹#,##0.00;-₹#,##0.00', 8: '₹#,##0.00;[Red]-₹#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'dd-mm-yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'dd-mm-yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 44: '_(₹* #,##0.00_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const parseXML = t => new DOMParser().parseFromString(t, 'application/xml');
const kids = (node, name) => node ? [...node.getElementsByTagNameNS('*', name)] : [];
const one = (node, name) => kids(node, name)[0] || null;
const attr = (n, a) => n ? n.getAttribute(a) : null;

export async function readXlsx(file) {
  const name = file.name || 'workbook.xlsx';
  if (/\.csv$/i.test(name)) return { error: 'This is a CSV file. CSV keeps only values, so formats, formulas and sheets are lost. Save as Excel Workbook (.xlsx) and upload again.', ext: 'csv' };
  if (/\.xls$/i.test(name)) return { error: 'This is an old .xls file. In Excel choose File › Save As › Excel Workbook (.xlsx), then upload that.', ext: 'xls' };
  if (/\.numbers$/i.test(name)) return { error: 'This is an Apple Numbers file. In Numbers choose File › Export To › Excel, then upload the .xlsx file.', ext: 'numbers' };
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
    const s = { name: attr(sh, 'name'), state: attr(sh, 'state') || 'visible', cells: {}, merges: kids(x, 'mergeCell').map(m => attr(m, 'ref')), cols: {}, validations: [], tables: [], comments: [], protect: null };
    for (const col of kids(x, 'col')) { const w = +attr(col, 'width'); for (let c = +attr(col, 'min'); c <= +attr(col, 'max') && c < 200; c++) s.cols[c - 1] = w; }
    const pane = one(x, 'pane'); if (pane && /frozen/.test(attr(pane, 'state') || '')) s.freeze = { rows: +(attr(pane, 'ySplit') || 0), cols: +(attr(pane, 'xSplit') || 0) };
    const shared = {};
    for (const c of kids(x, 'c')) {
      const r = attr(c, 'r'); const t = attr(c, 't'); const sIdx = +(attr(c, 's') || 0); const vEl = one(c, 'v'), fEl = one(c, 'f'); let v = vEl ? vEl.textContent : null; let f = fEl ? fEl.textContent : null;
      if (t === 's') v = ss[+v]; else if (t === 'b') v = v === '1'; else if (t === 'inlineStr') v = kids(c, 't').map(n => n.textContent).join(''); else if (t === 'e') v = { err: v }; else if (t === 'str') v = v; else if (v !== null) v = +v;
      if (fEl && attr(fEl, 't') === 'shared') { const si = attr(fEl, 'si'); if (f) shared[si] = { f, at: parseAddr(r) }; else if (shared[si]) { const p = parseAddr(r); f = shiftFormula('=' + shared[si].f, p.r - shared[si].at.r, p.c - shared[si].at.c).slice(1); } }
      const xf = styles.xfs[sIdx] || { fmt: 'General', font: 0, fill: 0, locked: true };
      s.cells[r] = { v, f: f ? '=' + f : null, fmt: xf.fmt, bold: (styles.fonts[xf.font] || {}).bold, italic: (styles.fonts[xf.font] || {}).italic, fill: (styles.fills[xf.fill] || {}).solid, align: xf.align, wrap: xf.wrap, locked: xf.locked };
    }
    for (const dv of kids(x, 'dataValidation')) { const f1 = one(dv, 'formula1'); s.validations.push({ type: attr(dv, 'type'), sqref: attr(dv, 'sqref'), formula: f1 ? f1.textContent : '' }); }
    // newer Excel writes list validations in an extension block (x14:dataValidation)
    const sp = one(x, 'sheetProtection'); if (sp) s.protect = { sheet: attr(sp, 'sheet') === '1' || attr(sp, 'sheet') === 'true' };
    const ps = one(x, 'pageSetup'); s.page = { orientation: attr(ps, 'orientation') || 'portrait', fitW: attr(ps, 'fitToWidth'), fitH: attr(ps, 'fitToHeight'), fitToPage: !!(one(x, 'pageSetUpPr') && attr(one(x, 'pageSetUpPr'), 'fitToPage') === '1'), scale: attr(ps, 'scale') };
    const hf = one(x, 'headerFooter'); s.headerFooter = hf ? hf.textContent : '';
    const af = one(x, 'autoFilter'); s.autoFilter = af ? attr(af, 'ref') : null;
    const tc = one(x, 'tabColor'); s.tabColor = tc ? (attr(tc, 'rgb') || attr(tc, 'theme') || 'set') : null;
    for (const tp of kids(x, 'tablePart')) { const rel = srels[attr(tp, 'r:id') || tp.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')]; if (!rel) continue; const tx = parseXML(await text(rel.target)); const tb = tx.documentElement; s.tables.push({ name: attr(tb, 'displayName') || attr(tb, 'name'), ref: attr(tb, 'ref'), totals: +(attr(tb, 'totalsRowCount') || 0) > 0 || attr(tb, 'totalsRowShown') === '1' && +(attr(tb, 'totalsRowCount') || 0) > 0, columns: kids(tx, 'tableColumn').map(c => attr(c, 'name')) }); }
    for (const rel of Object.values(srels)) if (/comments/.test(rel.type)) { const cx = await text(rel.target); if (cx) for (const cm of kids(parseXML(cx), 'comment')) s.comments.push(attr(cm, 'ref')); }
    for (const rel of Object.values(srels)) if (/threadedComment/.test(rel.type)) { const cx = await text(rel.target); if (cx) for (const cm of kids(parseXML(cx), 'threadedComment')) s.comments.push(attr(cm, 'ref')); }
    const ext = x.getElementsByTagName('x14:dataValidation'); for (const dv of [...ext]) { const f1 = dv.getElementsByTagName('xm:f')[0]; const sq = dv.getElementsByTagName('xm:sqref')[0]; s.validations.push({ type: attr(dv, 'type'), sqref: sq ? sq.textContent : '', formula: f1 ? f1.textContent : '' }); }
    sheets.push(s);
  }
  const props = await text('docProps/core.xml');
  return { name, sheets, names, author: props ? ((/<dc:creator>([^<]*)/.exec(props) || [])[1] || '') : '', hasMacros: !!zip.file('xl/vbaProject.bin') };
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
  const spec = { sheets: x.sheets.map(s => { const cells = {}; const allTables = x.sheets.flatMap(q => q.tables); for (const [a, c] of Object.entries(s.cells)) { if (c.f) cells[a] = { f: resolveStructured(c.f, allTables, a).replace(/_xlfn\./g, '').replace(/_xlws\./g, ''), fmt: c.fmt !== 'General' ? c.fmt : undefined }; else if (c.v !== null && c.v !== undefined) cells[a] = { v: c.v && c.v.err ? null : c.v, fmt: c.fmt !== 'General' ? c.fmt : undefined }; } return { name: s.name, cells }; }), names: {} };
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
        case 'textValues': { ok = Object.entries(item.cells).every(([a, want]) => { const p = parseAddr(a); return near(wb.value(si, p.r, p.c), want); }); msg = ok ? '' : item.msg || 'Some cells don’t hold the expected text yet.'; break; }
        case 'filled': { const p = parseAddr(item.cell); const v = wb.value(si, p.r, p.c); ok = v !== null && v !== ''; msg = ok ? '' : item.msg || `${item.cell} is still empty.`; break; }
        case 'filledRange': { const g = parseRange(item.range); ok = true; eachCell(g, (r, c) => { const v = wb.value(si, r, c); if (ok && (v === null || v === '')) { ok = false; msg = item.msg || `${addr(r, c)} is still blank.`; } }); break; }
        case 'definedName': { const d = Object.entries(x.names).find(([n]) => n.toLowerCase() === item.name.toLowerCase()); ok = !!d && (!item.ref || d[1].ref.replace(/\$/g, '').toUpperCase().endsWith(item.ref.toUpperCase())); msg = ok ? '' : `Create a name called ${item.name}${item.ref ? ' for ' + item.ref : ''} (Formulas › Define Name, or the Name Box).`; break; }
        case 'comment': ok = x.sheets[si].comments.some(a => !item.cell || a === item.cell); msg = ok ? '' : `Add a note or comment${item.cell ? ' to ' + item.cell : ''}.`; break;
        case 'hiddenSheet': ok = x.sheets.some(s => s.state !== 'visible'); msg = ok ? '' : 'Hide the sheet named in the task (right-click its tab › Hide).'; break;
        case 'autoFilter': ok = !!x.sheets[si].autoFilter || x.sheets[si].tables.length > 0; msg = ok ? '' : 'Turn on filter buttons with Data › Filter.'; break;
        case 'variant': { // put hidden data into the inputs, recalculate, compare outputs with the model formulas
          ok = true; for (const dataset of item.datasets) { const w2 = toWorkbook(x); for (const [a, v] of Object.entries(dataset)) { const [sh, ad] = a.includes('!') ? a.split('!') : [item.sheet, a]; const s2 = w2.sheetIndexByName(sh); const p = parseAddr(ad); const o = w2.cellObj(w2.sheets[s2], p.r, p.c, true); o.v = v; o.f = null; } w2.invalidate(); for (const out of item.outputs) { const [sh, ad] = out.cell.includes('!') ? out.cell.split('!') : [item.sheet, out.cell]; const s2 = w2.sheetIndexByName(sh); const p = parseAddr(ad); const o = w2.cellObj(w2.sheets[s2], p.r, p.c); if (o) o.cached = undefined; const got = w2.value(s2, p.r, p.c); const want = w2.evaluate(out.ref, s2); if (!near(got, want, 1e-6)) { ok = false; msg = `${out.label || out.cell} gave ${got} on a hidden set of data, but should give ${want}. Check that it points at the whole input range.`; break; } } if (!ok) break; } break; }
        default: ok = false; msg = 'Unknown check ' + item.type;
      }
    } catch (e) { ok = false; msg = 'This check couldn’t run: ' + e.message; }
    results.push({ label: item.label, ok, msg: ok ? (item.okMsg || '') : (item.msg && !msg ? item.msg : msg), points: item.points || 1 });
  }
  const total = results.reduce((a, r) => a + r.points, 0), score = results.reduce((a, r) => a + (r.ok ? r.points : 0), 0);
  return { score, total, pct: total ? Math.round(score / total * 100) : 0, results };
}
