// Office Scripts in Sheetah: a script is TypeScript with `function main(workbook: ExcelScript.Workbook, …)`.
// We strip the types (keeping every line where it was, so error lines match), guard the loops, and run it against an
// ExcelScript object model built over Sheetah's workbook engine. It runs in the learner's own browser tab, on a copy
// of the lab's workbook, like their VBA does.
import { Workbook, parseRange, addr, colName, isErr } from './engine.js';

// ---------------------------------------------------------------- TypeScript → JavaScript (types removed in place)
const KW = new Set(['let', 'const', 'var']);
function tokens(src) { const out = []; let i = 0; const n = src.length;
  while (i < n) { const c = src[i];
    if (c === '/' && src[i + 1] === '/') { let j = src.indexOf('\n', i); if (j < 0) j = n; out.push({ t: 'sp', v: src.slice(i, j) }); i = j; continue; }
    if (c === '/' && src[i + 1] === '*') { let j = src.indexOf('*/', i + 2); j = j < 0 ? n : j + 2; out.push({ t: 'sp', v: src.slice(i, j) }); i = j; continue; }
    if (/\s/.test(c)) { let j = i; while (j < n && /\s/.test(src[j])) j++; out.push({ t: 'sp', v: src.slice(i, j) }); i = j; continue; }
    if (c === '"' || c === "'" || c === '`') { let j = i + 1; while (j < n && src[j] !== c) { if (src[j] === '\\') j++; j++; } out.push({ t: 'str', v: src.slice(i, j + 1) }); i = j + 1; continue; }
    if (/[A-Za-z_$]/.test(c)) { let j = i; while (j < n && /[\w$]/.test(src[j])) j++; out.push({ t: 'id', v: src.slice(i, j) }); i = j; continue; }
    if (/[0-9]/.test(c)) { let j = i; while (j < n && /[\w.]/.test(src[j])) j++; out.push({ t: 'num', v: src.slice(i, j) }); i = j; continue; }
    const three = src.slice(i, i + 3), two = src.slice(i, i + 2);
    if (['===', '!==', '...', '**=', '>>>'].includes(three)) { out.push({ t: 'op', v: three }); i += 3; continue; }
    if (['=>', '==', '!=', '<=', '>=', '&&', '||', '++', '--', '+=', '-=', '*=', '/=', '??', '?.'].includes(two)) { out.push({ t: 'op', v: two }); i += 2; continue; }
    out.push({ t: 'op', v: c }); i++; }
  return out; }
const blank = s => s.replace(/[^\n]/g, ' '); // keep newlines so line numbers stay the same
// skip a type expression starting at k (just after ':' or 'as'); returns the index of the first token after it
function skipType(T, k, stops) { let d = 0; let i = k; let seen = false;
  for (; i < T.length; i++) { const x = T[i];
    if (x.t === 'sp') { if (!d && seen && /\n/.test(x.v)) { const nx = T.slice(i + 1).find(y => y.t !== 'sp'); if (nx && !['|', '&', '[', '<', '.'].includes(nx.v)) return i; } continue; }
    if (!d && x.t === 'op' && stops.includes(x.v) && (seen || !['{', '(', '['].includes(x.v))) return i;
    if (x.t === 'op' && ['(', '[', '{', '<'].includes(x.v)) { d++; seen = true; continue; }
    if (x.t === 'op' && [')', ']', '}', '>'].includes(x.v)) { if (!d) return i; d--; continue; }
    seen = true; }
  return i; }
export function stripTypes(src) { const T = tokens(src); const out = T.map(x => x.v); const nxt = i => { let j = i + 1; while (j < T.length && T[j].t === 'sp') j++; return j; }; const prv = i => { let j = i - 1; while (j >= 0 && T[j].t === 'sp') j--; return j; };
  const kill = (a, b) => { for (let j = a; j < b; j++) out[j] = blank(T[j].v); };
  // matching brackets
  const match = new Map(); const st = []; T.forEach((x, i) => { if (x.t === 'op' && '([{'.includes(x.v)) st.push(i); else if (x.t === 'op' && ')]}'.includes(x.v) && st.length) match.set(st.pop(), i); });
  const params = new Set(); // '(' that open a parameter list
  T.forEach((x, i) => { if (!(x.t === 'op' && x.v === '(')) return; const p = prv(i); const close = match.get(i); if (close === undefined) return; const after = nxt(close);
    const isFn = (p >= 0 && T[p].t === 'id' && prv(p) >= 0 && T[prv(p)].v === 'function') || (p >= 0 && T[p].v === 'function');
    const arrow = T[after] && (T[after].v === '=>' || (T[after].v === ':' && (() => { const e = skipType(T, after + 1, ['=>', '{']); return T[e] && T[e].v === '=>'; })()));
    if (isFn || arrow) params.add(i); });
  for (let i = 0; i < T.length; i++) { const x = T[i];
    if (x.t === 'id' && (x.v === 'interface' || (x.v === 'type' && T[nxt(i)] && T[nxt(i)].t === 'id' && T[nxt(nxt(i))] && T[nxt(nxt(i))].v === '='))) { // interface X {…} / type X = …;
      let j = i; if (x.v === 'interface') { while (j < T.length && T[j].v !== '{') j++; j = (match.get(j) ?? j) + 1; } else { j = skipType(T, nxt(nxt(i)) + 1, [';']); if (T[j] && T[j].v === ';') j++; } kill(i, j); i = j - 1; continue; }
    if (x.t === 'id' && x.v === 'as' && T[prv(i)] && (T[prv(i)].t === 'id' || T[prv(i)].t === 'num' || [')', ']', '}'].includes(T[prv(i)].v) || T[prv(i)].t === 'str')) { const e = skipType(T, nxt(i), [';', ',', ')', ']', '}', '=', '.', '?']); kill(i, e); i = e - 1; continue; }
    if (x.t === 'op' && x.v === '!' && T[i + 1] && (T[i + 1].v === '.' || T[i + 1].v === ')' || T[i + 1].v === ';' || T[i + 1].v === ',') && T[i - 1] && (T[i - 1].t === 'id' || T[i - 1].v === ')' || T[i - 1].v === ']')) { out[i] = ' '; continue; } // non-null x!.y
    if (x.t === 'op' && x.v === '<' && T[prv(i)] && T[prv(i)].t === 'id' && (T[prv(prv(i))] || {}).v === 'new') { const e = skipType(T, i, ['(']); kill(i, e); i = e - 1; continue; } // new Map<string, number>()
    if (x.t !== 'op' || x.v !== ':') continue;
    const p = prv(i); if (p < 0) continue;
    // let x: T = …   /   for (const r: T of …)
    if (T[p].t === 'id' && T[prv(p)] && KW.has(T[prv(p)].v)) { const e = skipType(T, i + 1, ['=', ';', ',', ')']); const ee = T[e] && T[e].t === 'id' && (T[e].v === 'of' || T[e].v === 'in') ? e : e; kill(i, ee); i = ee - 1; continue; }
    // a parameter: inside a parameter list, name: T (also name?: T)
    const owner = (() => { let d = 0; for (let j = i - 1; j >= 0; j--) { const y = T[j]; if (y.t !== 'op') continue; if (')]}'.includes(y.v)) d++; else if ('([{'.includes(y.v)) { if (!d) return j; d--; } } return -1; })();
    if (owner >= 0 && params.has(owner) && (T[p].t === 'id' || (T[p].v === '?' && T[prv(p)].t === 'id'))) { if (T[p].v === '?') out[p] = ' '; const e = skipType(T, i + 1, [',', ')', '=']); kill(i, e); i = e - 1; continue; }
    // return type: ): T {   or   ): T =>
    if (T[p].v === ')') { const open = [...match.entries()].find(([a, b]) => b === p); if (open && params.has(open[0])) { const e = skipType(T, i + 1, ['{', '=>']); kill(i, e); i = e - 1; continue; } }
  }
  return out.join(''); }
// a step counter at the top of every loop body, so an endless loop stops instead of freezing the tab
function guardLoops(js) { const T = tokens(js); const out = T.map(x => x.v); const match = new Map(); const st = []; T.forEach((x, i) => { if (x.t === 'op' && '([{'.includes(x.v)) st.push(i); else if (x.t === 'op' && ')]}'.includes(x.v) && st.length) match.set(st.pop(), i); });
  const nxt = i => { let j = i + 1; while (j < T.length && T[j].t === 'sp') j++; return j; };
  T.forEach((x, i) => { if (x.t !== 'id') return; if (x.v === 'for' || x.v === 'while') { const o = nxt(i); if (!T[o] || T[o].v !== '(') return; const c = match.get(o); if (c === undefined) return; const b = nxt(c); if (T[b] && T[b].v === '{') out[b] = '{__g();'; }
    if (x.v === 'do') { const b = nxt(i); if (T[b] && T[b].v === '{') out[b] = '{__g();'; } });
  return out.join(''); }

// ---------------------------------------------------------------- ExcelScript over Sheetah's workbook
const ENUMS = { ClearApplyTo: { all: 'All', formats: 'Formats', contents: 'Contents', hyperlinks: 'Hyperlinks', removeHyperlinks: 'RemoveHyperlinks' }, HorizontalAlignment: { general: 'General', left: 'Left', center: 'Center', right: 'Right', fill: 'Fill', justify: 'Justify' }, VerticalAlignment: { top: 'Top', center: 'Center', bottom: 'Bottom' },
  DeleteShiftDirection: { up: 'Up', left: 'Left' }, InsertShiftDirection: { down: 'Down', right: 'Right' }, BorderIndex: { edgeTop: 'EdgeTop', edgeBottom: 'EdgeBottom', edgeLeft: 'EdgeLeft', edgeRight: 'EdgeRight', insideVertical: 'InsideVertical', insideHorizontal: 'InsideHorizontal' }, BorderLineStyle: { none: 'None', continuous: 'Continuous', dash: 'Dash', double: 'Double' }, BorderWeight: { thin: 'Thin', medium: 'Medium', thick: 'Thick' },
  RangeCopyType: { all: 'All', formulas: 'Formulas', values: 'Values', formats: 'Formats' }, SortOrientation: { rows: 'Rows', columns: 'Columns' }, SheetVisibility: { visible: 'Visible', hidden: 'Hidden', veryHidden: 'VeryHidden' } };
const NAMED = { red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', white: '#FFFFFF', black: '#000000', orange: '#FFA500', gray: '#808080', grey: '#808080', lightgreen: '#90EE90', lightblue: '#ADD8E6' };
const hex = c => { const s = String(c || '').trim(); if (NAMED[s.toLowerCase()]) return NAMED[s.toLowerCase()]; const m = /^#?([0-9a-f]{6})$/i.exec(s); if (!m) throw new OsError(`Invalid color: ${s}`); return '#' + m[1].toUpperCase(); };
export class OsError extends Error { }
const fail = m => { throw new OsError(m); };
function toOut(v) { if (v === null || v === undefined) return ''; if (isErr(v)) return v.code; if (typeof v === 'object' && v.code) return v.code; return v; }
function put(wb, si, r, c, v) { if (v === null || v === undefined || v === '') { wb.clear(si, r, c, true); return; } if (typeof v === 'number' || typeof v === 'boolean') { wb.setRaw(wb.sheets[si], r, c, v); return; } wb.setInput(si, r, c, String(v)); }

export function osWorkbook(wb, ctx = {}) { const sheetObjs = new Map(); const log = ctx.log || [];
  const sheet = s => { if (!sheetObjs.has(s)) sheetObjs.set(s, makeSheet(s)); return sheetObjs.get(s); };
  const siOf = s => { const i = wb.sheets.indexOf(s); if (i < 0) fail('This worksheet has been deleted.'); return i; };
  const parse = (s, a) => { let g; try { g = parseRange(String(a).replace(/\$/g, '').replace(/^.*!/, '')); } catch (e) { g = null; } if (!g) fail(`The argument is invalid or missing or has an incorrect format: "${a}" isn't a range address.`); return g; };
  const ctxRef = { active: wb.sheets[ctx.active || 0] };
  function makeRange(s, g) { const si = () => siOf(s); const R = (r1, c1, r2, c2) => makeRange(s, { r1, c1, r2, c2 }); const rows = () => g.r2 - g.r1 + 1, cols = () => g.c2 - g.c1 + 1;
    const each = f => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) f(r, c); };
    const style = st => { const i = si(); each((r, c) => wb.setStyle(i, r, c, st)); };
    const first = () => wb.style(si(), g.r1, g.c1);
    const grid = f => { const out = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(f(r, c)); out.push(row); } return out; };
    const setGrid = (vals, f, what) => { if (!Array.isArray(vals) || vals.length !== rows() || vals.some(rw => !Array.isArray(rw) || rw.length !== cols())) fail(`The number of rows or columns in the input array doesn't match the size or dimensions of the range. (The range is ${rows()} × ${cols()}; the ${what} are ${Array.isArray(vals) ? vals.length : '?'} × ${Array.isArray(vals) && Array.isArray(vals[0]) ? vals[0].length : '?'}.)`); vals.forEach((rw, i) => rw.forEach((v, j) => f(g.r1 + i, g.c1 + j, v))); wb.invalidate(); };
    const fmt = { getFont: () => ({ setBold: b => style({ bold: !!b || null }), getBold: () => !!first().bold, setItalic: b => style({ italic: !!b || null }), getItalic: () => !!first().italic, setColor: c => style({ color: hex(c) === '#000000' ? null : hex(c) }), getColor: () => first().color || '#000000', setSize: n => style({ size: +n }), getSize: () => first().size || 11, setName: n => style({ font: String(n) }), setUnderline: () => { } }),
      getFill: () => ({ setColor: c => style({ fill: hex(c) }), getColor: () => first().fill || '#FFFFFF', clear: () => style({ fill: null }) }),
      autofitColumns: () => { const i = si(); const sh = wb.sheets[i]; for (let c = g.c1; c <= g.c2; c++) { let w = 4; const u = wb.usedRange(sh); if (u) for (let r = u.r1; r <= u.r2; r++) { const t = String(wb.display(i, r, c) ?? ''); w = Math.max(w, Math.ceil(t.length * (wb.style(i, r, c).bold ? 1.1 : 1)) + 1); } sh.widths[c] = w > 40 ? Math.min(420, w * 7 + 5) : w; } },
      autofitRows: () => { }, setColumnWidth: px => { const sh = wb.sheets[si()]; const ch = Math.round(px / 7); for (let c = g.c1; c <= g.c2; c++) sh.widths[c] = ch > 40 ? Math.round(px) : ch; }, setRowHeight: () => { },
      setHorizontalAlignment: a => style({ align: String(a).toLowerCase() === 'general' ? null : String(a).toLowerCase() }), setVerticalAlignment: () => { }, setWrapText: b => style({ wrap: !!b || null }), setIndentLevel: () => { },
      getRangeBorder: () => ({ setStyle: x => { if (String(x) !== 'None') style({ border: 'thin' }); }, setWeight: () => { }, setColor: () => { } }) };
    const self = {
      getAddress: () => `${s.name}!${addr(g.r1, g.c1)}${rows() > 1 || cols() > 1 ? ':' + addr(g.r2, g.c2) : ''}`,
      getRowCount: () => rows(), getColumnCount: () => cols(), getRowIndex: () => g.r1, getColumnIndex: () => g.c1, getCellCount: () => rows() * cols(),
      getValue: () => toOut(wb.value(si(), g.r1, g.c1)), getValues: () => grid((r, c) => toOut(wb.value(si(), r, c))),
      setValue: v => { const i = si(); if (Array.isArray(v)) fail('setValue takes one value; use setValues for an array.'); each((r, c) => put(wb, i, r, c, v)); wb.invalidate(); },
      setValues: vals => { const i = si(); setGrid(vals, (r, c, v) => put(wb, i, r, c, v), 'values'); },
      getFormula: () => { const f = wb.input(si(), g.r1, g.c1); return f ?? ''; }, getFormulas: () => grid((r, c) => wb.input(si(), r, c) ?? ''),
      setFormula: f => { const i = si(); const t = String(f); each((r, c) => wb.setInput(i, r, c, t)); wb.invalidate(); },
      setFormulas: vals => { const i = si(); setGrid(vals, (r, c, v) => wb.setInput(i, r, c, String(v)), 'formulas'); },
      getText: () => String(wb.display(si(), g.r1, g.c1) ?? ''), getTexts: () => grid((r, c) => String(wb.display(si(), r, c) ?? '')),
      getNumberFormat: () => first().fmt || 'General', setNumberFormat: f => style({ fmt: String(f) === 'General' ? null : String(f) }), setNumberFormatLocal: f => style({ fmt: String(f) }),
      setNumberFormats: vals => { const i = si(); setGrid(vals, (r, c, v) => wb.setStyle(i, r, c, { fmt: v }), 'formats'); },
      getFormat: () => fmt, getWorksheet: () => sheet(s),
      getCell: (r, c) => R(g.r1 + r, g.c1 + c, g.r1 + r, g.c1 + c), getRow: r => R(g.r1 + r, g.c1, g.r1 + r, g.c2), getColumn: c => R(g.r1, g.c1 + c, g.r2, g.c1 + c),
      getOffsetRange: (dr, dc) => R(g.r1 + dr, g.c1 + dc, g.r2 + dr, g.c2 + dc), getResizedRange: (dr, dc) => R(g.r1, g.c1, g.r2 + dr, g.c2 + dc),
      getLastRow: () => R(g.r2, g.c1, g.r2, g.c2), getLastColumn: () => R(g.r1, g.c2, g.r2, g.c2), getLastCell: () => R(g.r2, g.c2, g.r2, g.c2),
      getEntireRow: () => R(g.r1, 0, g.r2, 16383), getEntireColumn: () => R(0, g.c1, 1048575, g.c2),
      getRowsAbove: (n = 1) => R(g.r1 - n, g.c1, g.r1 - 1, g.c2), getRowsBelow: (n = 1) => R(g.r2 + 1, g.c1, g.r2 + n, g.c2), getColumnsAfter: (n = 1) => R(g.r1, g.c2 + 1, g.r2, g.c2 + n), getColumnsBefore: (n = 1) => R(g.r1, g.c1 - n, g.r2, g.c1 - 1),
      getUsedRange: () => { const u = wb.usedRange(s); if (!u) return undefined; const r1 = Math.max(u.r1, g.r1), c1 = Math.max(u.c1, g.c1), r2 = Math.min(u.r2, g.r2), c2 = Math.min(u.c2, g.c2); return r1 > r2 || c1 > c2 ? undefined : R(r1, c1, r2, c2); },
      getSurroundingRegion: () => { const i = si(); const filled = (r, c) => { const v = wb.value(i, r, c); return v !== null && v !== undefined && v !== ''; }; let { r1, c1, r2, c2 } = g; let grow = true; while (grow) { grow = false; if (r1 > 0) for (let c = c1 - (c1 > 0 ? 1 : 0); c <= c2 + 1; c++) if (filled(r1 - 1, c)) { r1--; grow = true; break; } for (let c = c1 - (c1 > 0 ? 1 : 0); c <= c2 + 1; c++) if (filled(r2 + 1, c)) { r2++; grow = true; break; } if (c1 > 0) for (let r = r1; r <= r2; r++) if (filled(r, c1 - 1)) { c1--; grow = true; break; } for (let r = r1; r <= r2; r++) if (filled(r, c2 + 1)) { c2++; grow = true; break; } } return R(r1, c1, r2, c2); },
      clear: (what = 'All') => { const i = si(); const w = String(what); each((r, c) => { if (w === 'Formats') wb.setStyle(i, r, c, { bold: null, italic: null, fill: null, color: null, fmt: null, align: null, wrap: null, size: null }); else wb.clear(i, r, c, w === 'Contents'); }); wb.invalidate(); },
      delete: dir => { const i = si(); const sh = wb.sheets[i]; const up = String(dir || 'Up') === 'Up'; const u = wb.usedRange(sh); if (!u) return; if (up) { for (let c = g.c1; c <= g.c2; c++) { for (let r = g.r1; r <= u.r2; r++) { const src = r + rows(); copyCell(wb, i, src, c, i, r, c); } } } else { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= u.c2; c++) copyCell(wb, i, r, c + cols(), i, r, c); } wb.invalidate(); },
      copyFrom: (src, type = 'All') => { const sg = src._g, ss = src._s; const si2 = siOf(ss); const i = si(); const vals = []; for (let r = 0; r <= sg.r2 - sg.r1; r++) for (let c = 0; c <= sg.c2 - sg.c1; c++) vals.push([r, c, wb.value(si2, sg.r1 + r, sg.c1 + c), wb.input(si2, sg.r1 + r, sg.c1 + c), { ...wb.style(si2, sg.r1 + r, sg.c1 + c) }]);
        for (const [r, c, v, f, st] of vals) { const t = String(type); if (t !== 'Formats') { if (t === 'Values') put(wb, i, g.r1 + r, g.c1 + c, toOut(v) === '' ? null : v); else if (f !== null && f !== undefined && String(f).startsWith('=')) wb.setInput(i, g.r1 + r, g.c1 + c, f); else put(wb, i, g.r1 + r, g.c1 + c, v); } if (t === 'All' || t === 'Formats') wb.setStyle(i, g.r1 + r, g.c1 + c, st); } wb.invalidate(); },
      getSort: () => ({ apply: (fields, matchCase, hasHeaders) => { const i = si(); const start = g.r1 + (hasHeaders ? 1 : 0); const rowsArr = []; for (let r = start; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push([wb.value(i, r, c), wb.input(i, r, c), { ...wb.style(i, r, c) }]); rowsArr.push(row); }
        const cmp = (a, b) => { for (const f of fields || []) { const k = f.key; const x = a[k][0], y = b[k][0]; let d = 0; if (typeof x === 'number' && typeof y === 'number') d = x - y; else d = String(x ?? '').localeCompare(String(y ?? '')); if (d) return f.ascending === false ? -d : d; } return 0; }; rowsArr.sort(cmp);
        rowsArr.forEach((row, k) => row.forEach(([v, f, st], j) => { const r = start + k, c = g.c1 + j; if (f && String(f).startsWith('=')) wb.setInput(i, r, c, f); else put(wb, i, r, c, v); wb.setStyle(i, r, c, st); })); wb.invalidate(); } }),
      select: () => { ctxRef.active = s; ctxRef.sel = { s, g }; },
      _g: g, _s: s };
    return self; }
  function makeSheet(s) { return {
      getName: () => s.name, setName: n => { const t = String(n); if (!t || /[\\/?*[\]:]/.test(t) || t.length > 31) fail(`The name "${t}" isn't valid for a worksheet.`); if (wb.sheets.some(x => x !== s && x.name.toLowerCase() === t.toLowerCase())) fail(`A worksheet with the name "${t}" already exists.`); s.name = t; },
      getRange: a => makeRange(s, a === undefined ? { r1: 0, c1: 0, r2: 1048575, c2: 16383 } : parse(s, a)), getCell: (r, c) => makeRange(s, { r1: r, c1: c, r2: r, c2: c }), getRangeByIndexes: (r, c, nr, nc) => makeRange(s, { r1: r, c1: c, r2: r + nr - 1, c2: c + nc - 1 }),
      getUsedRange: () => { const u = wb.usedRange(s); return u ? makeRange(s, u) : undefined; },
      activate: () => { ctxRef.active = s; }, delete: () => { if (wb.sheets.length === 1) fail('This operation is not permitted: a workbook must contain at least one visible worksheet.'); const i = siOf(s); wb.sheets.splice(i, 1); wb.invalidate(); if (ctxRef.active === s) ctxRef.active = wb.sheets[0]; },
      getPosition: () => siOf(s), setPosition: p => { const i = siOf(s); wb.sheets.splice(i, 1); wb.sheets.splice(Math.max(0, Math.min(p, wb.sheets.length)), 0, s); wb.invalidate(); },
      setTabColor: c => { s.tab = hex(c); }, getTabColor: () => s.tab || '', getFreezePanes: () => ({ freezeRows: n => { s.freeze = { rows: n, cols: (s.freeze && s.freeze.cols) || 0 }; }, freezeColumns: n => { s.freeze = { rows: (s.freeze && s.freeze.rows) || 0, cols: n }; }, unfreeze: () => { s.freeze = null; } }),
      setVisibility: v => { s.hiddenSheet = String(v) !== 'Visible'; }, getVisibility: () => s.hiddenSheet ? 'Hidden' : 'Visible',
      getTables: () => (s.tables || []).map(t => table(s, t)), addTable: (a, headers) => { const g = typeof a === 'string' ? parse(s, a) : a._g; const i = siOf(s); const nm = wb.addTable(i, addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2)); return table(s, (s.tables || []).find(x => x.name === nm)); },
      copy: (pos, rel) => { const i = siOf(s); const spec = wb.toSpec().sheets[i]; let k = 2; while (wb.sheets.some(x => x.name.toLowerCase() === `${s.name} (${k})`.toLowerCase())) k++; const w2 = new Workbook({ sheets: [{ ...spec, name: `${s.name} (${k})` }] }); const ns = w2.sheets[0]; wb.sheets.splice(i + 1, 0, ns); wb.invalidate(); return sheet(ns); } }; }
  function table(s, t) { const rng = () => { const g = parseRange(t.ref); return makeRange(s, g); };
    return { getName: () => t.name, setName: n => { t.name = String(n); }, getRange: rng, getHeaderRowRange: () => rng().getRow(0), getRangeBetweenHeaderAndTotal: () => { const g = parseRange(t.ref); return makeRange(s, { ...g, r1: g.r1 + 1, r2: g.r2 - (t.totals ? 1 : 0) }); },
      getRowCount: () => { const g = parseRange(t.ref); return g.r2 - g.r1 - (t.totals ? 1 : 0); }, getWorksheet: () => sheet(s),
      getColumnByName: n => { const g = parseRange(t.ref); const i = siOf(s); for (let c = g.c1; c <= g.c2; c++) if (String(wb.value(i, g.r1, c)).toLowerCase() === String(n).toLowerCase()) return { getName: () => n, getRangeBetweenHeaderAndTotal: () => makeRange(s, { r1: g.r1 + 1, c1: c, r2: g.r2 - (t.totals ? 1 : 0), c2: c }), getRange: () => makeRange(s, { r1: g.r1, c1: c, r2: g.r2, c2: c }) }; return undefined; },
      addRow: (idx, values) => { const g = parseRange(t.ref); const i = siOf(s); const r = g.r2 + 1; (values || []).forEach((v, j) => put(wb, i, r, g.c1 + j, v)); t.ref = addr(g.r1, g.c1) + ':' + addr(r, g.c2); wb.invalidate(); },
      addRows: (idx, rowsV) => { for (const v of rowsV || []) table(s, t).addRow(-1, v); } }; }
  const api = { getActiveWorksheet: () => sheet(ctxRef.active && wb.sheets.includes(ctxRef.active) ? ctxRef.active : wb.sheets[0]), getWorksheet: n => { const s = wb.sheets.find(x => x.name.toLowerCase() === String(n).toLowerCase()); return s ? sheet(s) : undefined; },
    getWorksheets: () => wb.sheets.map(sheet), getFirstWorksheet: () => sheet(wb.sheets[0]), getLastWorksheet: () => sheet(wb.sheets.at(-1)),
    addWorksheet: n => { let name = n; if (name === undefined) { let k = wb.sheets.length + 1; while (wb.sheets.some(x => x.name === 'Sheet' + k)) k++; name = 'Sheet' + k; } name = String(name); if (wb.sheets.some(x => x.name.toLowerCase() === name.toLowerCase())) fail(`A worksheet with the name "${name}" already exists. Please enter a unique name.`); if (/[\\/?*[\]:]/.test(name) || name.length > 31) fail(`The name "${name}" isn't valid for a worksheet.`); const s = wb.addSheet(name); wb.invalidate(); return sheet(s); },
    getTables: () => wb.sheets.flatMap(s => (s.tables || []).map(t => table(s, t))), getTable: n => { for (const s of wb.sheets) for (const t of s.tables || []) if (t.name.toLowerCase() === String(n).toLowerCase()) return table(s, t); return undefined; },
    getSelectedRange: () => { const sel = ctxRef.sel || ctx.sel; if (sel) return makeRange(sel.s || wb.sheets[sel.si || 0], sel.g); return makeRange(ctxRef.active || wb.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 0 }); },
    getActiveCell: () => api.getSelectedRange().getCell(0, 0), getName: () => ctx.name || 'Book1.xlsx',
    _ctx: ctxRef };
  return api; }
function copyCell(wb, si, r, c, ti, tr, tc) { const v = wb.value(si, r, c), f = wb.input(si, r, c), st = { ...wb.style(si, r, c) }; if (f !== null && f !== undefined && String(f).startsWith('=')) wb.setInput(ti, tr, tc, f); else if (v === null || v === undefined) wb.clear(ti, tr, tc, false); else wb.setRaw(wb.sheets[ti], tr, tc, v); wb.setStyle(ti, tr, tc, st); }

// ---------------------------------------------------------------- running a script
const fmtLog = a => typeof a === 'string' ? a : a === undefined ? 'undefined' : (() => { try { return JSON.stringify(a); } catch (e) { return String(a); } })();
function lineOf(e) { const s = String(e && e.stack || ''); const m = /<anonymous>:(\d+):(\d+)/.exec(s) || /> Function:(\d+):(\d+)/.exec(s) || /Function:(\d+):(\d+)/.exec(s); return m ? +m[1] - 2 : null; }
// runScript(code, wb, { args, active, sel, name }) → { out: [lines], error: {message, line} | null, value }
export function runScript(code, wb, o = {}) { const out = []; const res = { out, error: null, value: undefined };
  let js; try { js = guardLoops(stripTypes(String(code))); } catch (e) { res.error = { message: 'The script couldn’t be read: ' + e.message, line: null }; return res; }
  if (!/\bfunction\s+main\s*\(/.test(js)) { res.error = { message: 'No main function. A script starts: function main(workbook: ExcelScript.Workbook) { … }', line: 1 }; return res; }
  const console_ = { log: (...a) => out.push(a.map(fmtLog).join(' ')), info: (...a) => out.push(a.map(fmtLog).join(' ')), warn: (...a) => out.push(a.map(fmtLog).join(' ')), error: (...a) => out.push(a.map(fmtLog).join(' ')) };
  let n = 0; const __g = () => { if (++n > (o.maxSteps || 2e6)) throw new OsError('The script ran too long and was stopped: is there a loop that never ends?'); };
  const W = osWorkbook(wb, { active: o.active || 0, sel: o.sel, name: o.name, log: out });
  let fn; try { fn = new Function('ExcelScript', 'console', '__g', 'window', 'document', 'fetch', 'localStorage', 'globalThis', js + '\n;return main;'); } catch (e) { const m = /(\d+)/.exec(''); res.error = { message: 'Syntax error: ' + e.message, line: lineOf(e) }; return res; }
  try { const main = fn(ENUMS, console_, __g, undefined, undefined, undefined, undefined, undefined); res.value = main(W, ...(o.args || [])); if (res.value && typeof res.value.then === 'function') res.error = { message: 'This lab runs scripts straight through: leave out async and await.', line: null }; }
  catch (e) { res.error = { message: e instanceof OsError ? e.message : (e && e.message) || String(e), line: lineOf(e) }; }
  res.active = W._ctx.active ? wb.sheets.indexOf(W._ctx.active) : 0; return res; }
export { ENUMS as OS_ENUMS };
