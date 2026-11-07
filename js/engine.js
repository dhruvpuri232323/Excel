// Sheetah formula engine: a small, dependency-free spreadsheet calculator.
// Handles values, formulas, references ($A$1, ranges, other sheets, names),
// the common functions, Excel-style number formats and AutoFill series.

// ---------- addresses ----------
export function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
export function colIndex(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
export function addr(r, c) { return colName(c) + (r + 1); }
export function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(a).trim());
  if (!m) return null; return { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) };
}
export function parseRange(a) {
  const [p, q] = String(a).split(':'); const s = parseAddr(p); if (!s) return null;
  const e = q ? parseAddr(q) : s; if (!e) return null;
  return { r1: Math.min(s.r, e.r), c1: Math.min(s.c, e.c), r2: Math.max(s.r, e.r), c2: Math.max(s.c, e.c) };
}
export function rangeStr(g) { const a = addr(g.r1, g.c1), b = addr(g.r2, g.c2); return a === b ? a : a + ':' + b; }
export function eachCell(g, fn) { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c); }

// ---------- values ----------
export class XErr { constructor(code) { this.code = code; } toString() { return this.code; } }
export const ERR = { DIV0: new XErr('#DIV/0!'), NAME: new XErr('#NAME?'), VALUE: new XErr('#VALUE!'), REF: new XErr('#REF!'), NA: new XErr('#N/A'), NUM: new XErr('#NUM!'), NULL: new XErr('#NULL!'), SPILL: new XErr('#SPILL!'), CALC: new XErr('#CALC!') };
const ERR_BY_CODE = {}; for (const e of Object.values(ERR)) ERR_BY_CODE[e.code] = e;
export const isErr = v => v instanceof XErr;
const isBlank = v => v === null || v === undefined;

// Excel day serials (1900 system, with the 1900 leap-year quirk kept so serials match Excel)
const EPOCH = Date.UTC(1899, 11, 30);
export function dateToSerial(y, m, d) { let s = (Date.UTC(y, m - 1, d) - EPOCH) / 864e5; return s; }
export function serialToDate(s) { const ms = EPOCH + Math.floor(s) * 864e5; const dt = new Date(ms); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() }; }
export function todaySerial() { const n = new Date(); return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()); }

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const MONTHS_FULL = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
const DAYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
const DAYS_FULL = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

// Turn what a person typed into a typed value, the way Excel does (Indian locale: day first).
// Returns { v, fmt? } or { f } for a formula.
export function parseInput(text) {
  if (text === null || text === undefined) return { v: null };
  let t = String(text);
  if (t === '') return { v: null };
  if (t[0] === "'") return { v: t.slice(1), forcedText: true };
  if (t[0] === '=' && t.length > 1) return { f: t };
  const s = t.trim();
  if (/^(true|false)$/i.test(s)) return { v: s.toUpperCase() === 'TRUE' };
  if (ERR_BY_CODE[s.toUpperCase()]) return { v: ERR_BY_CODE[s.toUpperCase()] };
  // numbers, with ₹, thousands separators (western or Indian), %, sign, exponent
  let m = /^([-+])?(₹)?\s*((\d{1,3}(,\d{2,3})+|\d+)(\.\d*)?|\.\d+)(E[-+]?\d+)?\s*(%)?$/i.exec(s);
  if (m) {
    const num = parseFloat((m[3] + (m[7] || '')).replace(/,/g, '')) * (m[1] === '-' ? -1 : 1);
    if (m[8]) return { v: num / 100, fmt: (m[6] && m[6].length > 1) ? '0.00%' : '0%' };
    if (m[2]) return { v: num, fmt: '₹#,##0.00' };
    if (m[5]) return { v: num, fmt: /\d,\d\d,/.test(m[3]) ? '#,##,##0' : '#,##0' };
    return { v: num };
  }
  // (123) style negatives
  m = /^\((\d+(\.\d+)?)\)$/.exec(s); if (m) return { v: -parseFloat(m[1]) };
  // dates: 12-03-2026, 12/3/26, 2026-03-12, 12-Mar-2026, 12 Mar 2026, Mar-26
  m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/.exec(s);
  if (m) { let [d, mo, y] = [+m[1], +m[2], +m[3]]; if (y < 100) y += y < 30 ? 2000 : 1900; if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return { v: dateToSerial(y, mo, d), fmt: 'dd-mm-yyyy' }; }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return { v: dateToSerial(+m[1], +m[2], +m[3]), fmt: 'yyyy-mm-dd' };
  m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ,]*(\d{2}|\d{4})?$/.exec(s);
  if (m) { const mi = monthIdx(m[2]); if (mi >= 0) { let y = m[3] ? +m[3] : new Date().getFullYear(); if (y < 100) y += 2000; return { v: dateToSerial(y, mi + 1, +m[1]), fmt: 'd-mmm-yy' }; } }
  m = /^([A-Za-z]{3,9})[- ](\d{2}|\d{4})$/.exec(s);
  if (m) { const mi = monthIdx(m[1]); if (mi >= 0) { let y = +m[2]; if (y < 100) y += 2000; return { v: dateToSerial(y, mi + 1, 1), fmt: 'mmm-yy' }; } }
  // times 9:30, 18:45:10, 9:30 PM
  m = /^(\d{1,2}):(\d{2})(:(\d{2}))?\s*(am|pm)?$/i.exec(s);
  if (m) { let h = +m[1]; if (m[5]) { h = h % 12 + (/pm/i.test(m[5]) ? 12 : 0); } const v = (h * 3600 + +m[2] * 60 + (+m[4] || 0)) / 86400; return { v, fmt: m[5] ? 'h:mm AM/PM' : 'hh:mm' }; }
  return { v: t };
}
function monthIdx(w) { w = w.toLowerCase(); let i = MONTHS.indexOf(w.slice(0, 3)); if (i >= 0 && (w.length === 3 || MONTHS_FULL[i].startsWith(w))) return i; return -1; }

// ---------- tokenizer ----------
const OPS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', ':'];
export function tokenize(f) {
  const src = f[0] === '=' ? f.slice(1) : f; const off = f[0] === '=' ? 1 : 0;
  const toks = []; let i = 0;
  const push = (type, val, s, e, extra) => toks.push(Object.assign({ type, val, s: s + off, e: e + off }, extra || {}));
  while (i < src.length) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\t') { i++; continue; }
    if (ch === '"') { let j = i + 1, out = ''; while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { out += '"'; j += 2; continue; } break; } out += src[j++]; } if (j >= src.length) throw new SyntaxError('A text value is missing its closing quote'); push('str', out, i, j + 1); i = j + 1; continue; }
    if (ch === '#') { const m = /^#(DIV\/0!|NAME\?|VALUE!|REF!|N\/A|NUM!|NULL!|SPILL!|CALC!)/i.exec(src.slice(i)); if (m) { push('err', ERR_BY_CODE[m[0].toUpperCase()], i, i + m[0].length); i += m[0].length; continue; } }
    if (ch === '{') { // an array constant: {1,2,3} is a row, {1;2;3} a column
      const rows = [[]]; let j = i + 1;
      for (;;) {
        while (src[j] === ' ') j++;
        const rest = src.slice(j); let m2;
        if (rest[0] === '"') { let k = j + 1, o = ''; while (k < src.length) { if (src[k] === '"') { if (src[k + 1] === '"') { o += '"'; k += 2; continue; } break; } o += src[k++]; } rows.at(-1).push(o); j = k + 1; }
        else if ((m2 = /^[-+]?(\d+\.?\d*|\.\d+)(E[-+]?\d+)?/i.exec(rest))) { rows.at(-1).push(parseFloat(m2[0])); j += m2[0].length; }
        else if ((m2 = /^(TRUE|FALSE)/i.exec(rest))) { rows.at(-1).push(m2[0].toUpperCase() === 'TRUE'); j += m2[0].length; }
        else if ((m2 = /^#(DIV\/0!|NAME\?|VALUE!|REF!|N\/A|NUM!|NULL!)/i.exec(rest))) { rows.at(-1).push(ERR_BY_CODE[m2[0].toUpperCase()]); j += m2[0].length; }
        else throw new SyntaxError('An array constant can hold only numbers, text in quotes, TRUE/FALSE or errors');
        while (src[j] === ' ') j++;
        if (src[j] === ',') { j++; continue; } if (src[j] === ';') { rows.push([]); j++; continue; } if (src[j] === '}') { j++; break; }
        throw new SyntaxError('An array constant is missing its closing }');
      }
      if (rows.some(r => r.length !== rows[0].length)) throw new SyntaxError('Every row of an array constant needs the same number of items');
      push('arr', rows, i, j); i = j; continue;
    }
    if (ch === '@') { push('op', '@', i, i + 1); i++; continue; }
    { const mr = /^(\$?)(\d+):(\$?)(\d+)(?![\d.A-Za-z(])/.exec(src.slice(i)); if (mr && +mr[2] > 0 && +mr[4] > 0) { push('rows', null, i, i + mr[0].length, { sheet: null, sheetText: '', aAbs: !!mr[1], a: +mr[2] - 1, bAbs: !!mr[3], b: +mr[4] - 1 }); i += mr[0].length; continue; } }
    if (/[0-9.]/.test(ch)) { const m = /^(\d+\.?\d*|\.\d+)(E[-+]?\d+)?/i.exec(src.slice(i)); if (m) { push('num', parseFloat(m[0]), i, i + m[0].length); i += m[0].length; continue; } }
    // sheet-qualified or plain reference / name / function
    let m = /^('((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!/.exec(src.slice(i));
    let sheet = null, start = i;
    if (m) { sheet = m[2] ? m[2].replace(/''/g, "'") : m[3]; i += m[0].length; }
    m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(!.])/.exec(src.slice(i));
    if (m && colIndex(m[2]) < 16384 && colIndex(m[4]) < 16384) { push('cols', null, start, i + m[0].length, { sheet, sheetText: sheet !== null ? src.slice(start, i) : '', aAbs: !!m[1], a: colIndex(m[2]), bAbs: !!m[3], b: colIndex(m[4]) }); i += m[0].length; continue; }
    m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(#(?![A-Za-z]))?(?![A-Za-z0-9_(])/.exec(src.slice(i));
    if (m && parseInt(m[4], 10) > 0 && colIndex(m[2]) < 16384) {
      push('ref', null, start, i + m[0].length, { sheet, cAbs: !!m[1], c: colIndex(m[2]), rAbs: !!m[3], r: parseInt(m[4], 10) - 1, sheetText: sheet !== null ? src.slice(start, i) : '', spill: !!m[5] });
      i += m[0].length; continue;
    }
    if (sheet !== null) throw new SyntaxError('Expected a cell after the sheet name');
    m = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(src.slice(i));
    if (m) {
      const word = m[0]; const after = src.slice(i + word.length).trimStart();
      // files saved by Excel write newer functions as _xlfn.XLOOKUP, _xlfn._xlws.FILTER and LET names as _xlpm.x
      if (after[0] === '(') push('fn', word.toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, ''), i, i + word.length);
      else if (/^(TRUE|FALSE)$/i.test(word)) push('bool', word.toUpperCase() === 'TRUE', i, i + word.length);
      else push('name', word.replace(/^_xlpm\./i, ''), i, i + word.length);
      i += word.length; continue;
    }
    if (ch === '(' || ch === ')' || ch === ',' || ch === ';') { push(ch === ';' ? ',' : ch, ch, i, i + 1); i++; continue; }
    const op = OPS.find(o => src.startsWith(o, i));
    if (op) { push('op', op, i, i + op.length); i += op.length; continue; }
    throw new SyntaxError(`Excel doesn't recognise "${ch}" here`);
  }
  return toks;
}

// ---------- parser (Pratt) ----------
const BIN = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
export function parse(f) {
  const toks = tokenize(f); let p = 0;
  const peek = () => toks[p]; const next = () => toks[p++];
  function primary() {
    const t = next(); if (!t) throw new SyntaxError('The formula ends too early');
    if (t.type === 'num') return { k: 'num', v: t.val };
    if (t.type === 'str') return { k: 'str', v: t.val };
    if (t.type === 'bool') return { k: 'bool', v: t.val };
    if (t.type === 'err') return { k: 'err', v: t.val };
    if (t.type === 'arr') return { k: 'arr', v: t.val };
    if (t.type === 'op' && (t.val === '-' || t.val === '+')) { const e = unaryOperand(); return t.val === '-' ? { k: 'neg', e } : e; }
    if (t.type === 'op' && t.val === '@') return { k: 'at', e: unaryOperand() };
    if (t.type === '(') { const e = expr(0); const c = next(); if (!c || c.type !== ')') throw new SyntaxError('A bracket is missing: every ( needs a )'); return postfix(e); }
    if (t.type === 'ref') {
      if (t.spill) return postfix({ k: 'spillref', sheet: t.sheet, r: t.r, c: t.c });
      let node = { k: 'ref', sheet: t.sheet, r: t.r, c: t.c };
      if (peek() && peek().type === 'op' && peek().val === ':' && toks[p + 1] && toks[p + 1].type === 'ref') { next(); const u = next(); node = { k: 'range', sheet: t.sheet, r1: Math.min(t.r, u.r), c1: Math.min(t.c, u.c), r2: Math.max(t.r, u.r), c2: Math.max(t.c, u.c) }; }
      return postfix(node);
    }
    if (t.type === 'name') return postfix({ k: 'name', name: t.val });
    if (t.type === 'cols') return postfix({ k: 'range', whole: 'col', sheet: t.sheet, r1: 0, r2: 0, c1: Math.min(t.a, t.b), c2: Math.max(t.a, t.b) });
    if (t.type === 'rows') return postfix({ k: 'range', whole: 'row', sheet: t.sheet, c1: 0, c2: 0, r1: Math.min(t.a, t.b), r2: Math.max(t.a, t.b) });
    if (t.type === 'fn') {
      const o = next(); if (!o || o.type !== '(') throw new SyntaxError('Expected ( after ' + t.val);
      const args = [];
      if (peek() && peek().type === ')') { next(); return postfix({ k: 'call', fn: t.val, args }); }
      for (;;) {
        if (peek() && (peek().type === ',' || peek().type === ')')) args.push({ k: 'missing' }); else args.push(expr(0));
        const s = next(); if (!s) throw new SyntaxError(`${t.val}( is missing its closing bracket`);
        if (s.type === ')') break; if (s.type !== ',') throw new SyntaxError('Expected a comma or ) in ' + t.val);
      }
      return postfix({ k: 'call', fn: t.val, args });
    }
    throw new SyntaxError(`Unexpected "${t.val ?? t.type}" in the formula`);
  }
  function unaryOperand() { // unary minus binds tighter than ^ in Excel
    const t = peek(); if (t && t.type === 'op' && (t.val === '-' || t.val === '+')) { next(); const e = unaryOperand(); return t.val === '-' ? { k: 'neg', e } : e; }
    if (t && t.type === 'op' && t.val === '@') { next(); return { k: 'at', e: unaryOperand() }; }
    return postfix(primary());
  }
  function postfix(e) { while (peek() && peek().type === 'op' && peek().val === '%') { next(); e = { k: 'pct', e }; } return e; }
  function expr(minP) {
    let left = unaryOperand();
    for (;;) {
      const t = peek(); if (!t || t.type !== 'op' || !(t.val in BIN)) break;
      const pr = BIN[t.val]; if (pr <= minP) break;
      next(); const right = expr(pr); left = { k: 'bin', op: t.val, a: left, b: right };
    }
    return left;
  }
  const ast = expr(0);
  if (p < toks.length) { const t = toks[p]; throw new SyntaxError(t.type === ')' ? 'There is an extra )' : `Unexpected "${t.val ?? t.type}"`); }
  return ast;
}

// Move relative references when a formula is copied by (dr, dc). Returns new text; broken refs become #REF!
export function shiftFormula(f, dr, dc) {
  if (!f || f[0] !== '=') return f;
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  let out = '', last = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'cols' || t.type === 'rows') { const d = t.type === 'cols' ? dc : dr; const a = t.aAbs ? t.a : t.a + d, b = t.bAbs ? t.b : t.b + d; out += f.slice(last, t.s) + (a < 0 || b < 0 ? '#REF!' : t.sheetText + (t.type === 'cols' ? (t.aAbs ? '$' : '') + colName(a) + ':' + (t.bAbs ? '$' : '') + colName(b) : (t.aAbs ? '$' : '') + (a + 1) + ':' + (t.bAbs ? '$' : '') + (b + 1))); last = t.e; continue; }
    if (t.type !== 'ref') continue;
    const r = t.rAbs ? t.r : t.r + dr, c = t.cAbs ? t.c : t.c + dc;
    out += f.slice(last, t.s);
    if (r < 0 || c < 0) { out += '#REF!'; }
    else out += t.sheetText + (t.cAbs ? '$' : '') + colName(c) + (t.rAbs ? '$' : '') + (r + 1) + (t.spill ? '#' : '');
    last = t.e;
  }
  return out + f.slice(last);
}
// Cycle the reference under the caret through A1 -> $A$1 -> A$1 -> $A1 (the F4 key)
export function cycleRefAt(f, caret) {
  let toks; try { toks = tokenize(f); } catch (e) { return null; }
  const t = toks.find(t => t.type === 'ref' && caret >= t.s && caret <= t.e); if (!t) return null;
  const state = (t.cAbs ? 2 : 0) + (t.rAbs ? 1 : 0); // 0 A1, 3 $A$1, 1 A$1, 2 $A1
  const nextState = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
  const ca = nextState >= 2, ra = nextState % 2 === 1;
  const txt = t.sheetText + (ca ? '$' : '') + colName(t.c) + (ra ? '$' : '') + (t.r + 1) + (t.spill ? '#' : '');
  return { text: f.slice(0, t.s) + txt + f.slice(t.e), caret: t.s + txt.length };
}
export function refsIn(f) {
  try { return tokenize(f).filter(t => t.type === 'ref').map((t, i, arr) => t); } catch (e) { return []; }
}

// ---------- workbook ----------
export class Workbook {
  constructor(spec) {
    this.sheets = []; this.names = {}; this.listeners = new Set(); this._cache = null; this.circular = false; this._cur = []; this._let = []; this._resetSpill();
    if (spec) this.load(spec); else this.addSheet('Sheet1');
  }
  addSheet(name, rows = 100, cols = 26) { const s = { name, cells: new Map(), rows, cols, widths: {}, freeze: null, tables: [], validations: [] }; this.sheets.push(s); return s; }
  sheet(nameOrIdx) {
    if (typeof nameOrIdx === 'number') return this.sheets[nameOrIdx];
    const n = String(nameOrIdx).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === n);
  }
  load(spec) {
    this.sheets = []; this.names = {};
    for (const sp of spec.sheets || []) {
      const s = this.addSheet(sp.name || 'Sheet' + (this.sheets.length + 1), sp.rows || 100, sp.cols || 26);
      if (sp.widths) for (const [k, w] of Object.entries(sp.widths)) s.widths[colIndex(k)] = w;
      s.freeze = sp.freeze || null;
      for (const [a, val] of Object.entries(sp.cells || {})) {
        const p = parseAddr(a); if (!p) continue;
        this.setRaw(s, p.r, p.c, val, true);
      }
      for (const [rg, st] of Object.entries(sp.styles || {})) { const g = parseRange(rg); if (g) eachCell(g, (r, c) => { const cell = this.cellObj(s, r, c, true); cell.s = Object.assign({}, cell.s, st); }); }
      if (sp.tables) s.tables = sp.tables.slice();
      if (sp.validations) s.validations = sp.validations.slice();
      if (sp.data) { // array-of-rows shortcut starting at sp.dataAt (default A1)
        const at = parseAddr(sp.dataAt || 'A1');
        sp.data.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined && v !== '') this.setRaw(s, at.r + i, at.c + j, v, true); }));
      }
      if (sp.gen) { // deterministic generated data: {from:'A2', rows:N, cols:[{seq|cycle|rand|date}]}
        const at = parseAddr(sp.gen.from || 'A2'); let seed = sp.gen.seed || 7; const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
        for (let i = 0; i < sp.gen.rows; i++) { if (sp.gen.gapAt && i === sp.gen.gapAt) continue; sp.gen.cols.forEach((g, j) => { let v; if (g.seq !== undefined) v = (g.prefix || '') + (g.seq + i); else if (g.cycle) v = g.cycle[(i * (g.stride || 1)) % g.cycle.length]; else if (g.rand) v = Math.round((g.rand[0] + rnd() * (g.rand[1] - g.rand[0])) / (g.round || 1)) * (g.round || 1); else if (g.date !== undefined) v = g.date + Math.floor(i / (g.per || 1)); else if (g.f) v = g.f.replace(/\{r\}/g, at.r + i + 1); this.setRaw(s, at.r + i, at.c + j, v, true); if (g.fmt) { const cell = this.cellObj(s, at.r + i, at.c + j, true); cell.s = Object.assign({}, cell.s, { fmt: g.fmt }); } }); }
      }
      const used = this.usedRange(s); if (used) { s.rows = Math.max(s.rows, used.r2 + 21); s.cols = Math.max(s.cols, used.c2 + 6); }
    }
    for (const [n, ref] of Object.entries(spec.names || {})) this.names[n.toLowerCase()] = { name: n, ref };
    if (!this.sheets.length) this.addSheet('Sheet1');
    this.invalidate();
  }
  cellObj(s, r, c, create) { const k = r * 16384 + c; let o = s.cells.get(k); if (!o && create) { o = { v: null, f: null, s: null }; s.cells.set(k, o); } return o; }
  // raw value from authoring (number, string, bool, '=formula', or {v,f,fmt,...})
  setRaw(s, r, c, val, silent) {
    const cell = this.cellObj(s, r, c, true);
    if (val && typeof val === 'object' && !(val instanceof XErr)) {
      const { v, f, ...st } = val; cell.v = v ?? null; cell.f = f ?? null; if (Object.keys(st).length) cell.s = Object.assign({}, cell.s, st);
    } else if (typeof val === 'string' && val[0] === '=') { cell.f = val; cell.v = null; }
    else if (typeof val === 'string') { const p = parseInput(val); if (p.f) { cell.f = p.f; cell.v = null; } else { cell.v = p.v; cell.f = null; if (p.fmt) cell.s = Object.assign({}, cell.s, { fmt: p.fmt }); } }
    else { cell.v = val; cell.f = null; }
    if (!silent) this.invalidate();
  }
  // what a person typed
  setInput(sheetIdx, r, c, text) {
    const s = this.sheets[sheetIdx]; const p = parseInput(text);
    const cell = this.cellObj(s, r, c, true);
    if (p.f) { cell.f = p.f; cell.v = null; }
    else { cell.f = null; cell.v = p.v; if (p.fmt && (!cell.s || !cell.s.fmt || cell.s.fmt === 'General')) cell.s = Object.assign({}, cell.s, { fmt: p.fmt }); if (p.forcedText) cell.s = Object.assign({}, cell.s, { text: true }); else if (cell.s && cell.s.text) { delete cell.s.text; } }
    if (cell.v === null && !cell.f && !cell.s) s.cells.delete(r * 16384 + c);
    this.invalidate();
  }
  clear(sheetIdx, r, c, keepStyle = true) { const s = this.sheets[sheetIdx]; const cell = this.cellObj(s, r, c); if (!cell) return; if (keepStyle && cell.s) { cell.v = null; cell.f = null; } else s.cells.delete(r * 16384 + c); this.invalidate(); }
  input(sheetIdx, r, c) { // text shown in the formula bar
    const cell = this.cellObj(this.sheets[sheetIdx], r, c); if (!cell) return '';
    if (cell.f) return cell.f; const v = cell.v; if (isBlank(v)) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) return v.code;
    if (typeof v === 'number') { const fm = cell.s && cell.s.fmt; if (fm && isDateFmt(fm)) return formatValue(v, v % 1 ? 'dd-mm-yyyy hh:mm' : 'dd-mm-yyyy'); if (fm && fm.includes('%')) return +(v * 100).toPrecision(12) + '%'; return String(+v.toPrecision(15)); }
    return (cell.s && cell.s.text ? "'" : '') + v;
  }
  style(sheetIdx, r, c) { const cell = this.cellObj(this.sheets[sheetIdx], r, c); return (cell && cell.s) || {}; }
  setStyle(sheetIdx, r, c, st) { const cell = this.cellObj(this.sheets[sheetIdx], r, c, true); cell.s = Object.assign({}, cell.s, st); for (const k in cell.s) if (cell.s[k] === null || cell.s[k] === undefined) delete cell.s[k]; this.emit(); }
  formula(sheetIdx, r, c) { const cell = this.cellObj(this.sheets[sheetIdx], r, c); return cell && cell.f; }
  usedRange(s) { let g = null; for (const [k, cell] of s.cells) { if (isBlank(cell.v) && !cell.f) continue; const r = Math.floor(k / 16384), c = k % 16384; if (!g) g = { r1: r, c1: c, r2: r, c2: c }; else { g.r1 = Math.min(g.r1, r); g.c1 = Math.min(g.c1, c); g.r2 = Math.max(g.r2, r); g.c2 = Math.max(g.c2, c); } } return g; }
  invalidate() { this._cache = new Map(); this.circular = false; this._resetSpill(); this.emit(); }
  // ---- spilling: a formula whose answer is many cells fills the cells below and to the right ----
  _resetSpill() { this._spill = new Map(); this._anchors = new Map(); this._ready = new Set(); this._cands = new Map(); }
  _spillCands(si) { // formula cells that could give more than one cell
    if (this._cands.has(si)) return this._cands.get(si); const s = this.sheets[si]; const all = Object.keys(this.names).length > 0; const out = [];
    if (s) for (const [k, o] of s.cells) if (o.f && (all || formulaMayArr(o))) out.push([Math.floor(k / 16384), k % 16384]);
    this._cands.set(si, out); return out;
  }
  _scanSpills(si, r, c) { // while a formula is being worked out, only formulas above and to the left of (r,c) can spill into it
    if (!this._cache) this._cache = new Map(); const nested = this._cur.length > 0 && r !== undefined; let skipped = false;
    for (const [rr, cc] of this._spillCands(si)) { if (nested && (rr > r || cc > c)) { skipped = true; continue; } const key = si + ':' + (rr * 16384 + cc); const cv = this._cache.get(key); if (cv === IN_PROGRESS) { skipped = true; continue; } if (cv === undefined) this.value(si, rr, cc); }
    if (!skipped) this._ready.add(si);
  }
  // the spilled value shown in an empty cell, or null
  spillAt(si, r, c) { if (!this._cache) this._cache = new Map(); if (!this._ready.has(si)) this._scanSpills(si, r, c); return this._spill.get(si + ':' + (r * 16384 + c)) || null; }
  // the spill block that (r,c) belongs to: { g, anchor:{r,c}, blocked }
  spillBlock(si, r, c) {
    const own = this._anchors.get(si + ':' + (r * 16384 + c)); if (own || this.formula(si, r, c)) { this.value(si, r, c); const a = this._anchors.get(si + ':' + (r * 16384 + c)); return a ? { ...a, anchor: { r, c } } : null; }
    const e = this.spillAt(si, r, c); if (!e) return null; const a = this._anchors.get(si + ':' + (e.ar * 16384 + e.ac)); return a ? { ...a, anchor: { r: e.ar, c: e.ac } } : null;
  }
  hasSpills(si) { if (!this._spillCands(si).length) return false; if (!this._ready.has(si)) this._scanSpills(si); for (const k of this._anchors.keys()) if (k.startsWith(si + ':')) return true; return false; }
  _settle(si, r, c, v) {
    if (v && v.range) {
      const R = v.vals.length, C = R ? v.vals[0].length : 0;
      if (!R || !C) return ERR.CALC;
      if (R === 1 && C === 1) v = v.vals[0][0];
      else {
        const s = this.sheets[si]; const key = si + ':' + (r * 16384 + c); const g = { r1: r, c1: c, r2: r + R - 1, c2: c + C - 1 };
        if (g.r2 >= 1048576 || g.c2 >= 16384) { this._anchors.set(key, { g, blocked: true }); return ERR.SPILL; }
        for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { if (!i && !j) continue; const k2 = (r + i) * 16384 + c + j; const o = s.cells.get(k2); const sp = this._spill.get(si + ':' + k2);
          if ((o && (o.f || !isBlank(o.v))) || (sp && sp.key !== key)) { this._anchors.set(key, { g, blocked: true }); return ERR.SPILL; } }
        for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { if (!i && !j) continue; const x = v.vals[i][j]; this._spill.set(si + ':' + ((r + i) * 16384 + c + j), { v: isBlank(x) ? 0 : x, ar: r, ac: c, key }); }
        this._anchors.set(key, { g }); if (g.r2 >= s.rows) s.rows = g.r2 + 6; if (g.c2 >= s.cols) s.cols = g.c2 + 3;
        v = v.vals[0][0];
      }
    }
    return isBlank(v) ? 0 : v;
  }
  // every value a formula gives, as rows of cells (one cell for an ordinary formula)
  evaluateArray(f, si = 0) { try { const v = this.evalNode(parse(f), si); if (v && v.range) return v.vals.map(r => r.map(x => isBlank(x) ? 0 : x)); return [[isBlank(v) ? 0 : v]]; } catch (e) { return [[e instanceof XErr ? e : ERR.VALUE]]; } }
  // the whole answer of the formula in (r,c), spilled cells included
  spillArray(si, r, c) { const v = this.value(si, r, c); const a = this._anchors.get(si + ':' + (r * 16384 + c)); if (!a || a.blocked) return [[v]]; const out = []; for (let i = a.g.r1; i <= a.g.r2; i++) { const row = []; for (let j = a.g.c1; j <= a.g.c2; j++) row.push(this.value(si, i, j)); out.push(row); } return out; }
  onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  emit() { for (const fn of this.listeners) fn(); }
  sheetIndexByName(n) { if (n === null || n === undefined) return -1; const k = n.toLowerCase(); return this.sheets.findIndex(s => s.name.toLowerCase() === k); }
  defineName(name, ref) { this.names[name.toLowerCase()] = { name, ref }; this.invalidate(); }

  // computed value of a cell
  value(si, r, c) {
    const s = this.sheets[si]; if (!s) return ERR.REF; const cell = this.cellObj(s, r, c);
    if (!this._cache) this._cache = new Map();
    if (!cell || (!cell.f && isBlank(cell.v))) { const sp = this.spillAt(si, r, c); return sp ? sp.v : (cell ? cell.v : null); }
    if (!cell.f) return cell.v;
    const key = si + ':' + (r * 16384 + c);
    if (this._cache.has(key)) { const v = this._cache.get(key); if (v === IN_PROGRESS) { this.circular = true; return 0; } return v; }
    this._cache.set(key, IN_PROGRESS); this._cur.push({ si, r, c });
    let v; try { const ast = cell._ast && cell._astSrc === cell.f ? cell._ast : (cell._ast = parse(cell.f), cell._astSrc = cell.f, cell._ast); v = this._settle(si, r, c, this.evalNode(ast, si)); }
    catch (e) { v = e instanceof SyntaxError ? ERR.NAME : (e instanceof XErr ? e : ERR.VALUE); cell._parseError = e instanceof SyntaxError ? e.message : null; }
    finally { this._cur.pop(); }
    this._cache.set(key, v); return v;
  }
  // evaluate a formula in the context of a sheet without storing it
  evaluate(f, si = 0) { try { let v = this.evalNode(parse(f), si); if (v && v.range) v = v.vals.length && v.vals[0].length ? v.vals[0][0] : ERR.CALC; return isBlank(v) ? 0 : v; } catch (e) { return e instanceof SyntaxError ? ERR.NAME : e instanceof XErr ? e : ERR.VALUE; } }
  display(si, r, c, width) { const v = this.value(si, r, c); const st = this.style(si, r, c); return formatValue(v, st.fmt, width); }

  resolveSheet(name, si) { if (name === null || name === undefined) return si; const i = this.sheetIndexByName(name); if (i < 0) throw ERR.REF; return i; }
  rangeVal(si, g) {
    const vals = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(this.value(si, r, c)); vals.push(row); }
    const out = { range: true, vals, si, g }; if (g.r1 === g.r2 && g.c1 === g.c2) out.single = vals[0][0]; return out;
  }
  evalNode(n, si) {
    switch (n.k) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return n.v;
      case 'missing': return null;
      case 'ref': { const i = this.resolveSheet(n.sheet, si); const v = this.value(i, n.r, n.c); return { range: true, vals: [[v]], single: v, si: i, g: { r1: n.r, c1: n.c, r2: n.r, c2: n.c } }; }
      case 'range': { const i = this.resolveSheet(n.sheet, si); if (n.whole) { const sh = this.sheets[i]; const out = this.rangeVal(i, n.whole === 'col' ? { r1: 0, r2: sh.rows - 1, c1: n.c1, c2: n.c2 } : { r1: n.r1, r2: n.r2, c1: 0, c2: sh.cols - 1 }); out.whole = n.whole; return out; } return this.rangeVal(i, n); }
      case 'arr': return mkArr(n.v.map(r => r.slice()));
      case 'spillref': { const i = this.resolveSheet(n.sheet, si); this.value(i, n.r, n.c); const a = this._anchors.get(i + ':' + (n.r * 16384 + n.c)); if (!a || a.blocked) return ERR.REF; return this.rangeVal(i, a.g); }
      case 'at': { const v = this.evalNode(n.e, si); if (!isMulti(v)) return scalar(v); const cur = this._cur[this._cur.length - 1]; if (v.g && cur && cur.si === v.si) { const g = v.g; if (g.c1 === g.c2 && cur.r >= g.r1 && cur.r <= g.r2) return v.vals[cur.r - g.r1][0]; if (g.r1 === g.r2 && cur.c >= g.c1 && cur.c <= g.c2) return v.vals[0][cur.c - g.c1]; return ERR.VALUE; } return v.vals[0][0]; }
      case 'name': {
        const lk = n.name.toLowerCase(); for (let k = this._let.length - 1; k >= 0; k--) if (this._let[k].has(lk)) return this._let[k].get(lk);
        const d = this.names[lk]; if (!d) return ERR.NAME;
        const m = /^(?:'?(.+?)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(d.ref.replace(/^=/, ''));
        if (m) { const i = this.resolveSheet(m[1] ?? null, si); const g = parseRange(m[2].replace(/\$/g, '')); return this.rangeVal(i, g); }
        const v = this.evaluate(d.ref[0] === '=' ? d.ref : '=' + d.ref, si); return v;
      }
      case 'neg': case 'pct': { const x = this.evalNode(n.e, si); const f = y => { const v = num(y); return isErr(v) ? v : n.k === 'neg' ? -v : v / 100; }; return isMulti(x) ? mapArr([x], ([y]) => f(y)) : f(scalar(x)); }
      case 'bin': { const a = this.evalNode(n.a, si), b = this.evalNode(n.b, si); if (isMulti(a) || isMulti(b)) return mapArr([a, b], ([x, y]) => binop(n.op, x, y)); return binop(n.op, scalar(a), scalar(b)); }
      case 'call': {
        const fn = FN[n.fn]; if (!fn) return ERR.NAME;
        if (fn.lazy) return fn.run.call(this, n.args, si);
        const args = n.args.map(a => this.evalNode(a, si));
        return fn(args, this, si);
      }
    }
    return ERR.VALUE;
  }
  // copy a block of cells (values, formulas shifted, styles) — used by fill and paste
  copyCell(si, from, to, opts = {}) {
    const s = this.sheets[si]; const src = this.cellObj(s, from.r, from.c); const dst = this.cellObj(s, to.r, to.c, true);
    if (!src) { dst.v = null; dst.f = null; if (!opts.keepStyle) dst.s = null; return; }
    dst.f = src.f ? shiftFormula(src.f, to.r - from.r, to.c - from.c) : null; dst.v = src.f ? null : src.v; if (!opts.valuesOnly) dst.s = src.s ? Object.assign({}, src.s) : null;
  }
  toSpec() { // serialise back to authoring format
    return { sheets: this.sheets.map(s => { const cells = {}, styles = {}; for (const [k, cell] of s.cells) { const a = addr(Math.floor(k / 16384), k % 16384); if (cell.f) cells[a] = cell.f; else if (!isBlank(cell.v)) cells[a] = isErr(cell.v) ? cell.v.code : typeof cell.v === 'string' && (cell.v[0] === "'" || cell.v[0] === '=' || parseInput(cell.v).v !== cell.v) ? "'" + cell.v : cell.v; /* text that looks like a number stays text */ if (cell.s && Object.keys(cell.s).length) styles[a] = cell.s; } const widths = {}; for (const [c, w] of Object.entries(s.widths)) widths[colName(+c)] = w; return { name: s.name, cells, styles, widths, freeze: s.freeze, rows: s.rows, cols: s.cols }; }), names: Object.fromEntries(Object.values(this.names).map(d => [d.name, d.ref])) };
  }
}
const IN_PROGRESS = Symbol('calc');


// ---------- coercion ----------
function scalar(v) { if (v && v.range) return v.single !== undefined ? (isBlank(v.single) ? null : v.single) : ERR.VALUE; return v; }
function num(v) {
  if (isErr(v)) return v; if (isBlank(v)) return 0; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') { if (v.trim() === '') return ERR.VALUE; const p = parseInput(v); if (typeof p.v === 'number') return p.v; return ERR.VALUE; }
  return ERR.VALUE;
}
function str(v) { if (isErr(v)) return v; if (isBlank(v)) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return generalNum(v); return String(v); }
function bool(v) { if (isErr(v)) return v; if (isBlank(v)) return false; if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0; if (/^true$/i.test(v)) return true; if (/^false$/i.test(v)) return false; return ERR.VALUE; }
function typeRank(v) { return typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2; }
export function compareValues(a, b) {
  if (isBlank(a)) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (isBlank(b)) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeRank(a), tb = typeRank(b); if (ta !== tb) return ta - tb;
  if (ta === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  if (ta === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}
function binop(op, a, b) {
  if (isErr(a)) return a; if (isErr(b)) return b;
  switch (op) {
    case '&': return str(a) + str(b);
    case '=': return compareValues(a, b) === 0; case '<>': return compareValues(a, b) !== 0;
    case '<': return compareValues(a, b) < 0; case '>': return compareValues(a, b) > 0;
    case '<=': return compareValues(a, b) <= 0; case '>=': return compareValues(a, b) >= 0;
  }
  const x = num(a), y = num(b); if (isErr(x)) return x; if (isErr(y)) return y;
  switch (op) {
    case '+': return fix(x + y); case '-': return fix(x - y); case '*': return fix(x * y);
    case '/': return y === 0 ? ERR.DIV0 : fix(x / y);
    case '^': { const r = Math.pow(x, y); return isFinite(r) ? r : ERR.NUM; }
  }
  return ERR.VALUE;
}
function fix(x) { return Math.abs(x) < 1e15 ? +x.toPrecision(15) : x; }

// ---------- functions ----------
// Each gets evaluated args (ranges come as {range, vals}). Ranges skip text/blank/bool; direct args coerce.
function* flatNums(args, opts = {}) {
  for (const a of args) {
    if (a && a.range) { for (const row of a.vals) for (const v of row) { if (isErr(v)) throw v; if (typeof v === 'number') yield v; else if (opts.rangeBool && typeof v === 'boolean') yield v ? 1 : 0; } }
    else { if (isBlank(a)) { if (opts.blankZero) yield 0; continue; } if (isErr(a)) throw a; const n = num(a); if (isErr(n)) throw n; yield n; }
  }
}
function* flatAll(args) { for (const a of args) { if (a && a.range) { for (const row of a.vals) for (const v of row) yield v; } else yield a; } }
const wrap = fn => (args, wb, si) => { try { return fn(args, wb, si); } catch (e) { if (isErr(e)) return e; throw e; } };
const S = v => scalar(v);
function numArg(v) { const n = num(S(v)); if (isErr(n)) throw n; return n; }
function strArg(v) { const s = str(S(v)); if (isErr(s)) throw s; return s; }
function round(x, d) { const m = Math.pow(10, d); return Math.sign(x) * Math.round(Math.abs(x) * m + 1e-9) / m; }
function criteriaFn(crit) {
  crit = S(crit); if (typeof crit !== 'string') return v => compareValues(v, crit) === 0 && !isBlank(v);
  const m = /^(<>|<=|>=|=|<|>)?(.*)$/.exec(crit); const op = m[1] || '=', rhsText = m[2]; const p = parseInput(rhsText); let rhs = p.v;
  if (typeof rhs === 'string' && (op === '=' || op === '<>') && /[*?]/.test(rhs)) { const re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*') + '$', 'i'); return v => (op === '=') === (typeof v === 'string' && re.test(v)); }
  if (rhsText === '') return op === '<>' ? v => !isBlank(v) && v !== '' : v => isBlank(v) || v === '';
  return v => { if (isBlank(v)) return op === '<>'; if (typeof rhs === 'number' && typeof v !== 'number') return op === '<>'; const c = compareValues(v, rhs); return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op]; };
}
export const FN = {
  SUM: wrap(a => { let t = 0; for (const n of flatNums(a)) t += n; return fix(t); }),
  AVERAGE: wrap(a => { let t = 0, k = 0; for (const n of flatNums(a)) { t += n; k++; } return k ? fix(t / k) : ERR.DIV0; }),
  MIN: wrap(a => { let m = Infinity; for (const n of flatNums(a)) m = Math.min(m, n); return m === Infinity ? 0 : m; }),
  MAX: wrap(a => { let m = -Infinity; for (const n of flatNums(a)) m = Math.max(m, n); return m === -Infinity ? 0 : m; }),
  COUNT: wrap(a => { let k = 0; for (const x of a) { if (x && x.range) { for (const row of x.vals) for (const v of row) if (typeof v === 'number') k++; } else if (typeof x === 'number' || typeof x === 'boolean' || (typeof x === 'string' && typeof parseInput(x).v === 'number')) k++; } return k; }),
  COUNTA: wrap(a => { let k = 0; for (const v of flatAll(a)) if (!isBlank(v)) k++; return k; }),
  SUBTOTAL: wrap(a => { const code = Math.round(Number(a[0])) % 100; const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 9: 'SUM' }; const f = map[code]; if (!f) return ERR.VALUE; return FN[f](a.slice(1)); }),
  COUNTBLANK: wrap(a => { let k = 0; for (const v of flatAll(a)) if (isBlank(v) || v === '') k++; return k; }),
  PRODUCT: wrap(a => { let t = 1, any = false; for (const n of flatNums(a)) { t *= n; any = true; } return any ? fix(t) : 0; }),
  ROUND: wrap(a => round(numArg(a[0]), numArg(a[1] ?? 0))),
  ROUNDUP: wrap(a => { const d = numArg(a[1] ?? 0), m = Math.pow(10, d), x = numArg(a[0]); return Math.sign(x) * Math.ceil(+(Math.abs(x) * m).toPrecision(12)) / m; }),
  ROUNDDOWN: wrap(a => { const d = numArg(a[1] ?? 0), m = Math.pow(10, d), x = numArg(a[0]); return Math.sign(x) * Math.floor(+(Math.abs(x) * m).toPrecision(12)) / m; }),
  INT: wrap(a => Math.floor(numArg(a[0]))),
  ABS: wrap(a => Math.abs(numArg(a[0]))),
  SQRT: wrap(a => { const x = numArg(a[0]); return x < 0 ? ERR.NUM : Math.sqrt(x); }),
  POWER: wrap(a => Math.pow(numArg(a[0]), numArg(a[1]))),
  MOD: wrap(a => { const x = numArg(a[0]), y = numArg(a[1]); if (y === 0) return ERR.DIV0; return fix(x - y * Math.floor(x / y)); }),
  IF: { lazy: true, run(args, si) { const c = bool(S(this.evalNode(args[0], si))); if (isErr(c)) return c; const n = c ? args[1] : args[2]; if (!n) return c ? 0 : false; const v = this.evalNode(n, si); return n.k === 'missing' ? 0 : v; } },
  IFERROR: { lazy: true, run(args, si) { const v = S(this.evalNode(args[0], si)); return isErr(v) ? S(this.evalNode(args[1], si)) : v; } },
  AND: wrap(a => { let r = true, any = false; for (const v of flatAll(a)) { if (isBlank(v) || typeof v === 'string') continue; const b = bool(v); if (isErr(b)) return b; r = r && b; any = true; } return any ? r : ERR.VALUE; }),
  OR: wrap(a => { let r = false, any = false; for (const v of flatAll(a)) { if (isBlank(v) || typeof v === 'string') continue; const b = bool(v); if (isErr(b)) return b; r = r || b; any = true; } return any ? r : ERR.VALUE; }),
  NOT: wrap(a => { const b = bool(S(a[0])); return isErr(b) ? b : !b; }),
  LEN: wrap(a => strArg(a[0]).length),
  LEFT: wrap(a => strArg(a[0]).slice(0, a[1] === undefined ? 1 : numArg(a[1]))),
  RIGHT: wrap(a => { const s = strArg(a[0]), n = a[1] === undefined ? 1 : numArg(a[1]); return n ? s.slice(-n) : ''; }),
  MID: wrap(a => strArg(a[0]).substr(numArg(a[1]) - 1, numArg(a[2]))),
  UPPER: wrap(a => strArg(a[0]).toUpperCase()), LOWER: wrap(a => strArg(a[0]).toLowerCase()),
  PROPER: wrap(a => strArg(a[0]).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase())),
  TRIM: wrap(a => strArg(a[0]).replace(/ +/g, ' ').trim()),
  CONCAT: wrap(a => { let s = ''; for (const v of flatAll(a)) { if (isErr(v)) return v; s += str(v); } return s; }),
  CONCATENATE: wrap(a => a.map(x => strArg(x)).join('')),
  TEXT: wrap(a => formatValue(S(a[0]), strArg(a[1]))),
  VALUE: wrap(a => { const v = S(a[0]); if (typeof v === 'number') return v; const p = parseInput(str(v)); return typeof p.v === 'number' ? p.v : ERR.VALUE; }),
  TODAY: wrap(() => todaySerial()),
  NOW: wrap(() => { const n = new Date(); return todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400; }),
  DATE: wrap(a => dateToSerial(numArg(a[0]), numArg(a[1]), numArg(a[2]))),
  YEAR: wrap(a => serialToDate(numArg(a[0])).y), MONTH: wrap(a => serialToDate(numArg(a[0])).m), DAY: wrap(a => serialToDate(numArg(a[0])).d),
  SUMIF: wrap(a => { const rg = a[0], test = criteriaFn(a[1]), sr = a[2] || a[0]; let t = 0; if (!rg || !rg.range) throw ERR.VALUE; rg.vals.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = sr.range ? (sr.vals[i] || [])[j] : null; if (typeof x === 'number') t += x; } })); return fix(t); }),
  COUNTIF: wrap(a => { const rg = a[0], test = criteriaFn(a[1]); let k = 0; if (!rg || !rg.range) throw ERR.VALUE; for (const row of rg.vals) for (const v of row) if (test(v)) k++; return k; }),
  AVERAGEIF: wrap(a => { const rg = a[0], test = criteriaFn(a[1]), sr = a[2] || a[0]; let t = 0, k = 0; rg.vals.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = (sr.vals[i] || [])[j]; if (typeof x === 'number') { t += x; k++; } } })); return k ? fix(t / k) : ERR.DIV0; }),
  ISBLANK: wrap(a => { const v = S(a[0]); return isBlank(v); }), ISNUMBER: wrap(a => typeof S(a[0]) === 'number'), ISTEXT: wrap(a => typeof S(a[0]) === 'string'),
  ISERROR: { lazy: true, run(args, si) { return isErr(S(this.evalNode(args[0], si))); } },
  PI: wrap(() => Math.PI),
  RANK: wrap(a => rankEq(a)), 'RANK.EQ': wrap(a => rankEq(a)),
  LARGE: wrap(a => { const xs = [...flatNums([a[0]])].sort((x, y) => y - x); const k = numArg(a[1]); return k >= 1 && k <= xs.length ? xs[k - 1] : ERR.NUM; }),
  SMALL: wrap(a => { const xs = [...flatNums([a[0]])].sort((x, y) => x - y); const k = numArg(a[1]); return k >= 1 && k <= xs.length ? xs[k - 1] : ERR.NUM; }),
  MEDIAN: wrap(a => { const xs = [...flatNums(a)].sort((x, y) => x - y); if (!xs.length) return ERR.NUM; const m = xs.length >> 1; return xs.length % 2 ? xs[m] : (xs[m - 1] + xs[m]) / 2; }),
};
function rankEq(a) { const x = numArg(a[0]); const xs = [...flatNums([a[1]])]; const desc = !(a[2] && numArg(a[2])); if (!xs.includes(x)) return ERR.NA; return 1 + xs.filter(y => desc ? y > x : y < x).length; }

// ---------- arrays: many values at once (dynamic arrays) ----------
const isArr = v => !!(v && v.range);
const dimsOf = v => isArr(v) ? [v.vals.length, v.vals.length ? v.vals[0].length : 0] : [1, 1];
function isMulti(v) { return isArr(v) && (v.vals.length > 1 || (v.vals[0] && v.vals[0].length > 1)); }
export function mkArr(vals) { const o = { range: true, arr: true, vals }; if (vals.length === 1 && vals[0].length === 1) o.single = vals[0][0]; return o; }
const valsOf = v => isArr(v) ? v.vals : [[v]];
const toArr = v => isArr(v) ? v : mkArr([[v]]);
const transpose = vals => vals[0].map((_, j) => vals.map(r => r[j]));
function elAt(v, i, j) { if (!isArr(v)) return v; const R = v.vals.length, C = v.vals[0].length; const ii = R === 1 ? 0 : i, jj = C === 1 ? 0 : j; if (ii >= R || jj >= C) return ERR.NA; return v.vals[ii][jj]; }
// Excel's broadcasting: a single row or column stretches to match; a shorter array gives #N/A where it runs out
function mapArr(list, f) {
  let R = 1, C = 1; for (const v of list) { if (!isArr(v)) continue; const [r, c] = dimsOf(v); if (r > 1) R = Math.max(R, r); if (c > 1) C = Math.max(C, c); }
  const vals = []; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { let x = f(list.map(v => elAt(v, i, j)), i, j); if (isArr(x)) x = x.vals[0] ? x.vals[0][0] : ERR.CALC; row.push(x); } vals.push(row); }
  return mkArr(vals);
}
// functions that work on one value at a time give a whole array when handed one
const LIFT_POS = {};
function lift(name, pos) {
  const fn = FN[name]; if (!fn || fn.lazy) return;
  const isPos = !pos ? () => true : Array.isArray(pos) ? i => pos.includes(i) : pos; LIFT_POS[name] = isPos;
  FN[name] = (args, wb, si) => {
    const idx = []; for (let i = 0; i < args.length; i++) if (isPos(i, args.length) && isMulti(args[i])) idx.push(i);
    if (!idx.length) return fn(args, wb, si);
    return mapArr(idx.map(i => args[i]), els => { const a = args.slice(); idx.forEach((i, k) => { a[i] = els[k]; }); return fn(a, wb, si); });
  };
}
// does this part of a formula possibly give more than one cell? (only those formulas are checked for spilling)
const ALWAYS_ARR = new Set(['FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'TEXTSPLIT', 'TRANSPOSE', 'OFFSET', 'INDIRECT', 'EXPAND', 'WRAPROWS', 'WRAPCOLS', 'ANCHORARRAY', 'LET']);
const PASS_ALL = new Set(['IF', 'IFS', 'SWITCH', 'IFERROR', 'IFNA', 'CHOOSE']);
function mayArr(n) {
  if (!n) return false;
  switch (n.k) {
    case 'range': return !!n.whole || n.r1 !== n.r2 || n.c1 !== n.c2;
    case 'spillref': case 'name': return true;
    case 'arr': return n.v.length > 1 || n.v[0].length > 1;
    case 'neg': case 'pct': return mayArr(n.e);
    case 'bin': return mayArr(n.a) || mayArr(n.b);
    case 'call': {
      const f = n.fn, A = n.args;
      if (ALWAYS_ARR.has(f)) return true;
      if (f === 'ROW' || f === 'COLUMN') return A.length > 0 && mayArr(A[0]);
      if (f === 'INDEX') { const z = x => x && x.k === 'num' && x.v === 0; const twoD = x => x && x.k === 'range' && x.r1 !== x.r2 && x.c1 !== x.c2; return z(A[1]) || z(A[2]) || (A.length < 3 && twoD(A[0])) || mayArr(A[1]) || mayArr(A[2]) || (A[0] && A[0].k !== 'range' && mayArr(A[0])); }
      if (f === 'XLOOKUP') { if (mayArr(A[0])) return true; const la = A[1], ra = A[2]; if (!la || !ra || la.k !== 'range' || ra.k !== 'range') return mayArr(ra); return la.c1 === la.c2 ? ra.c2 > ra.c1 : ra.r2 > ra.r1; }
      if (PASS_ALL.has(f)) return A.some(mayArr);
      const pos = LIFT_POS[f]; if (pos) return A.some((x, i) => pos(i, A.length) && mayArr(x));
      return false;
    }
  }
  return false;
}
// Does a formula work on many values somewhere inside, like =SUM(--(C2:C9>=33))? Older files store such formulas as
// array formulas; without that, Excel applies implicit intersection (@) and gives a different answer.
export function needsArray(f) { try { return arrInside(parse(f)); } catch (e) { return false; } }
function arrInside(n) {
  if (!n) return false;
  switch (n.k) {
    case 'neg': case 'pct': return mayArr(n.e) || arrInside(n.e);
    case 'bin': return mayArr(n.a) || mayArr(n.b) || arrInside(n.a) || arrInside(n.b);
    case 'call': {
      if (ALWAYS_ARR.has(n.fn) && n.fn !== 'OFFSET' && n.fn !== 'INDIRECT') return true;
      const pos = LIFT_POS[n.fn]; if (pos && n.args.some((x, i) => pos(i, n.args.length) && mayArr(x))) return true;
      if (PASS_ALL.has(n.fn) && n.args.some(mayArr)) return true;
      return n.args.some(arrInside);
    }
  }
  return false;
}
function formulaMayArr(cell) { try { const ast = cell._ast && cell._astSrc === cell.f ? cell._ast : (cell._ast = parse(cell.f), cell._astSrc = cell.f, cell._ast); return mayArr(ast); } catch (e) { return false; } }

// ---------- argument helpers ----------
const given = v => v !== undefined && v !== null; // an argument that was typed (not left out)
const optNum = (v, d) => given(v) ? numArg(v) : d;
function boolArg(v) { const b = bool(S(v)); if (isErr(b)) throw b; return b; }
const optBool = (v, d) => given(v) ? boolArg(v) : d;
const leap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const dim = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate(); // days in month m (1-12) of year y
const ser = (y, m, d) => dateToSerial(y, m, d);
const ymd = s => serialToDate(s);
const wdOf = d => ((Math.floor(d) + 6) % 7 + 7) % 7; // 0 = Sunday
function dateArg(v) { const n = numArg(v); if (n < 0) throw ERR.NUM; return Math.floor(n); }
function holidaySet(h) { const set = new Set(); if (!given(h)) return set; for (const v of flatAll([h])) { if (isBlank(v)) continue; if (isErr(v)) throw v; const n = num(v); if (isErr(n)) throw ERR.VALUE; set.add(Math.floor(n)); } return set; }
function weekendMask(w) {
  const m = Array(7).fill(false); if (!given(w)) w = 1; else w = S(w);
  if (typeof w === 'string') { if (!/^[01]{7}$/.test(w) || w === '1111111') throw ERR.VALUE; for (let i = 0; i < 7; i++) m[(i + 1) % 7] = w[i] === '1'; return m; }
  const n = numArg(w); const pairs = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] };
  if (pairs[n]) { pairs[n].forEach(d => { m[d] = true; }); return m; }
  if (n >= 11 && n <= 17) { m[n - 11] = true; return m; }
  throw ERR.NUM;
}
function netdays(s, e, mask, hol) { let sign = 1; if (s > e) { [s, e] = [e, s]; sign = -1; } let n = 0; for (let d = s; d <= e; d++) if (!mask[wdOf(d)] && !hol.has(d)) n++; return n * sign; }
function workday(s, k, mask, hol) { let d = s, left = Math.abs(Math.trunc(k)); const step = k < 0 ? -1 : 1; let guard = 0; while (left > 0 && guard++ < 4e6) { d += step; if (!mask[wdOf(d)] && !hol.has(d)) left--; } return d; }
function isoWeek(s) { const wd = (wdOf(s) + 6) % 7; const th = s - wd + 3; const y = ymd(th).y; return Math.floor((th - ser(y, 1, 1)) / 7) + 1; }
function secsOf(v) { if (v < 0) throw ERR.NUM; return Math.round((v - Math.floor(v)) * 86400) % 86400; }
const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function wildRe(s, whole = true) { const b = s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '[\\s\\S]*' + (whole ? '' : '?')).replace(/\?/g, '[\\s\\S]').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?'); return new RegExp(whole ? '^' + b + '$' : b, 'i'); }
// exact matching for lookups: same kind of value, text ignores case, blanks never match; wildcards when allowed
function matcher(x, wild) {
  if (wild && typeof x === 'string' && /[*?~]/.test(x)) { const re = wildRe(x); return v => typeof v === 'string' && re.test(v); }
  return v => !isBlank(v) && !isBlank(x) && typeof v === typeof x && compareValues(v, x) === 0;
}
// approximate match the way Excel does it: a binary search that assumes the list is sorted
function approxIdx(list, x, desc) {
  let lo = 0, hi = list.length - 1, ans = -1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; const v = list[mid]; if (isBlank(v) || typeRank(v) !== typeRank(x)) { hi = mid - 1; continue; } const c = compareValues(v, x); if (!desc ? c <= 0 : c >= 0) { ans = mid; lo = mid + 1; } else hi = mid - 1; }
  return ans;
}
function xfind(list, x, mode, smode) {
  const order = smode === -1 || smode === -2 ? [...list.keys()].reverse() : [...list.keys()];
  const eq = matcher(x, mode === 2);
  for (const i of order) if (eq(list[i])) return i;
  if (mode === -1 || mode === 1) { let best = -1; for (const i of order) { const v = list[i]; if (isBlank(v) || typeRank(v) !== typeRank(x)) continue; const c = compareValues(v, x); if (mode === -1 ? c < 0 : c > 0) { if (best < 0 || (mode === -1 ? compareValues(v, list[best]) > 0 : compareValues(v, list[best]) < 0)) best = i; } } return best; }
  return -1;
}
const vec = a => { const [R, C] = dimsOf(a); if (R > 1 && C > 1) return null; return R > 1 ? a.vals.map(r => r[0]) : a.vals[0]; };
function sortCmp(a, b, o) { const ba = isBlank(a), bb = isBlank(b); if (ba || bb) return ba && bb ? 0 : ba ? 1 : -1; const ea = isErr(a), eb = isErr(b); if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1; return compareValues(a, b) * o; }
function ifsMask(pairs) {
  const [R, C] = dimsOf(pairs[0][0]); const mask = Array.from({ length: R }, () => Array(C).fill(true));
  for (const [rg, cr] of pairs) { if (!isArr(rg)) throw ERR.VALUE; const [r2, c2] = dimsOf(rg); if (r2 !== R || c2 !== C) throw ERR.VALUE; const test = criteriaFn(cr); for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) if (mask[i][j] && !test(rg.vals[i][j])) mask[i][j] = false; }
  return mask;
}
function ifsPick(a, first) { // SUMIFS-style: a[0] is the range to add up, then range/criteria pairs
  if (a.length < 3 || a.length % 2 === 0) throw ERR.VALUE; const pairs = []; for (let i = first; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
  const mask = ifsMask(pairs); const vr = a[0]; if (!isArr(vr)) throw ERR.VALUE; const [R, C] = dimsOf(pairs[0][0]); const [r2, c2] = dimsOf(vr); if (r2 !== R || c2 !== C) throw ERR.VALUE;
  const out = []; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) if (mask[i][j]) { const v = vr.vals[i][j]; if (isErr(v)) throw v; if (typeof v === 'number') out.push(v); } return out;
}
const numsOf = a => [...flatNums(a)];
const meanOf = x => x.reduce((s, v) => s + v, 0) / x.length;
function variance(x, sample) { if (x.length < (sample ? 2 : 1)) throw ERR.DIV0; const m = meanOf(x); return x.reduce((s, v) => s + (v - m) * (v - m), 0) / (x.length - (sample ? 1 : 0)); }
function pctInc(x, k) { if (!x.length || k < 0 || k > 1) throw ERR.NUM; x = x.slice().sort((p, q) => p - q); const pos = k * (x.length - 1), lo = Math.floor(pos); return fix(x[lo] + (x[Math.min(lo + 1, x.length - 1)] - x[lo]) * (pos - lo)); }
function pctExc(x, k) { x = x.slice().sort((p, q) => p - q); const pos = k * (x.length + 1) - 1; if (!x.length || pos < 0 || pos > x.length - 1) throw ERR.NUM; const lo = Math.floor(pos); return fix(x[lo] + (x[Math.min(lo + 1, x.length - 1)] - x[lo]) * (pos - lo)); }
function takeIdx(n, k, drop) {
  if (k === null) return [...Array(n).keys()]; const m = Math.min(Math.abs(k), n);
  if (!drop) return k > 0 ? [...Array(m).keys()] : [...Array(m).keys()].map(i => n - m + i);
  return k >= 0 ? [...Array(n - m).keys()].map(i => i + m) : [...Array(n - m).keys()];
}
function pickIdx(list, n) { return list.map(k => { if (k === 0 || Math.abs(k) > n) throw ERR.VALUE; return k > 0 ? k - 1 : n + k; }); }
function textAround(a, after) {
  const t = strArg(a[0]); const ds = valsOf(a[1]).flat().map(x => str(x)); const inst = Math.trunc(optNum(a[2], 1)); const ci = given(a[3]) && numArg(a[3]) === 1; const me = given(a[4]) && numArg(a[4]) === 1;
  if (inst === 0 || Math.abs(inst) > Math.max(1, t.length)) throw ERR.VALUE;
  const T = ci ? t.toLowerCase() : t; const D = ds.map(d => ci ? d.toLowerCase() : d).filter(d => d !== '');
  let h;
  if (!D.length) h = inst > 0 ? { k: 0, len: 0 } : { k: T.length, len: 0 };
  else {
    const hits = []; let i = 0; while (i <= T.length) { let best = null; for (const d of D) { const k = T.indexOf(d, i); if (k >= 0 && (!best || k < best.k || (k === best.k && d.length > best.len))) best = { k, len: d.length }; } if (!best) break; hits.push(best); i = best.k + best.len; }
    if (me) { if (inst > 0) hits.push({ k: T.length, len: 0 }); else hits.unshift({ k: 0, len: 0 }); }
    h = inst > 0 ? hits[inst - 1] : hits[hits.length + inst];
  }
  if (!h) return given(a[5]) ? S(a[5]) : ERR.NA;
  return after ? t.slice(h.k + h.len) : t.slice(0, h.k);
}
function stackRows(parts) { const W = Math.max(...parts.map(p => dimsOf(p)[1])); const vals = []; for (const p of parts) for (const r of p.vals) { const row = r.slice(); while (row.length < W) row.push(ERR.NA); vals.push(row); } return mkArr(vals); }
// SUBTOTAL and AGGREGATE leave out cells that are themselves SUBTOTAL or AGGREGATE results, so subtotals aren't counted twice
function skipNested(x, wb) { if (!x || !x.g || !wb) return x; const g = x.g; let any = false; const vals = x.vals.map((r, i) => r.map((v, j) => { const f = wb.formula(x.si, g.r1 + i, g.c1 + j); if (f && /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) { any = true; return null; } return v; })); return any ? mkArr(vals) : x; }
function lazyEach(run) { return { lazy: true, run }; }

Object.assign(FN, {
  // ---- logical ----
  TRUE: () => true, FALSE: () => false, NA: () => ERR.NA,
  IF: lazyEach(function (args, si) {
    const cv = this.evalNode(args[0], si);
    const branch = k => args[k] ? (args[k].k === 'missing' ? 0 : this.evalNode(args[k], si)) : (k === 2 ? false : 0);
    if (isMulti(cv)) { const tv = branch(1), fv = branch(2); return mapArr([cv, tv, fv], ([c, t, f]) => { const b = bool(c); if (isErr(b)) return b; return b ? t : f; }); }
    const c = bool(S(cv)); if (isErr(c)) return c; return branch(c ? 1 : 2);
  }),
  IFS: lazyEach(function (args, si) {
    if (args.length < 2 || args.length % 2) return ERR.VALUE;
    for (let k = 0; k < args.length; k += 2) {
      const cv = this.evalNode(args[k], si);
      if (isMulti(cv)) { const all = args.map((n, i) => i === k ? cv : this.evalNode(n, si)); return mapArr(all, xs => { for (let p = 0; p < xs.length; p += 2) { const b = bool(xs[p]); if (isErr(b)) return b; if (b) return xs[p + 1]; } return ERR.NA; }); }
      const b = bool(S(cv)); if (isErr(b)) return b; if (b) return this.evalNode(args[k + 1], si);
    }
    return ERR.NA;
  }),
  SWITCH: lazyEach(function (args, si) {
    if (args.length < 3) return ERR.VALUE; const pairs = Math.floor((args.length - 1) / 2), hasDef = (args.length - 1) % 2 === 1;
    const same = (x, v) => isBlank(x) ? (isBlank(v) || v === 0 || v === '') : !isBlank(v) && typeRank(x) === typeRank(v) && compareValues(x, v) === 0;
    const xv = this.evalNode(args[0], si);
    if (isMulti(xv)) { const all = [xv, ...args.slice(1).map(n => this.evalNode(n, si))]; return mapArr(all, xs => { const x = xs[0]; if (isErr(x)) return x; for (let k = 0; k < pairs; k++) { const v = xs[1 + 2 * k]; if (isErr(v)) return v; if (same(x, v)) return xs[2 + 2 * k]; } return hasDef ? xs[xs.length - 1] : ERR.NA; }); }
    const x = S(xv); if (isErr(x)) return x;
    for (let k = 0; k < pairs; k++) { const v = S(this.evalNode(args[1 + 2 * k], si)); if (isErr(v)) return v; if (same(x, v)) return this.evalNode(args[2 + 2 * k], si); }
    return hasDef ? this.evalNode(args[args.length - 1], si) : ERR.NA;
  }),
  IFERROR: lazyEach(function (args, si) { const v = this.evalNode(args[0], si); const fb = () => args[1] ? this.evalNode(args[1], si) : 0; if (isMulti(v)) { const f = fb(); return mapArr([v, f], ([x, y]) => isErr(x) ? (isBlank(y) ? 0 : y) : x); } const x = S(v); return isErr(x) ? fb() : x; }),
  IFNA: lazyEach(function (args, si) { const v = this.evalNode(args[0], si); const na = x => isErr(x) && x.code === '#N/A'; const fb = () => args[1] ? this.evalNode(args[1], si) : 0; if (isMulti(v)) { const f = fb(); return mapArr([v, f], ([x, y]) => na(x) ? (isBlank(y) ? 0 : y) : x); } const x = S(v); return na(x) ? fb() : x; }),
  XOR: wrap(a => { let n = 0, any = false; for (const v of flatAll(a)) { if (isBlank(v) || typeof v === 'string') continue; const b = bool(v); if (isErr(b)) return b; any = true; if (b) n++; } return any ? n % 2 === 1 : ERR.VALUE; }),
  ISERROR: wrap(a => isErr(S(a[0]))), ISNA: wrap(a => { const v = S(a[0]); return isErr(v) && v.code === '#N/A'; }), ISERR: wrap(a => { const v = S(a[0]); return isErr(v) && v.code !== '#N/A'; }),
  ISLOGICAL: wrap(a => typeof S(a[0]) === 'boolean'), ISNONTEXT: wrap(a => typeof S(a[0]) !== 'string'),
  ISEVEN: wrap(a => Math.trunc(numArg(a[0])) % 2 === 0), ISODD: wrap(a => Math.abs(Math.trunc(numArg(a[0])) % 2) === 1),
  N: wrap(a => { const v = S(a[0]); if (isErr(v)) return v; return typeof v === 'number' ? v : v === true ? 1 : 0; }), T: wrap(a => { const v = S(a[0]); if (isErr(v)) return v; return typeof v === 'string' ? v : ''; }),
  'ERROR.TYPE': wrap(a => { const v = S(a[0]); const i = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#CALC!'].indexOf(isErr(v) ? v.code : ''); return i >= 0 ? i + 1 : ERR.NA; }),
  // ---- text ----
  FIND: wrap(a => { const f = strArg(a[0]), t = strArg(a[1]); const st = Math.trunc(optNum(a[2], 1)); if (st < 1 || st > Math.max(1, t.length)) throw ERR.VALUE; const i = t.indexOf(f, st - 1); if (i < 0) throw ERR.VALUE; return i + 1; }),
  SEARCH: wrap(a => { const f = strArg(a[0]), t = strArg(a[1]); const st = Math.trunc(optNum(a[2], 1)); if (st < 1 || st > Math.max(1, t.length)) throw ERR.VALUE; if (f === '') return st; const m = wildRe(f, false).exec(t.slice(st - 1)); if (!m) throw ERR.VALUE; return m.index + st; }),
  SUBSTITUTE: wrap(a => { const t = strArg(a[0]), o = strArg(a[1]), n = strArg(a[2]); if (o === '') return t; if (!given(a[3])) return t.split(o).join(n); const k = Math.trunc(numArg(a[3])); if (k < 1) throw ERR.VALUE; let i = -o.length; for (let c = 0; c < k; c++) { i = t.indexOf(o, i + o.length); if (i < 0) return t; } return t.slice(0, i) + n + t.slice(i + o.length); }),
  REPLACE: wrap(a => { const t = strArg(a[0]); const s = Math.trunc(numArg(a[1])), n = Math.trunc(numArg(a[2])); const nw = strArg(a[3]); if (s < 1 || n < 0) throw ERR.VALUE; return t.slice(0, s - 1) + nw + t.slice(s - 1 + n); }),
  CLEAN: wrap(a => strArg(a[0]).replace(/[\x00-\x1F]/g, '')),
  REPT: wrap(a => { const t = strArg(a[0]), n = Math.trunc(numArg(a[1])); if (n < 0 || t.length * n > 32767) throw ERR.VALUE; return t.repeat(n); }),
  EXACT: wrap(a => strArg(a[0]) === strArg(a[1])),
  CHAR: wrap(a => { const n = Math.trunc(numArg(a[0])); if (n < 1 || n > 255) throw ERR.VALUE; return String.fromCharCode(n); }),
  CODE: wrap(a => { const t = strArg(a[0]); if (!t) throw ERR.VALUE; return t.charCodeAt(0); }),
  UNICHAR: wrap(a => { const n = Math.trunc(numArg(a[0])); if (n < 1) throw ERR.VALUE; return String.fromCodePoint(n); }), UNICODE: wrap(a => { const t = strArg(a[0]); if (!t) throw ERR.VALUE; return t.codePointAt(0); }),
  TEXTJOIN: wrap(a => { const delims = valsOf(a[0]).flat().map(x => str(x)); const ign = boolArg(a[1]); const parts = []; for (const v of flatAll(a.slice(2))) { if (isErr(v)) throw v; const s = str(v); if (ign && s === '') continue; parts.push(s); } let out = ''; parts.forEach((p, i) => { if (i) out += delims[(i - 1) % delims.length]; out += p; }); if (out.length > 32767) throw ERR.CALC; return out; }),
  TEXTSPLIT: wrap(a => {
    const t = strArg(a[0]); const list = x => given(x) ? valsOf(x).flat().map(v => str(v)).filter(v => v !== '') : [];
    const cd = list(a[1]), rd = list(a[2]); const ign = optBool(a[3], false); const ci = given(a[4]) && numArg(a[4]) === 1; const pad = given(a[5]) ? S(a[5]) : ERR.NA;
    if (!cd.length && !rd.length) throw ERR.VALUE;
    const split = (s, ds) => ds.length ? s.split(new RegExp(ds.slice().sort((p, q) => q.length - p.length).map(reEsc).join('|'), ci ? 'gi' : 'g')) : [s];
    let rows = split(t, rd); if (ign) rows = rows.filter(r => r !== '');
    let grid = rows.map(r => { let cs = split(r, cd); if (ign) cs = cs.filter(x => x !== ''); return cs; }); if (ign) grid = grid.filter(r => r.length);
    if (!grid.length) return ERR.CALC; const W = Math.max(...grid.map(r => r.length));
    return mkArr(grid.map(r => { const o = r.slice(); while (o.length < W) o.push(pad); return o; }));
  }),
  TEXTBEFORE: wrap(a => textAround(a, false)), TEXTAFTER: wrap(a => textAround(a, true)),
  NUMBERVALUE: wrap(a => { let t = strArg(a[0]).replace(/\s/g, ''); const dec = given(a[1]) ? strArg(a[1])[0] : '.', grp = given(a[2]) ? strArg(a[2])[0] : ','; if (t === '') return 0; let pct = 0; while (t.endsWith('%')) { pct++; t = t.slice(0, -1); } t = t.split(grp).join(''); if (dec !== '.') t = t.replace(dec, '.'); if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) throw ERR.VALUE; return parseFloat(t) / Math.pow(100, pct); }),
  // ---- dates and times ----
  DATE: wrap(a => { let y = Math.trunc(numArg(a[0])); if (y < 0 || y > 9999) throw ERR.NUM; if (y < 1900) y += 1900; const s = dateToSerial(y, Math.trunc(numArg(a[1])), Math.trunc(numArg(a[2]))); if (s < 0) throw ERR.NUM; return s; }),
  EDATE: wrap(a => { const s = dateArg(a[0]); const k = Math.trunc(numArg(a[1])); const { y, m, d } = ymd(s); const t = m - 1 + k; const ty = y + Math.floor(t / 12), tm = ((t % 12) + 12) % 12 + 1; return ser(ty, tm, Math.min(d, dim(ty, tm))); }),
  EOMONTH: wrap(a => { const s = dateArg(a[0]); const k = Math.trunc(numArg(a[1])); const { y, m } = ymd(s); const t = m - 1 + k; const ty = y + Math.floor(t / 12), tm = ((t % 12) + 12) % 12 + 1; return ser(ty, tm, dim(ty, tm)); }),
  WEEKDAY: wrap(a => { const s = dateArg(a[0]); const t = Math.trunc(optNum(a[1], 1)); const wd = wdOf(s); if (t === 1 || t === 17) return wd + 1; if (t === 2 || t === 11) return (wd + 6) % 7 + 1; if (t === 3) return (wd + 6) % 7; if (t >= 12 && t <= 16) return (wd - (t - 10) + 7) % 7 + 1; throw ERR.NUM; }),
  WEEKNUM: wrap(a => { const s = dateArg(a[0]); const t = Math.trunc(optNum(a[1], 1)); if (t === 21) return isoWeek(s); const st = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 }[t]; if (st === undefined) throw ERR.NUM; const j1 = ser(ymd(s).y, 1, 1); return Math.floor((s - j1 + (wdOf(j1) - st + 7) % 7) / 7) + 1; }),
  ISOWEEKNUM: wrap(a => isoWeek(dateArg(a[0]))),
  NETWORKDAYS: wrap(a => netdays(dateArg(a[0]), dateArg(a[1]), weekendMask(1), holidaySet(a[2]))),
  'NETWORKDAYS.INTL': wrap(a => netdays(dateArg(a[0]), dateArg(a[1]), weekendMask(a[2]), holidaySet(a[3]))),
  WORKDAY: wrap(a => workday(dateArg(a[0]), numArg(a[1]), weekendMask(1), holidaySet(a[2]))),
  'WORKDAY.INTL': wrap(a => workday(dateArg(a[0]), numArg(a[1]), weekendMask(a[2]), holidaySet(a[3]))),
  DATEDIF: wrap(a => {
    const s = dateArg(a[0]), e = dateArg(a[1]); const u = strArg(a[2]).toUpperCase(); if (s > e) throw ERR.NUM; const A = ymd(s), B = ymd(e);
    let months = (B.y - A.y) * 12 + (B.m - A.m); if (B.d < A.d) months--;
    switch (u) {
      case 'D': return e - s; case 'M': return months; case 'Y': return Math.floor(months / 12); case 'YM': return months % 12;
      case 'MD': { let d = B.d - A.d; if (d < 0) { const pm = B.m === 1 ? 12 : B.m - 1, py = B.m === 1 ? B.y - 1 : B.y; d += dim(py, pm); } return d; }
      case 'YD': { let st = ser(B.y, A.m, Math.min(A.d, dim(B.y, A.m))); if (st > e) st = ser(B.y - 1, A.m, Math.min(A.d, dim(B.y - 1, A.m))); return e - st; }
    }
    throw ERR.NUM;
  }),
  YEARFRAC: wrap(a => {
    let s = dateArg(a[0]), e = dateArg(a[1]); const b = Math.trunc(optNum(a[2], 0)); if (s > e) [s, e] = [e, s]; const A = ymd(s), B = ymd(e); const lastFeb = D => D.m === 2 && D.d === dim(D.y, 2);
    switch (b) {
      case 0: { let d1 = A.d, d2 = B.d; if (lastFeb(A) && lastFeb(B)) d2 = 30; if (lastFeb(A)) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; if (d1 === 31) d1 = 30; return fix(((B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1)) / 360); }
      case 1: { if (A.y === B.y || (B.y === A.y + 1 && (A.m > B.m || (A.m === B.m && A.d >= B.d)))) { const den = A.y === B.y ? (leap(A.y) ? 366 : 365) : ((leap(A.y) && s <= ser(A.y, 2, 29)) || (leap(B.y) && e >= ser(B.y, 2, 29))) ? 366 : 365; return fix((e - s) / den); } return fix((e - s) / ((ser(B.y + 1, 1, 1) - ser(A.y, 1, 1)) / (B.y - A.y + 1))); }
      case 2: return fix((e - s) / 360); case 3: return fix((e - s) / 365);
      case 4: return fix(((B.y - A.y) * 360 + (B.m - A.m) * 30 + (Math.min(B.d, 30) - Math.min(A.d, 30))) / 360);
    }
    throw ERR.NUM;
  }),
  DAYS: wrap(a => Math.floor(numArg(a[0])) - Math.floor(numArg(a[1]))),
  TIME: wrap(a => { const t = Math.trunc(numArg(a[0])) * 3600 + Math.trunc(numArg(a[1])) * 60 + Math.trunc(numArg(a[2])); if (t < 0) throw ERR.NUM; return (t % 86400) / 86400; }),
  HOUR: wrap(a => Math.floor(secsOf(numArg(a[0])) / 3600)), MINUTE: wrap(a => Math.floor(secsOf(numArg(a[0])) / 60) % 60), SECOND: wrap(a => secsOf(numArg(a[0])) % 60),
  DATEVALUE: wrap(a => { const v = S(a[0]); if (typeof v === 'number') throw ERR.VALUE; const p = parseInput(strArg(v)); if (typeof p.v !== 'number' || !p.fmt || !isDateFmt(p.fmt)) throw ERR.VALUE; return Math.floor(p.v); }),
  TIMEVALUE: wrap(a => { const p = parseInput(strArg(a[0])); if (typeof p.v !== 'number' || !p.fmt) throw ERR.VALUE; return p.v - Math.floor(p.v); }),
  // ---- maths ----
  MROUND: wrap(a => { const n = numArg(a[0]), m = numArg(a[1]); if (m === 0) return 0; if (n * m < 0) throw ERR.NUM; return fix(Math.sign(n) * Math.round(Math.abs(n / m) + 1e-9) * Math.abs(m)); }),
  'CEILING.MATH': wrap(a => { const n = numArg(a[0]); const s = given(a[1]) ? Math.abs(numArg(a[1])) : 1; const mode = optNum(a[2], 0); if (s === 0) return 0; const q = +(n / s).toPrecision(12); return fix(n >= 0 || !mode ? Math.ceil(q) * s : -Math.ceil(-q) * s); }),
  'FLOOR.MATH': wrap(a => { const n = numArg(a[0]); const s = given(a[1]) ? Math.abs(numArg(a[1])) : 1; const mode = optNum(a[2], 0); if (s === 0) return 0; const q = +(n / s).toPrecision(12); return fix(n >= 0 || !mode ? Math.floor(q) * s : -Math.floor(-q) * s); }),
  CEILING: wrap(a => { const n = numArg(a[0]), s = numArg(a[1]); if (s === 0) return 0; if (n > 0 && s < 0) throw ERR.NUM; const q = +(Math.abs(n) / Math.abs(s)).toPrecision(12); return fix(n < 0 && s < 0 ? -Math.ceil(q) * Math.abs(s) : n < 0 ? -Math.floor(q) * Math.abs(s) : Math.ceil(q) * Math.abs(s)); }),
  FLOOR: wrap(a => { const n = numArg(a[0]), s = numArg(a[1]); if (s === 0) throw ERR.DIV0; if (n > 0 && s < 0) throw ERR.NUM; const q = +(Math.abs(n) / Math.abs(s)).toPrecision(12); return fix(n < 0 && s < 0 ? -Math.floor(q) * Math.abs(s) : n < 0 ? -Math.ceil(q) * Math.abs(s) : Math.floor(q) * Math.abs(s)); }),
  TRUNC: wrap(a => { const d = Math.trunc(optNum(a[1], 0)); const m = Math.pow(10, d); return fix(Math.trunc(+(numArg(a[0]) * m).toPrecision(15)) / m); }),
  SIGN: wrap(a => Math.sign(numArg(a[0]))), EXP: wrap(a => Math.exp(numArg(a[0]))),
  LN: wrap(a => { const x = numArg(a[0]); if (x <= 0) throw ERR.NUM; return Math.log(x); }), LOG10: wrap(a => { const x = numArg(a[0]); if (x <= 0) throw ERR.NUM; return Math.log10(x); }),
  LOG: wrap(a => { const x = numArg(a[0]), b = optNum(a[1], 10); if (x <= 0 || b <= 0 || b === 1) throw ERR.NUM; return fix(Math.log(x) / Math.log(b)); }),
  EVEN: wrap(a => { const x = numArg(a[0]); const v = Math.ceil(Math.abs(x) / 2 - 1e-12) * 2; return x < 0 ? -v : v; }), ODD: wrap(a => { const x = numArg(a[0]); let v = Math.ceil(Math.abs(x) - 1e-12); if (v % 2 === 0) v++; return x < 0 ? -v : v; }),
  QUOTIENT: wrap(a => { const d = numArg(a[1]); if (d === 0) throw ERR.DIV0; return Math.trunc(numArg(a[0]) / d); }),
  RAND: wrap(() => Math.random()),
  RANDBETWEEN: wrap(a => { const lo = Math.ceil(numArg(a[0])), hi = Math.floor(numArg(a[1])); if (lo > hi) throw ERR.NUM; return lo + Math.floor(Math.random() * (hi - lo + 1)); }),
  RANDARRAY: wrap(a => { const R = Math.trunc(optNum(a[0], 1)), C = Math.trunc(optNum(a[1], 1)), lo = optNum(a[2], 0), hi = optNum(a[3], 1), whole = optBool(a[4], false); if (R < 1 || C < 1) return ERR.CALC; if (lo > hi || (whole && (lo % 1 || hi % 1))) throw ERR.VALUE; return mkArr(Array.from({ length: R }, () => Array.from({ length: C }, () => whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)))); }),
  // ---- conditional sums and counts ----
  COUNTIFS: wrap(a => { if (a.length < 2 || a.length % 2) throw ERR.VALUE; const pairs = []; for (let i = 0; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]); return ifsMask(pairs).flat().filter(Boolean).length; }),
  SUMIFS: wrap(a => fix(ifsPick(a, 1).reduce((s, v) => s + v, 0))),
  AVERAGEIFS: wrap(a => { const x = ifsPick(a, 1); return x.length ? fix(meanOf(x)) : ERR.DIV0; }),
  MAXIFS: wrap(a => { const x = ifsPick(a, 1); return x.length ? Math.max(...x) : 0; }), MINIFS: wrap(a => { const x = ifsPick(a, 1); return x.length ? Math.min(...x) : 0; }),
  SUMPRODUCT: wrap(a => { if (!a.length) throw ERR.VALUE; const arrs = a.map(x => isArr(x) ? x : mkArr([[x]])); const [R, C] = dimsOf(arrs[0]); for (const x of arrs) { const [r, c] = dimsOf(x); if (r !== R || c !== C) throw ERR.VALUE; } let t = 0; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { let p = 1; for (const x of arrs) { const v = x.vals[i][j]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : 0; } t += p; } return fix(t); }),
  SUBTOTAL: wrap((a, wb) => { const code = Math.trunc(numArg(a[0])) % 100; const name = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P' }[code]; if (!name) throw ERR.VALUE; return FN[name](a.slice(1).map(x => skipNested(x, wb))); }),
  AGGREGATE: wrap((a, wb) => {
    const fnum = Math.trunc(numArg(a[0])), opt = Math.trunc(optNum(a[1], 0)); const name = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' }[fnum];
    if (!name || opt < 0 || opt > 7) throw ERR.VALUE; const skipErr = [2, 3, 6, 7].includes(opt);
    const clean = x => { if (opt <= 3) x = skipNested(x, wb); return !skipErr || !isArr(x) ? x : mkArr(x.vals.map(r => r.map(v => isErr(v) ? null : v))); };
    return fnum >= 14 ? FN[name]([clean(a[2]), a[3]]) : FN[name](a.slice(2).map(clean));
  }),
  // ---- statistics ----
  'STDEV.S': wrap(a => Math.sqrt(variance(numsOf(a), true))), 'STDEV.P': wrap(a => Math.sqrt(variance(numsOf(a), false))),
  'VAR.S': wrap(a => fix(variance(numsOf(a), true))), 'VAR.P': wrap(a => fix(variance(numsOf(a), false))),
  'MODE.SNGL': wrap(a => { const x = numsOf(a); const c = new Map(); let best = null, bn = 1; for (const v of x) { const n = (c.get(v) || 0) + 1; c.set(v, n); if (n > bn) { bn = n; best = v; } } if (best === null) return ERR.NA; const first = x.find(v => c.get(v) === bn); return first; }),
  'PERCENTILE.INC': wrap(a => pctInc(numsOf([a[0]]), numArg(a[1]))), 'PERCENTILE.EXC': wrap(a => pctExc(numsOf([a[0]]), numArg(a[1]))),
  'QUARTILE.INC': wrap(a => { const q = Math.trunc(numArg(a[1])); if (q < 0 || q > 4) throw ERR.NUM; return pctInc(numsOf([a[0]]), q / 4); }),
  'QUARTILE.EXC': wrap(a => { const q = Math.trunc(numArg(a[1])); if (q < 1 || q > 3) throw ERR.NUM; return pctExc(numsOf([a[0]]), q / 4); }),
  AVERAGEA: wrap(a => { let t = 0, k = 0; for (const v of flatAll(a)) { if (isBlank(v)) continue; if (isErr(v)) return v; t += typeof v === 'number' ? v : v === true ? 1 : 0; k++; } return k ? fix(t / k) : ERR.DIV0; }),
  // ---- lookups and references ----
  VLOOKUP: wrap(a => { const x = S(a[0]); if (isErr(x)) throw x; const t = a[1]; if (!isArr(t)) throw ERR.VALUE; const ci = Math.trunc(numArg(a[2])); const approx = a[3] === undefined ? true : a[3] === null ? false : boolArg(a[3]); if (ci < 1) throw ERR.VALUE; if (ci > t.vals[0].length) throw ERR.REF; const col = t.vals.map(r => r[0]); const i = approx ? approxIdx(col, x) : col.findIndex(matcher(x, true)); return i < 0 ? ERR.NA : t.vals[i][ci - 1]; }),
  HLOOKUP: wrap(a => { const x = S(a[0]); if (isErr(x)) throw x; const t = a[1]; if (!isArr(t)) throw ERR.VALUE; const ri = Math.trunc(numArg(a[2])); const approx = a[3] === undefined ? true : a[3] === null ? false : boolArg(a[3]); if (ri < 1) throw ERR.VALUE; if (ri > t.vals.length) throw ERR.REF; const row = t.vals[0]; const i = approx ? approxIdx(row, x) : row.findIndex(matcher(x, true)); return i < 0 ? ERR.NA : t.vals[ri - 1][i]; }),
  MATCH: wrap(a => { const x = S(a[0]); if (isErr(x)) throw x; if (!isArr(a[1])) throw ERR.NA; const list = vec(a[1]); if (!list) throw ERR.NA; const t = a[2] === undefined ? 1 : a[2] === null ? 0 : Math.sign(numArg(a[2])); const i = t === 0 ? list.findIndex(matcher(x, true)) : approxIdx(list, x, t < 0); return i < 0 ? ERR.NA : i + 1; }),
  XMATCH: wrap(a => { const x = S(a[0]); if (isErr(x)) throw x; if (!isArr(a[1])) throw ERR.NA; const list = vec(a[1]); if (!list) throw ERR.VALUE; const i = xfind(list, x, Math.trunc(optNum(a[2], 0)), Math.trunc(optNum(a[3], 1))); return i < 0 ? ERR.NA : i + 1; }),
  XLOOKUP: wrap((a, wb) => {
    const x = S(a[0]); if (isErr(x)) throw x; const la = a[1], ra = a[2]; if (!isArr(la) || !isArr(ra)) throw ERR.VALUE;
    const [R, C] = dimsOf(la); if (R > 1 && C > 1) throw ERR.VALUE; const vertical = C === 1; const list = vertical ? la.vals.map(r => r[0]) : la.vals[0];
    const [RR, RC] = dimsOf(ra); if (vertical ? RR !== list.length : RC !== list.length) throw ERR.VALUE;
    const i = xfind(list, x, Math.trunc(optNum(a[4], 0)), Math.trunc(optNum(a[5], 1)));
    if (i < 0) return given(a[3]) ? a[3] : ERR.NA;
    if (ra.g) return wb.rangeVal(ra.si, vertical ? { r1: ra.g.r1 + i, r2: ra.g.r1 + i, c1: ra.g.c1, c2: ra.g.c2 } : { r1: ra.g.r1, r2: ra.g.r2, c1: ra.g.c1 + i, c2: ra.g.c1 + i });
    return mkArr(vertical ? [ra.vals[i].slice()] : ra.vals.map(r => [r[i]]));
  }),
  INDEX: wrap((a, wb) => {
    const arr = toArr(a[0]); const [R, C] = dimsOf(arr); let r = given(a[1]) ? Math.trunc(numArg(a[1])) : 0, c = given(a[2]) ? Math.trunc(numArg(a[2])) : 0;
    if (a[2] === undefined) { if (R === 1 && C > 1) { c = r; r = 1; } else if (C === 1) c = 1; }
    if (R === 1 && r === 0 && C === 1) r = 1;
    if (r < 0 || c < 0 || r > R || c > C) throw ERR.REF;
    const g = { r1: r ? r - 1 : 0, r2: r ? r - 1 : R - 1, c1: c ? c - 1 : 0, c2: c ? c - 1 : C - 1 };
    if (arr.g) return wb.rangeVal(arr.si, { r1: arr.g.r1 + g.r1, r2: arr.g.r1 + g.r2, c1: arr.g.c1 + g.c1, c2: arr.g.c1 + g.c2 });
    const out = []; for (let i = g.r1; i <= g.r2; i++) out.push(arr.vals[i].slice(g.c1, g.c2 + 1)); return mkArr(out);
  }),
  CHOOSE: lazyEach(function (args, si) {
    const iv = this.evalNode(args[0], si);
    if (isMulti(iv)) { const vs = args.slice(1).map(n => this.evalNode(n, si)); return mapArr([iv, ...vs], ([k, ...xs]) => { const n = num(k); if (isErr(n)) return n; const t = Math.trunc(n); return t >= 1 && t <= xs.length ? xs[t - 1] : ERR.VALUE; }); }
    const n = num(S(iv)); if (isErr(n)) return n; const t = Math.trunc(n); if (t < 1 || t >= args.length) return ERR.VALUE; return this.evalNode(args[t], si);
  }),
  LOOKUP: wrap(a => { const x = S(a[0]); if (isErr(x)) throw x; const lv = toArr(a[1]); const [R, C] = dimsOf(lv); let list, res; if (given(a[2])) { list = R >= C ? lv.vals.map(r => r[0]) : lv.vals[0]; const rv = toArr(a[2]); const [r2, c2] = dimsOf(rv); res = r2 >= c2 ? rv.vals.map(r => r[0]) : rv.vals[0]; } else if (R >= C) { list = lv.vals.map(r => r[0]); res = lv.vals.map(r => r[C - 1]); } else { list = lv.vals[0]; res = lv.vals[R - 1]; } const i = approxIdx(list, x); return i < 0 || i >= res.length ? ERR.NA : res[i]; }),
  ROW: lazyEach(function (args, si) { if (!args.length || args[0].k === 'missing') { const cur = this._cur[this._cur.length - 1]; return cur ? cur.r + 1 : ERR.VALUE; } const v = this.evalNode(args[0], si); if (!v || !v.g) return ERR.VALUE; const g = v.g; return g.r1 === g.r2 ? g.r1 + 1 : mkArr(Array.from({ length: g.r2 - g.r1 + 1 }, (_, i) => [g.r1 + i + 1])); }),
  COLUMN: lazyEach(function (args, si) { if (!args.length || args[0].k === 'missing') { const cur = this._cur[this._cur.length - 1]; return cur ? cur.c + 1 : ERR.VALUE; } const v = this.evalNode(args[0], si); if (!v || !v.g) return ERR.VALUE; const g = v.g; return g.c1 === g.c2 ? g.c1 + 1 : mkArr([Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => g.c1 + i + 1)]); }),
  ROWS: wrap(a => dimsOf(a[0])[0]), COLUMNS: wrap(a => dimsOf(a[0])[1]),
  TRANSPOSE: wrap(a => mkArr(transpose(toArr(a[0]).vals))),
  OFFSET: wrap((a, wb) => {
    const base = a[0]; if (!base || !base.g) throw ERR.VALUE; const g0 = base.g; const dr = Math.trunc(numArg(a[1])), dc = Math.trunc(numArg(a[2]));
    const h = given(a[3]) ? Math.trunc(numArg(a[3])) : g0.r2 - g0.r1 + 1, w = given(a[4]) ? Math.trunc(numArg(a[4])) : g0.c2 - g0.c1 + 1; if (!h || !w) throw ERR.REF;
    const r1 = g0.r1 + dr, c1 = g0.c1 + dc; const g = { r1: h > 0 ? r1 : r1 + h + 1, r2: h > 0 ? r1 + h - 1 : r1, c1: w > 0 ? c1 : c1 + w + 1, c2: w > 0 ? c1 + w - 1 : c1 };
    if (g.r1 < 0 || g.c1 < 0 || g.r2 >= 1048576 || g.c2 >= 16384) throw ERR.REF; return wb.rangeVal(base.si, g);
  }),
  INDIRECT: wrap((a, wb, si) => {
    const t = strArg(a[0]).trim(); const a1 = optBool(a[1], true); const m = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/.exec(t); let sheet = null, ref = t; if (m) { sheet = m[1] ? m[1].replace(/''/g, "'") : m[2]; ref = m[3]; }
    const i = sheet === null ? si : wb.sheetIndexByName(sheet); if (i < 0) throw ERR.REF;
    if (!a1) { const r = /^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(ref); if (!r) throw ERR.REF; const g = { r1: +r[1] - 1, c1: +r[2] - 1, r2: +(r[3] || r[1]) - 1, c2: +(r[4] || r[2]) - 1 }; return wb.rangeVal(i, { r1: Math.min(g.r1, g.r2), r2: Math.max(g.r1, g.r2), c1: Math.min(g.c1, g.c2), c2: Math.max(g.c1, g.c2) }); }
    const g = parseRange(ref.replace(/\$/g, '')); if (g) return wb.rangeVal(i, g);
    if (sheet === null && wb.names[ref.toLowerCase()]) return wb.evalNode({ k: 'name', name: ref }, si);
    throw ERR.REF;
  }),
  ADDRESS: wrap(a => { const r = Math.trunc(numArg(a[0])), c = Math.trunc(numArg(a[1])); const ab = Math.trunc(optNum(a[2], 1)); const a1 = optBool(a[3], true); const sh = given(a[4]) ? strArg(a[4]) : ''; if (r < 1 || c < 1 || ab < 1 || ab > 4) throw ERR.VALUE; const ra = ab === 1 || ab === 2, ca = ab === 1 || ab === 3; let s = a1 ? (ca ? '$' : '') + colName(c - 1) + (ra ? '$' : '') + r : `R${ra ? r : `[${r}]`}C${ca ? c : `[${c}]`}`; if (sh) s = (/[^A-Za-z0-9_]/.test(sh) ? `'${sh.replace(/'/g, "''")}'` : sh) + '!' + s; return s; }),
  FORMULATEXT: wrap((a, wb) => { const v = a[0]; if (!v || !v.g) throw ERR.NA; return wb.formula(v.si, v.g.r1, v.g.c1) || ERR.NA; }),
  ISFORMULA: wrap((a, wb) => { const v = a[0]; if (!v || !v.g) throw ERR.VALUE; return !!wb.formula(v.si, v.g.r1, v.g.c1); }),
  ANCHORARRAY: lazyEach(function (args, si) { const n = args[0]; if (!n || n.k !== 'ref') return ERR.REF; return this.evalNode({ k: 'spillref', sheet: n.sheet, r: n.r, c: n.c }, si); }),
  SINGLE: lazyEach(function (args, si) { return args[0] ? this.evalNode({ k: 'at', e: args[0] }, si) : ERR.VALUE; }),
  // ---- dynamic arrays ----
  FILTER: wrap(a => {
    const arr = toArr(a[0]), inc = toArr(a[1]); const [R, C] = dimsOf(arr), [ir, ic] = dimsOf(inc);
    const byRow = ic === 1 && ir === R ? true : ir === 1 && ic === C ? false : null; if (byRow === null) throw ERR.VALUE;
    const keep = []; for (let k = 0; k < (byRow ? R : C); k++) { const v = byRow ? inc.vals[k][0] : inc.vals[0][k]; if (isErr(v)) throw v; if (typeof v === 'string') throw ERR.VALUE; if (bool(v) === true) keep.push(k); }
    if (!keep.length) return given(a[2]) ? a[2] : ERR.CALC;
    return mkArr(byRow ? keep.map(k => arr.vals[k].slice()) : arr.vals.map(r => keep.map(k => r[k])));
  }),
  SORT: wrap(a => {
    const arr = toArr(a[0]); const [R, C] = dimsOf(arr); const byCol = optBool(a[3], false);
    const idxs = given(a[1]) ? valsOf(a[1]).flat().map(x => Math.trunc(numArg(x))) : [1]; const ords = given(a[2]) ? valsOf(a[2]).flat().map(x => numArg(x)) : [1];
    for (const k of idxs) if (k < 1 || k > (byCol ? R : C)) throw ERR.VALUE; for (const o of ords) if (o !== 1 && o !== -1) throw ERR.VALUE;
    const items = byCol ? transpose(arr.vals) : arr.vals;
    const sorted = items.map((r, i) => [r, i]).sort((A, B) => { for (let k = 0; k < idxs.length; k++) { const c = sortCmp(A[0][idxs[k] - 1], B[0][idxs[k] - 1], ords[Math.min(k, ords.length - 1)]); if (c) return c; } return A[1] - B[1]; }).map(x => x[0].slice());
    return mkArr(byCol ? transpose(sorted) : sorted);
  }),
  SORTBY: wrap(a => {
    const arr = toArr(a[0]); const [R, C] = dimsOf(arr); const keys = [];
    for (let i = 1; i < a.length; i += 2) { const by = toArr(a[i]); const o = given(a[i + 1]) ? numArg(a[i + 1]) : 1; if (o !== 1 && o !== -1) throw ERR.VALUE; const [br, bc] = dimsOf(by); keys.push({ by, o, col: br === 1 && bc > 1 }); }
    if (!keys.length) throw ERR.VALUE; const byCols = keys[0].col;
    for (const k of keys) { const [br, bc] = dimsOf(k.by); if (byCols ? (br !== 1 || bc !== C) : (bc !== 1 || br !== R)) throw ERR.VALUE; }
    const order = [...Array(byCols ? C : R).keys()].sort((x, y) => { for (const k of keys) { const c = sortCmp(byCols ? k.by.vals[0][x] : k.by.vals[x][0], byCols ? k.by.vals[0][y] : k.by.vals[y][0], k.o); if (c) return c; } return x - y; });
    return mkArr(byCols ? arr.vals.map(r => order.map(i => r[i])) : order.map(i => arr.vals[i].slice()));
  }),
  UNIQUE: wrap(a => {
    const arr = toArr(a[0]); const byCol = optBool(a[1], false), once = optBool(a[2], false); const items = byCol ? transpose(arr.vals) : arr.vals;
    const keyOf = r => JSON.stringify(r.map(v => isBlank(v) ? '' : typeof v === 'string' ? 's:' + v.toLowerCase() : isErr(v) ? 'e:' + v.code : typeof v + ':' + v));
    const count = new Map(), first = []; for (const r of items) { const k = keyOf(r); if (!count.has(k)) { count.set(k, 0); first.push([k, r]); } count.set(k, count.get(k) + 1); }
    const out = first.filter(([k]) => !once || count.get(k) === 1).map(([, r]) => r.slice()); if (!out.length) return ERR.CALC;
    return mkArr(byCol ? transpose(out) : out);
  }),
  SEQUENCE: wrap(a => { const R = Math.trunc(numArg(a[0])), C = given(a[1]) ? Math.trunc(numArg(a[1])) : 1; const st = optNum(a[2], 1), sp = optNum(a[3], 1); if (R < 1 || C < 1) return ERR.CALC; if (R * C > 1048576) throw ERR.NUM; const vals = []; let k = 0; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) row.push(fix(st + sp * k++)); vals.push(row); } return mkArr(vals); }),
  VSTACK: wrap(a => stackRows(a.map(toArr))),
  HSTACK: wrap(a => { const parts = a.map(toArr); return mkArr(transpose(stackRows(parts.map(p => mkArr(transpose(p.vals)))).vals)); }),
  TAKE: wrap(a => { const arr = toArr(a[0]); const [R, C] = dimsOf(arr); const ri = takeIdx(R, given(a[1]) ? Math.trunc(numArg(a[1])) : null, false), ci = takeIdx(C, given(a[2]) ? Math.trunc(numArg(a[2])) : null, false); if (!ri.length || !ci.length) return ERR.CALC; return mkArr(ri.map(i => ci.map(j => arr.vals[i][j]))); }),
  DROP: wrap(a => { const arr = toArr(a[0]); const [R, C] = dimsOf(arr); const ri = takeIdx(R, given(a[1]) ? Math.trunc(numArg(a[1])) : null, true), ci = takeIdx(C, given(a[2]) ? Math.trunc(numArg(a[2])) : null, true); if (!ri.length || !ci.length) return ERR.CALC; return mkArr(ri.map(i => ci.map(j => arr.vals[i][j]))); }),
  CHOOSECOLS: wrap(a => { const arr = toArr(a[0]); const C = dimsOf(arr)[1]; const idx = pickIdx(a.slice(1).flatMap(x => valsOf(x).flat()).map(x => Math.trunc(numArg(x))), C); if (!idx.length) throw ERR.VALUE; return mkArr(arr.vals.map(r => idx.map(j => r[j]))); }),
  CHOOSEROWS: wrap(a => { const arr = toArr(a[0]); const R = dimsOf(arr)[0]; const idx = pickIdx(a.slice(1).flatMap(x => valsOf(x).flat()).map(x => Math.trunc(numArg(x))), R); if (!idx.length) throw ERR.VALUE; return mkArr(idx.map(i => arr.vals[i].slice())); }),
  TOCOL: wrap(a => { const arr = toArr(a[0]); const ig = Math.trunc(optNum(a[1], 0)), byCol = optBool(a[2], false); let xs = (byCol ? transpose(arr.vals) : arr.vals).flat(); xs = xs.filter(v => !((ig === 1 || ig === 3) && isBlank(v)) && !((ig === 2 || ig === 3) && isErr(v))); if (!xs.length) return ERR.CALC; return mkArr(xs.map(v => [v])); }),
  TOROW: wrap(a => { const arr = toArr(a[0]); const ig = Math.trunc(optNum(a[1], 0)), byCol = optBool(a[2], false); let xs = (byCol ? transpose(arr.vals) : arr.vals).flat(); xs = xs.filter(v => !((ig === 1 || ig === 3) && isBlank(v)) && !((ig === 2 || ig === 3) && isErr(v))); if (!xs.length) return ERR.CALC; return mkArr([xs]); }),
  WRAPROWS: wrap(a => { const xs = toArr(a[0]).vals.flat(); const n = Math.trunc(numArg(a[1])); const pad = given(a[2]) ? S(a[2]) : ERR.NA; if (n < 1) throw ERR.NUM; const out = []; for (let i = 0; i < xs.length; i += n) { const r = xs.slice(i, i + n); while (r.length < n) r.push(pad); out.push(r); } return mkArr(out); }),
  WRAPCOLS: wrap(a => { const xs = toArr(a[0]).vals.flat(); const n = Math.trunc(numArg(a[1])); const pad = given(a[2]) ? S(a[2]) : ERR.NA; if (n < 1) throw ERR.NUM; const cols = []; for (let i = 0; i < xs.length; i += n) { const c = xs.slice(i, i + n); while (c.length < n) c.push(pad); cols.push(c); } return mkArr(transpose(cols)); }),
  EXPAND: wrap(a => { const arr = toArr(a[0]); const [R, C] = dimsOf(arr); const r = given(a[1]) ? Math.trunc(numArg(a[1])) : R, c = given(a[2]) ? Math.trunc(numArg(a[2])) : C; const pad = given(a[3]) ? S(a[3]) : ERR.NA; if (r < R || c < C) throw ERR.VALUE; return mkArr(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => i < R && j < C ? arr.vals[i][j] : pad))); }),
  LET: lazyEach(function (args, si) {
    if (args.length < 3 || args.length % 2 === 0) return ERR.VALUE; const scope = new Map(); this._let.push(scope);
    try { for (let i = 0; i < args.length - 1; i += 2) { const nm = args[i]; if (!nm || nm.k !== 'name') return ERR.VALUE; scope.set(nm.name.toLowerCase(), this.evalNode(args[i + 1], si)); } return this.evalNode(args[args.length - 1], si); }
    finally { this._let.pop(); }
  }),
});
FN.STDEV = FN['STDEV.S']; FN.VAR = FN['VAR.S']; FN.STDEVP = FN['STDEV.P']; FN.VARP = FN['VAR.P']; FN.MODE = FN['MODE.SNGL']; FN.PERCENTILE = FN['PERCENTILE.INC']; FN.QUARTILE = FN['QUARTILE.INC'];
// one value at a time, in every argument
for (const n of ['ABS', 'INT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'SQRT', 'POWER', 'MOD', 'NOT', 'LEN', 'LEFT', 'RIGHT', 'MID', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'TEXT', 'VALUE', 'DATE', 'YEAR', 'MONTH', 'DAY',
  'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISNA', 'ISERR', 'ISLOGICAL', 'ISNONTEXT', 'ISEVEN', 'ISODD', 'N', 'T', 'ERROR.TYPE', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'CLEAN', 'REPT', 'EXACT', 'CHAR', 'CODE', 'UNICHAR', 'UNICODE', 'NUMBERVALUE',
  'EDATE', 'EOMONTH', 'WEEKDAY', 'WEEKNUM', 'ISOWEEKNUM', 'DATEDIF', 'YEARFRAC', 'DAYS', 'TIME', 'HOUR', 'MINUTE', 'SECOND', 'DATEVALUE', 'TIMEVALUE',
  'MROUND', 'CEILING.MATH', 'FLOOR.MATH', 'CEILING', 'FLOOR', 'TRUNC', 'SIGN', 'EXP', 'LN', 'LOG10', 'LOG', 'EVEN', 'ODD', 'QUOTIENT', 'RANDBETWEEN', 'ADDRESS']) lift(n);
// one value at a time only in some arguments
lift('TEXTBEFORE', [0]); lift('TEXTAFTER', [0]); lift('NETWORKDAYS', [0, 1]); lift('NETWORKDAYS.INTL', [0, 1]); lift('WORKDAY', [0, 1]); lift('WORKDAY.INTL', [0, 1]);
lift('COUNTIF', [1]); lift('SUMIF', [1]); lift('AVERAGEIF', [1]); lift('COUNTIFS', i => i % 2 === 1); for (const n of ['SUMIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS']) lift(n, i => i >= 2 && i % 2 === 0);
lift('VLOOKUP', [0]); lift('HLOOKUP', [0]); lift('MATCH', [0]); lift('XMATCH', [0]); lift('XLOOKUP', [0]); lift('LOOKUP', [0]); lift('INDEX', [1, 2]);
lift('LARGE', [1]); lift('SMALL', [1]); lift('RANK', [0]); lift('RANK.EQ', [0]); lift('PERCENTILE.INC', [1]); lift('QUARTILE.INC', [1]);

// ---------- formulas as Excel stores them in a file ----------
// Newer functions carry a prefix in the file (_xlfn.XLOOKUP); without it Excel shows #NAME? until the cell is retyped.
const XLFN = new Set(['AGGREGATE', 'CEILING.MATH', 'FLOOR.MATH', 'CONCAT', 'DAYS', 'FORMULATEXT', 'IFNA', 'IFS', 'ISFORMULA', 'ISOWEEKNUM', 'MAXIFS', 'MINIFS', 'MODE.SNGL', 'NETWORKDAYS.INTL', 'WORKDAY.INTL', 'NUMBERVALUE', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'SWITCH', 'TEXTJOIN', 'UNICHAR', 'UNICODE', 'XOR', 'XLOOKUP', 'XMATCH', 'LET', 'RANDARRAY', 'SEQUENCE', 'UNIQUE', 'SORTBY', 'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND', 'ANCHORARRAY', 'SINGLE', 'ERROR.TYPE']);
export function toFileFormula(f, names = {}) {
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  const hasLet = toks.some(t => t.type === 'fn' && t.val === 'LET'); let out = '', last = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]; let rep = null, end = t.e;
    if (t.type === 'fn') rep = t.val === 'FILTER' || t.val === 'SORT' ? '_xlfn._xlws.' + t.val : XLFN.has(t.val) ? '_xlfn.' + t.val : null;
    else if (t.type === 'ref' && t.spill) rep = '_xlfn.ANCHORARRAY(' + t.sheetText + (t.cAbs ? '$' : '') + colName(t.c) + (t.rAbs ? '$' : '') + (t.r + 1) + ')';
    else if (t.type === 'name' && hasLet && !names[t.val.toLowerCase()]) rep = '_xlpm.' + t.val;
    else if (t.type === 'op' && t.val === '@') { const n1 = toks[i + 1]; if (n1 && n1.type === 'ref') { let j = i + 1; if (toks[j + 1] && toks[j + 1].val === ':' && toks[j + 2] && toks[j + 2].type === 'ref') j += 2; rep = '_xlfn.SINGLE(' + f.slice(n1.s, toks[j].e) + ')'; end = toks[j].e; i = j; } else rep = ''; }
    if (rep !== null) { out += f.slice(last, t.s) + rep; last = end; }
  }
  return out + f.slice(last);
}

export const FUNCTION_NAMES = Object.keys(FN).filter(n => n !== 'ANCHORARRAY' && n !== 'SINGLE').sort();

// ---------- number formats ----------
export function isDateFmt(f) { if (!f) return false; if (/\[(h+|m+|s+)\]/i.test(f)) return true; const g = f.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dy]/i.test(g) || /(^|[^0#?])m{3,}/i.test(g) || /h/i.test(g) && /:/.test(g); }
function generalNum(v) {
  if (v === 0) return '0'; const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) { let s = v.toExponential(5).replace(/\.?0+e/, 'e'); return s.replace('e', 'E').replace(/E\+?(-?)(\d)$/, 'E+$10$2').replace('E+-', 'E-'); }
  let s = String(+v.toPrecision(11)); if (s.includes('e')) s = v.toFixed(10).replace(/\.?0+$/, ''); return s;
}
function groupWestern(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function groupIndian(intStr) { if (intStr.length <= 3) return intStr; const last = intStr.slice(-3), rest = intStr.slice(0, -3); return rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last; }
export function formatValue(v, fmt, width) {
  if (isBlank(v)) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') { if (fmt && fmt.includes('@') ) { const sec = fmt.split(';'); const t = sec[3] || (sec.length === 1 ? sec[0] : null); if (t) return t.replace(/"([^"]*)"/g, '$1').replace('@', v); } return v; }
  if (!fmt || fmt === 'General') { const s = generalNum(v); if (width && s.length * 7.2 > width - 6) { // General squeezes decimals, then shows ####
      for (let p = 10; p >= 1; p--) { const t = String(+v.toPrecision(p)); if (t.length * 7.2 <= width - 6 && !t.includes('e')) return t; } return '#'.repeat(Math.max(1, Math.floor((width - 6) / 7.2))); } return s; }
  const secs = splitFmt(fmt); let sec = secs[0], val = v;
  const cond = x => /^\s*\[(<=|>=|<>|<|>|=)(-?\d+(?:\.\d+)?)\]/.exec(x);
  if (secs.some(x => cond(x))) { // [>=10000000]…;[>=100000]…;… : the first section whose condition holds
    const hold = (op, n) => ({ '<': v < n, '>': v > n, '=': v === n, '<=': v <= n, '>=': v >= n, '<>': v !== n }[op]);
    let pick = secs.slice(0, 2).find(x => { const c = cond(x); return c && hold(c[1], +c[2]); }); if (pick === undefined) pick = secs.find(x => !cond(x)) ?? secs[secs.length - 1];
    const out = fmtNumber(v, pick.replace(/^\s*\[[^\]]*\]/, ''), false); if (width && out.length * 7.2 > width - 6) return '#'.repeat(Math.max(1, Math.floor((width - 6) / 7.2))); return out;
  }
  if (secs.length > 1 && v < 0) { sec = secs[1]; val = -v; } else if (secs.length > 2 && v === 0) sec = secs[2];
  let out;
  if (isDateFmt(sec)) out = fmtDate(val, sec); else out = fmtNumber(val, sec, secs.length > 1 && v < 0);
  if (width && out.length * 7.2 > width - 6) return '#'.repeat(Math.max(1, Math.floor((width - 6) / 7.2)));
  return out;
}
function splitFmt(f) { const out = []; let cur = '', q = false; for (const ch of f) { if (ch === '"') q = !q; if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out; }
// a digit pattern with characters fixed between the digits, like ##\,##\,##0: digits fill from the right
function fmtTemplate(v, sec, negSection) {
  const toks = []; let i = 0; const src = sec.replace(/"([^"]*)"/g, (m, t) => t.split('').map(ch => '\\' + ch).join(''));
  while (i < src.length) { const ch = src[i]; if (ch === '\\') { toks.push({ lit: src[i + 1] || '' }); i += 2; } else if ('#0?'.includes(ch)) { toks.push({ ph: ch }); i++; } else if (ch === '.') { toks.push({ dot: true }); i++; } else { toks.push({ lit: ch, outside: true }); i++; } }
  const first = toks.findIndex(t => t.ph), last = toks.length - 1 - [...toks].reverse().findIndex(t => t.ph); const d = toks.findIndex(t => t.dot);
  const intEnd = d >= 0 && d < last ? d : last + 1; const decs = d >= 0 ? toks.slice(d + 1, last + 1).filter(t => t.ph).length : 0;
  const neg = v < 0 && !negSection; let [ip, fp] = round(Math.abs(v), decs).toFixed(decs).split('.'); if (ip === '0' && toks.slice(first, intEnd).every(t => !t.ph || t.ph !== '0')) ip = '';
  let out = '', k = ip.length - 1;
  for (let j = intEnd - 1; j >= first; j--) { const t = toks[j]; if (t.ph) { if (k >= 0) out = ip[k--] + out; else if (t.ph === '0') out = '0' + out; } else if (t.lit !== undefined && k >= 0) out = t.lit + out; }
  if (k >= 0) out = ip.slice(0, k + 1) + out;
  const pre = toks.slice(0, first).map(t => t.lit || '').join(''), post = toks.slice(last + 1).map(t => t.lit || '').join('');
  return (neg ? '-' : '') + pre + out + (decs ? '.' + fp : '') + post;
}
function fmtNumber(v, sec, negSection) {
  if (/[#0?]\\.[#0?]/.test(sec.replace(/"[^"]*"/g, ''))) return fmtTemplate(v, sec.replace(/\[(Red|Green|Blue|Black|Color\d+)\]/gi, ''), negSection);
  sec = sec.replace(/\[(Red|Green|Blue|Black|Color\d+)\]/gi, '').replace(/\[\$([^\]-]*)(-[^\]]*)?\]/g, '"$1"').replace(/_./g, ' ').replace(/\*./g, '');
  let pct = (sec.match(/%/g) || []).length; let x = v * Math.pow(100, pct);
  const m = /([#0?,.]*[0#?][#0?,.]*)/.exec(sec.replace(/"[^"]*"/g, m => 'Q'.repeat(m.length)));
  if (!m) return sec.replace(/"([^"]*)"/g, '$1');
  const code = sec.substr(m.index, m[0].length);
  const pre = sec.slice(0, m.index).replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const post = sec.slice(m.index + code.length).replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const dot = code.indexOf('.'); const decPart = dot >= 0 ? code.slice(dot + 1).replace(/,/g, '') : '';
  let scale = 0; let intCode = dot >= 0 ? code.slice(0, dot) : code; while (intCode.endsWith(',')) { scale++; intCode = intCode.slice(0, -1); } x = x / Math.pow(1000, scale);
  const decs = decPart.length; const neg = x < 0 && !negSection; x = Math.abs(x);
  let [ip, fp] = round(x, decs).toFixed(decs).split('.');
  const minInt = (intCode.match(/0/g) || []).length; if (ip === '0' && minInt === 0) ip = ''; while (ip.length < minInt) ip = '0' + ip;
  if (intCode.includes(',')) { const indian = /#,##,##/.test(intCode); ip = indian ? groupIndian(ip) : groupWestern(ip); }
  if (fp !== undefined) { const opt = (decPart.match(/#/g) || []).length; if (opt) fp = fp.replace(new RegExp('0{1,' + opt + '}$'), ''); }
  let body = ip + (decs ? '.' + (fp || '') : ''); if (body.endsWith('.') && !decPart.includes('0')) body = body.slice(0, -1);
  return (neg ? '-' : '') + pre + body + post;
}
function fmtDate(v, sec) {
  const d = serialToDate(v); const frac = v - Math.floor(v); let secs = Math.round(frac * 86400); const H = Math.floor(secs / 3600) % 24, Mi = Math.floor(secs / 60) % 60, Se = secs % 60;
  const ampm = /AM\/PM/i.test(sec); const h12 = ((H + 11) % 12) + 1;
  const tot = Math.round(v * 86400); // elapsed time: [h] keeps counting past 24 hours, [m] and [s] likewise
  const tokens = sec.replace(/\[(?!h+\]|m+\]|s+\])[^\]]*\]/gi, '').match(/"[^"]*"|\\.|\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|./gi) || [];
  let out = '', lastWasH = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i], tl = t.toLowerCase();
    if (t[0] === '"') { out += t.slice(1, -1); continue; } if (t[0] === '\\') { out += t[1]; continue; }
    if (t[0] === '[') { const k = tl[1], w = t.length - 2; const n = k === 'h' ? Math.floor(tot / 3600) : k === 'm' ? Math.floor(tot / 60) : tot; out += String(n).padStart(w, '0'); lastWasH = k === 'h'; continue; }
    if (tl === 'yyyy') out += d.y; else if (tl === 'yy') out += String(d.y).slice(-2);
    else if (tl === 'mmmm') out += cap(MONTHS_FULL[d.m - 1]); else if (tl === 'mmm') out += cap(MONTHS[d.m - 1]); else if (tl === 'mmmmm') out += MONTHS[d.m - 1][0].toUpperCase();
    else if (tl === 'mm' || tl === 'm') { const nextTok = (tokens.slice(i + 1).find(x => /[a-z]/i.test(x)) || '').toLowerCase(); const isMin = lastWasH || nextTok.startsWith('s'); const val = isMin ? Mi : d.m; out += tl === 'mm' ? String(val).padStart(2, '0') : val; }
    else if (tl === 'dddd') out += cap(DAYS_FULL[d.wd]); else if (tl === 'ddd') out += cap(DAYS[d.wd]); else if (tl === 'dd') out += String(d.d).padStart(2, '0'); else if (tl === 'd') out += d.d;
    else if (tl === 'hh') out += String(ampm ? h12 : H).padStart(2, '0'); else if (tl === 'h') out += ampm ? h12 : H;
    else if (tl === 'ss') out += String(Se).padStart(2, '0'); else if (tl === 's') out += Se;
    else if (tl === 'am/pm') out += H < 12 ? 'AM' : 'PM';
    else out += t;
    if (/[a-z]/i.test(t)) lastWasH = tl[0] === 'h';
  }
  return out;
}
const cap = s => s[0].toUpperCase() + s.slice(1);

// ---------- AutoFill ----------
// Given the source values (typed inputs) of a selection along the fill direction, produce the next n values.
export function fillSeries(src, n, opts = {}) {
  // src: array of {v, f, s} cells (one line). Returns array of n {v,f,s, from:index}
  const out = []; const len = src.length;
  const nums = src.map(c => c && !c.f && typeof c.v === 'number' ? c.v : null);
  const allNum = nums.every(x => x !== null);
  const textNum = src.map(c => c && !c.f && typeof c.v === 'string' ? /^(.*?)(\d+)$/.exec(c.v) : null);
  const monthOf = src.map(c => c && !c.f && typeof c.v === 'string' ? monthWord(c.v) : null);
  const dayOf = src.map(c => c && !c.f && typeof c.v === 'string' ? dayWord(c.v) : null);
  const isDate = src.map(c => c && c.s && c.s.fmt && isDateFmt(c.s.fmt));
  for (let i = 0; i < n; i++) {
    const k = i % len; const base = src[k]; const step = Math.floor(i / len) + 1;
    if (!base) { out.push({ v: null, f: null, s: null, from: k }); continue; }
    if (base.f) { out.push({ f: base.f, from: k, shift: true, s: base.s }); continue; }
    if (opts.copy) { out.push({ v: base.v, s: base.s, from: k }); continue; }
    if (allNum && (len > 1 || isDate[0] || opts.series)) {
      let d = len > 1 ? (nums[len - 1] - nums[0]) / (len - 1) : 1;
      if (len > 1 && isDate[0] && opts.monthStep) d = 0;
      out.push({ v: fix(nums[len - 1] + d * (i + 1)), s: src[len - 1].s, from: len - 1 }); continue;
    }
    if (monthOf.every(x => x)) { const st = len > 1 ? (monthOf[len - 1].i - monthOf[0].i + 12) % 12 / (len - 1) || 1 : 1; const m = monthOf[len - 1]; const idx = ((m.i + st * (i + 1)) % 12 + 12) % 12; out.push({ v: m.style(idx), s: src[len - 1].s, from: len - 1 }); continue; }
    if (dayOf.every(x => x)) { const st = len > 1 ? ((dayOf[len - 1].i - dayOf[0].i + 7) % 7) / (len - 1) || 1 : 1; const dd = dayOf[len - 1]; const idx = ((dd.i + st * (i + 1)) % 7 + 7) % 7; out.push({ v: dd.style(idx), s: src[len - 1].s, from: len - 1 }); continue; }
    if (textNum[k] && textNum.every(x => x)) { const lastN = textNum[len - 1]; const d = len > 1 ? (parseInt(lastN[2]) - parseInt(textNum[0][2])) / (len - 1) || 1 : 1; const nv = parseInt(lastN[2]) + d * (i + 1); out.push({ v: lastN[1] + String(nv).padStart(lastN[2].length, '0'), s: src[len - 1].s, from: len - 1 }); continue; }
    out.push({ v: base.v, s: base.s, from: k });
  }
  return out;
}
function monthWord(s) { const w = s.trim(); const lo = w.toLowerCase(); let i = MONTHS_FULL.indexOf(lo); if (i >= 0) return { i, style: j => caseLike(w, MONTHS_FULL[j]) }; i = MONTHS.indexOf(lo); if (i >= 0) return { i, style: j => caseLike(w, MONTHS[j]) }; return null; }
function dayWord(s) { const w = s.trim(); const lo = w.toLowerCase(); let i = DAYS_FULL.indexOf(lo); if (i >= 0) return { i, style: j => caseLike(w, DAYS_FULL[j]) }; i = DAYS.indexOf(lo); if (i >= 0) return { i, style: j => caseLike(w, DAYS[j]) }; return null; }
function caseLike(model, word) { if (model === model.toUpperCase()) return word.toUpperCase(); if (model[0] === model[0].toUpperCase()) return cap(word); return word; }

// ---------- formula explainer ----------
// Returns steps: [{label, text, value}] walking from references to the final result.
export function explainFormula(wb, si, f) {
  const steps = []; let ast;
  try { ast = parse(f); } catch (e) { return [{ label: 'This formula has a typo', text: e.message, value: ERR.NAME }]; }
  const fmtV = v => v && v.range ? (v.single !== undefined ? show(v.single) : `${v.vals.length} × ${v.vals[0].length} values: ` + v.vals.flat().slice(0, 6).map(show).join(', ') + (v.vals.flat().length > 6 ? ' …' : '')) : show(v);
  const show = v => isBlank(v) ? '(blank → 0)' : isErr(v) ? v.code : typeof v === 'string' ? '"' + v + '"' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : generalNum(v);
  const text = n => {
    switch (n.k) {
      case 'num': return generalNum(n.v); case 'str': return '"' + n.v + '"'; case 'bool': return n.v ? 'TRUE' : 'FALSE'; case 'err': return n.v.code;
      case 'ref': return (n.sheet ? n.sheet + '!' : '') + addr(n.r, n.c); case 'range': return (n.sheet ? n.sheet + '!' : '') + (n.whole === 'col' ? colName(n.c1) + ':' + colName(n.c2) : n.whole === 'row' ? (n.r1 + 1) + ':' + (n.r2 + 1) : addr(n.r1, n.c1) + ':' + addr(n.r2, n.c2));
      case 'name': return n.name; case 'neg': return '-' + text(n.e); case 'pct': return text(n.e) + '%';
      case 'bin': return text(n.a) + n.op + text(n.b); case 'call': return n.fn + '(' + n.args.map(text).join(',') + ')'; case 'missing': return '';
      case 'arr': return '{' + n.v.map(r => r.map(show).join(',')).join(';') + '}'; case 'spillref': return (n.sheet ? n.sheet + '!' : '') + addr(n.r, n.c) + '#'; case 'at': return '@' + text(n.e);
    }
  };
  const walk = n => {
    if (n.k === 'ref' || n.k === 'name') { const v = wb.evalNode(n, si); steps.push({ label: n.k === 'name' ? 'Name' : 'Cell', text: text(n), value: fmtV(v) }); return; }
    if (n.k === 'spillref') { const v = wb.evalNode(n, si); steps.push({ label: 'Spill range', text: text(n), value: fmtV(v) }); return; }
    if (n.k === 'at') { walk(n.e); steps.push({ label: 'One value', text: text(n), value: fmtV(wb.evalNode(n, si)) }); return; }
    if (n.k === 'range') { const v = wb.evalNode(n, si); const flat = v.vals.flat().filter(x => !isBlank(x)).slice(0, 8).map(show); steps.push({ label: 'Range', text: text(n), value: flat.join(', ') + (v.vals.flat().length > 8 ? ' …' : '') }); return; }
    if (n.k === 'bin') { walk(n.a); walk(n.b); steps.push({ label: 'Calculate', text: text(n), value: fmtV(wb.evalNode(n, si)) }); return; }
    if (n.k === 'neg' || n.k === 'pct') { walk(n.e); steps.push({ label: 'Calculate', text: text(n), value: fmtV(wb.evalNode(n, si)) }); return; }
    if (n.k === 'call') { n.args.forEach(walk); steps.push({ label: 'Function', text: text(n), value: fmtV(wb.evalNode(n, si)) }); return; }
  };
  walk(ast); return steps;
}

// ---------- structure: insert/delete rows & columns, sort, regions, snapshots ----------
function adjustFormula(f, sheetName, ownSheet, axis, at, n) {
  // axis 'r' or 'c'; n>0 insert n at index `at`; n<0 delete -n starting at `at`
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  let out = '', last = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i]; if (t.type !== 'ref') continue;
    const target = t.sheet === null ? ownSheet : t.sheet;
    if (target.toLowerCase() !== sheetName.toLowerCase()) continue;
    let r = t.r, c = t.c; let broken = false;
    const isRangeEnd = i > 0 && toks[i - 1].type === 'op' && toks[i - 1].val === ':';
    const isRangeStart = toks[i + 1] && toks[i + 1].type === 'op' && toks[i + 1].val === ':' && toks[i + 2] && toks[i + 2].type === 'ref';
    const pos = axis === 'r' ? r : c;
    let np = pos;
    if (n > 0) { if (pos >= at) np = pos + n; }
    else { const k = -n; if (pos >= at + k) np = pos - k; else if (pos >= at) { if (isRangeStart) np = at; else if (isRangeEnd) np = at - 1; else broken = true; } }
    if (np < 0) broken = true;
    if (axis === 'r') r = np; else c = np;
    out += f.slice(last, t.s) + (broken ? '#REF!' : t.sheetText + (t.cAbs ? '$' : '') + colName(c) + (t.rAbs ? '$' : '') + (r + 1) + (t.spill ? '#' : ''));
    last = t.e;
  }
  return out + f.slice(last);
}
Workbook.prototype.snapshot = function () { return this.sheets.map(s => ({ name: s.name, rows: s.rows, cols: s.cols, widths: { ...s.widths }, cells: new Map([...s.cells].map(([k, o]) => [k, { v: o.v, f: o.f, s: o.s ? { ...o.s } : null }])) })).concat([{ names: JSON.parse(JSON.stringify(this.names)) }]); };
Workbook.prototype.restore = function (snap) { const names = snap[snap.length - 1].names; snap.slice(0, -1).forEach((sn, i) => { let s = this.sheets[i]; if (!s) s = this.addSheet(sn.name); s.name = sn.name; s.rows = sn.rows; s.cols = sn.cols; s.widths = { ...sn.widths }; s.cells = new Map([...sn.cells].map(([k, o]) => [k, { ...o, s: o.s ? { ...o.s } : null }])); }); this.sheets.length = snap.length - 1; this.names = names; this.invalidate(); };
Workbook.prototype.shiftStructure = function (si, axis, at, n) {
  const s = this.sheets[si]; const moved = new Map();
  for (const [k, o] of s.cells) {
    let r = Math.floor(k / 16384), c = k % 16384; const pos = axis === 'r' ? r : c;
    if (n < 0 && pos >= at && pos < at - n) continue;
    let np = pos; if (n > 0 && pos >= at) np = pos + n; if (n < 0 && pos >= at - n) np = pos + n;
    if (axis === 'r') r = np; else c = np; moved.set(r * 16384 + c, o);
  }
  s.cells = moved;
  if (axis === 'c') { const w = {}; for (const [c, v] of Object.entries(s.widths)) { const ci = +c; if (n < 0 && ci >= at && ci < at - n) continue; w[n > 0 ? (ci >= at ? ci + n : ci) : (ci >= at - n ? ci + n : ci)] = v; } s.widths = w; s.cols = Math.max(s.cols + n, 10); }
  else s.rows = Math.max(s.rows + n, 20);
  for (const sh of this.sheets) for (const o of sh.cells.values()) if (o.f) o.f = adjustFormula(o.f, s.name, sh.name, axis, at, n);
  for (const d of Object.values(this.names)) { const f = adjustFormula('=' + d.ref.replace(/^=/, ''), s.name, '\u0000', axis, at, n); d.ref = f.slice(1); }
  this.invalidate();
};
Workbook.prototype.insertRows = function (si, at, n = 1) { this.shiftStructure(si, 'r', at, n); };
Workbook.prototype.deleteRows = function (si, at, n = 1) { this.shiftStructure(si, 'r', at, -n); };
Workbook.prototype.insertCols = function (si, at, n = 1) { this.shiftStructure(si, 'c', at, n); };
Workbook.prototype.deleteCols = function (si, at, n = 1) { this.shiftStructure(si, 'c', at, -n); };
Workbook.prototype.isEmpty = function (si, r, c) { const o = this.cellObj(this.sheets[si], r, c); return (!o || (isBlank(o.v) && !o.f)) && !this.spillAt(si, r, c); };
// The block of filled cells around (r,c), like Ctrl+A / Ctrl+Shift+8
Workbook.prototype.currentRegion = function (si, r, c) {
  let g = { r1: r, c1: c, r2: r, c2: c }; let grown = true;
  const filled = (rr, cc) => rr >= 0 && cc >= 0 && !this.isEmpty(si, rr, cc);
  while (grown) {
    grown = false;
    const tryEdge = (cond, apply) => { if (cond()) { apply(); grown = true; } };
    tryEdge(() => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (filled(g.r1 - 1, cc)) return true; return false; }, () => g.r1--);
    tryEdge(() => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (filled(g.r2 + 1, cc)) return true; return false; }, () => g.r2++);
    tryEdge(() => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (filled(rr, g.c1 - 1)) return true; return false; }, () => g.c1--);
    tryEdge(() => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (filled(rr, g.c2 + 1)) return true; return false; }, () => g.c2++);
    g.r1 = Math.max(0, g.r1); g.c1 = Math.max(0, g.c1);
  }
  return g;
};
// Sort rows of region g by column `byCol` (absolute index); header row kept if hasHeader
Workbook.prototype.sortRegion = function (si, g, keys, hasHeader) {
  const s = this.sheets[si]; const start = hasHeader ? g.r1 + 1 : g.r1;
  const rows = []; for (let r = start; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) cells.push(this.cellObj(s, r, c)); rows.push({ r, cells, vals: keys.map(k => this.value(si, r, k.col)) }); }
  rows.sort((a, b) => { for (let i = 0; i < keys.length; i++) { const x = a.vals[i], y = b.vals[i]; if (isBlank(x) && !isBlank(y)) return 1; if (isBlank(y) && !isBlank(x)) return -1; const c = compareValues(x, y); if (c) return keys[i].desc ? -c : c; } return a.r - b.r; });
  rows.forEach((row, i) => { const nr = start + i; row.cells.forEach((o, j) => { const k = nr * 16384 + g.c1 + j; if (o) { const copy = { v: o.v, f: o.f ? shiftFormula(o.f, nr - row.r, 0) : null, s: o.s }; s.cells.set(k, copy); } else s.cells.delete(k); }); });
  this.invalidate();
};
Workbook.prototype.guessHeader = function (si, g) { if (g.r2 <= g.r1) return false; let textTop = 0, numBelow = 0; for (let c = g.c1; c <= g.c2; c++) { const a = this.value(si, g.r1, c), b = this.value(si, g.r1 + 1, c); if (typeof a === 'string') textTop++; if (typeof b !== 'string' || typeof a === 'string') numBelow++; } return textTop === g.c2 - g.c1 + 1; };
