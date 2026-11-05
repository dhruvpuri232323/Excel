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
export const ERR = { DIV0: new XErr('#DIV/0!'), NAME: new XErr('#NAME?'), VALUE: new XErr('#VALUE!'), REF: new XErr('#REF!'), NA: new XErr('#N/A'), NUM: new XErr('#NUM!'), NULL: new XErr('#NULL!') };
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
    if (ch === '#') { const m = /^#(DIV\/0!|NAME\?|VALUE!|REF!|N\/A|NUM!|NULL!)/i.exec(src.slice(i)); if (m) { push('err', ERR_BY_CODE[m[0].toUpperCase()], i, i + m[0].length); i += m[0].length; continue; } }
    if (/[0-9.]/.test(ch)) { const m = /^(\d+\.?\d*|\.\d+)(E[-+]?\d+)?/i.exec(src.slice(i)); if (m) { push('num', parseFloat(m[0]), i, i + m[0].length); i += m[0].length; continue; } }
    // sheet-qualified or plain reference / name / function
    let m = /^('((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!/.exec(src.slice(i));
    let sheet = null, start = i;
    if (m) { sheet = m[2] ? m[2].replace(/''/g, "'") : m[3]; i += m[0].length; }
    m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/.exec(src.slice(i));
    if (m && parseInt(m[4], 10) > 0 && colIndex(m[2]) < 16384) {
      push('ref', null, start, i + m[0].length, { sheet, cAbs: !!m[1], c: colIndex(m[2]), rAbs: !!m[3], r: parseInt(m[4], 10) - 1, sheetText: sheet !== null ? src.slice(start, i) : '' });
      i += m[0].length; continue;
    }
    if (sheet !== null) throw new SyntaxError('Expected a cell after the sheet name');
    m = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(src.slice(i));
    if (m) {
      const word = m[0]; const after = src.slice(i + word.length).trimStart();
      if (after[0] === '(') push('fn', word.toUpperCase(), i, i + word.length);
      else if (/^(TRUE|FALSE)$/i.test(word)) push('bool', word.toUpperCase() === 'TRUE', i, i + word.length);
      else push('name', word, i, i + word.length);
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
    if (t.type === 'op' && (t.val === '-' || t.val === '+')) { const e = unaryOperand(); return t.val === '-' ? { k: 'neg', e } : e; }
    if (t.type === '(') { const e = expr(0); const c = next(); if (!c || c.type !== ')') throw new SyntaxError('A bracket is missing: every ( needs a )'); return postfix(e); }
    if (t.type === 'ref') {
      let node = { k: 'ref', sheet: t.sheet, r: t.r, c: t.c };
      if (peek() && peek().type === 'op' && peek().val === ':' && toks[p + 1] && toks[p + 1].type === 'ref') { next(); const u = next(); node = { k: 'range', sheet: t.sheet, r1: Math.min(t.r, u.r), c1: Math.min(t.c, u.c), r2: Math.max(t.r, u.r), c2: Math.max(t.c, u.c) }; }
      return postfix(node);
    }
    if (t.type === 'name') return postfix({ k: 'name', name: t.val });
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
    const t = toks[i]; if (t.type !== 'ref') continue;
    const r = t.rAbs ? t.r : t.r + dr, c = t.cAbs ? t.c : t.c + dc;
    out += f.slice(last, t.s);
    if (r < 0 || c < 0) { out += '#REF!'; }
    else out += t.sheetText + (t.cAbs ? '$' : '') + colName(c) + (t.rAbs ? '$' : '') + (r + 1);
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
  const txt = t.sheetText + (ca ? '$' : '') + colName(t.c) + (ra ? '$' : '') + (t.r + 1);
  return { text: f.slice(0, t.s) + txt + f.slice(t.e), caret: t.s + txt.length };
}
export function refsIn(f) {
  try { return tokenize(f).filter(t => t.type === 'ref').map((t, i, arr) => t); } catch (e) { return []; }
}

// ---------- workbook ----------
export class Workbook {
  constructor(spec) {
    this.sheets = []; this.names = {}; this.listeners = new Set(); this._cache = null; this.circular = false;
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
  invalidate() { this._cache = new Map(); this.circular = false; this.emit(); }
  onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  emit() { for (const fn of this.listeners) fn(); }
  sheetIndexByName(n) { if (n === null || n === undefined) return -1; const k = n.toLowerCase(); return this.sheets.findIndex(s => s.name.toLowerCase() === k); }
  defineName(name, ref) { this.names[name.toLowerCase()] = { name, ref }; this.invalidate(); }

  // computed value of a cell
  value(si, r, c) {
    const s = this.sheets[si]; if (!s) return ERR.REF; const cell = this.cellObj(s, r, c);
    if (!cell) return null; if (!cell.f) return cell.v;
    if (!this._cache) this._cache = new Map();
    const key = si + ':' + (r * 16384 + c);
    if (this._cache.has(key)) { const v = this._cache.get(key); if (v === IN_PROGRESS) { this.circular = true; return 0; } return v; }
    this._cache.set(key, IN_PROGRESS);
    let v; try { const ast = cell._ast && cell._astSrc === cell.f ? cell._ast : (cell._ast = parse(cell.f), cell._astSrc = cell.f, cell._ast); v = this.evalNode(ast, si); if (v && v.range) v = v.single !== undefined ? v.single : ERR.VALUE; if (isBlank(v)) v = 0; }
    catch (e) { v = e instanceof SyntaxError ? ERR.NAME : (e instanceof XErr ? e : ERR.VALUE); cell._parseError = e instanceof SyntaxError ? e.message : null; }
    this._cache.set(key, v); return v;
  }
  // evaluate a formula in the context of a sheet without storing it
  evaluate(f, si = 0) { try { let v = this.evalNode(parse(f), si); if (v && v.range) v = v.single !== undefined ? v.single : ERR.VALUE; return isBlank(v) ? 0 : v; } catch (e) { return e instanceof SyntaxError ? ERR.NAME : ERR.VALUE; } }
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
      case 'range': { const i = this.resolveSheet(n.sheet, si); return this.rangeVal(i, n); }
      case 'name': {
        const d = this.names[n.name.toLowerCase()]; if (!d) return ERR.NAME;
        const m = /^(?:'?(.+?)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(d.ref.replace(/^=/, ''));
        if (m) { const i = this.resolveSheet(m[1] ?? null, si); const g = parseRange(m[2].replace(/\$/g, '')); return this.rangeVal(i, g); }
        const v = this.evaluate(d.ref[0] === '=' ? d.ref : '=' + d.ref, si); return v;
      }
      case 'neg': { const v = num(scalar(this.evalNode(n.e, si))); return isErr(v) ? v : -v; }
      case 'pct': { const v = num(scalar(this.evalNode(n.e, si))); return isErr(v) ? v : v / 100; }
      case 'bin': return binop(n.op, scalar(this.evalNode(n.a, si)), scalar(this.evalNode(n.b, si)));
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
    return { sheets: this.sheets.map(s => { const cells = {}, styles = {}; for (const [k, cell] of s.cells) { const a = addr(Math.floor(k / 16384), k % 16384); if (cell.f) cells[a] = cell.f; else if (!isBlank(cell.v)) cells[a] = isErr(cell.v) ? cell.v.code : cell.v; if (cell.s && Object.keys(cell.s).length) styles[a] = cell.s; } const widths = {}; for (const [c, w] of Object.entries(s.widths)) widths[colName(+c)] = w; return { name: s.name, cells, styles, widths, freeze: s.freeze, rows: s.rows, cols: s.cols }; }), names: Object.fromEntries(Object.values(this.names).map(d => [d.name, d.ref])) };
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
export const FUNCTION_NAMES = Object.keys(FN).sort();

// ---------- number formats ----------
export function isDateFmt(f) { if (!f) return false; const g = f.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''); return /[dy]/i.test(g) || /(^|[^0#?])m{3,}/i.test(g) || /h/i.test(g) && /:/.test(g); }
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
  if (secs.length > 1 && v < 0) { sec = secs[1]; val = -v; } else if (secs.length > 2 && v === 0) sec = secs[2];
  let out;
  if (isDateFmt(sec)) out = fmtDate(val, sec); else out = fmtNumber(val, sec, secs.length > 1 && v < 0);
  if (width && out.length * 7.2 > width - 6) return '#'.repeat(Math.max(1, Math.floor((width - 6) / 7.2)));
  return out;
}
function splitFmt(f) { const out = []; let cur = '', q = false; for (const ch of f) { if (ch === '"') q = !q; if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out; }
function fmtNumber(v, sec, negSection) {
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
  const tokens = sec.replace(/\[[^\]]*\]/g, '').match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|./gi) || [];
  let out = '', lastWasH = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i], tl = t.toLowerCase();
    if (t[0] === '"') { out += t.slice(1, -1); continue; } if (t[0] === '\\') { out += t[1]; continue; }
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
  const fmtV = v => v && v.range ? (v.single !== undefined ? show(v.single) : `${v.vals.length}×${v.vals[0].length} cells`) : show(v);
  const show = v => isBlank(v) ? '(blank → 0)' : isErr(v) ? v.code : typeof v === 'string' ? '"' + v + '"' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : generalNum(v);
  const text = n => {
    switch (n.k) {
      case 'num': return generalNum(n.v); case 'str': return '"' + n.v + '"'; case 'bool': return n.v ? 'TRUE' : 'FALSE'; case 'err': return n.v.code;
      case 'ref': return (n.sheet ? n.sheet + '!' : '') + addr(n.r, n.c); case 'range': return (n.sheet ? n.sheet + '!' : '') + addr(n.r1, n.c1) + ':' + addr(n.r2, n.c2);
      case 'name': return n.name; case 'neg': return '-' + text(n.e); case 'pct': return text(n.e) + '%';
      case 'bin': return text(n.a) + n.op + text(n.b); case 'call': return n.fn + '(' + n.args.map(text).join(',') + ')'; case 'missing': return '';
    }
  };
  const walk = n => {
    if (n.k === 'ref' || n.k === 'name') { const v = wb.evalNode(n, si); steps.push({ label: n.k === 'name' ? 'Name' : 'Cell', text: text(n), value: fmtV(v) }); return; }
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
    out += f.slice(last, t.s) + (broken ? '#REF!' : t.sheetText + (t.cAbs ? '$' : '') + colName(c) + (t.rAbs ? '$' : '') + (r + 1));
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
Workbook.prototype.isEmpty = function (si, r, c) { const o = this.cellObj(this.sheets[si], r, c); return !o || (isBlank(o.v) && !o.f); };
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
