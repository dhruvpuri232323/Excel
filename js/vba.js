// VBA in the browser: a lexer, parser and interpreter for the VBA that P6 teaches, with Excel's own messages.
// Modules are parsed into procedures; Run executes a Sub as a generator that yields before every statement (so the
// lab can step with F8, stop at breakpoints and show Locals) and yields { ui } for MsgBox, InputBox and forms.
// The Excel objects (Range, Worksheets, Workbooks…) come from a host object: see vbaxl.js.
import { dateToSerial, serialToDate, formatValue } from './engine.js';

// ---------------------------------------------------------------- values
export const EMPTY = { vb: 'Empty' }, NULLV = { vb: 'Null' }, NOTHING = { vb: 'Nothing' }, MISSING = { vb: 'Missing' };
export class VDate { constructor(n) { this.n = n; } }
export class VbaError extends Error { constructor(num, desc, source) { super(desc); this.num = num; this.desc = desc; this.source = source || 'VBAProject'; } }
export class CompileError extends Error { constructor(msg, line, mod, word) { super(msg); this.line = line; this.mod = mod; this.word = word; } }
class EndSignal { }
const ERRS = { 3: 'Return without GoSub', 5: 'Invalid procedure call or argument', 6: 'Overflow', 9: 'Subscript out of range', 10: 'This array is fixed or temporarily locked', 11: 'Division by zero', 13: 'Type mismatch', 20: 'Resume without error', 28: 'Out of stack space',
  52: 'Bad file name or number', 53: 'File not found', 58: 'File already exists', 75: 'Path/File access error', 76: 'Path not found', 91: 'Object variable or With block variable not set', 92: 'For loop not initialized', 94: 'Invalid use of Null',
  424: 'Object required', 429: 'ActiveX component can\'t create object', 438: 'Object doesn\'t support this property or method', 449: 'Argument not optional', 450: 'Wrong number of arguments or invalid property assignment', 457: 'This key is already associated with an element of this collection', 1004: 'Application-defined or object-defined error' };
export const vbErr = (n, d) => new VbaError(n, d || ERRS[n] || 'Application-defined or object-defined error');
export const isObj = v => !!(v && typeof v === 'object' && v.vbObj);
export class VObject { get vbObj() { return true; } get tn() { return 'Object'; } get(n) { throw vbErr(438); } put(n) { throw vbErr(438); } }
export class VArray {
  constructor(bounds, type = 'variant', data) { this.b = bounds.map(x => x.slice()); this.type = type; const n = this.size; this.d = data || Array.from({ length: n }, () => defaultOf(type)); }
  get size() { return this.b.reduce((p, [lo, hi]) => p * Math.max(0, hi - lo + 1), this.b.length ? 1 : 0); }
  idx(ix) { if (ix.length !== this.b.length) throw vbErr(9); let k = 0; for (let i = 0; i < ix.length; i++) { const [lo, hi] = this.b[i]; const v = Math.round(toNum(ix[i])); if (v < lo || v > hi) throw vbErr(9); k = k * (hi - lo + 1) + (v - lo); } return k; }
  get1(ix) { return this.d[this.idx(ix)]; } set1(ix, v) { this.d[this.idx(ix)] = v; }
  clone() { return new VArray(this.b, this.type, this.d.map(x => x instanceof VArray ? x.clone() : x)); }
}
// VBA's and Excel's enum types (Dim answer As VbMsgBoxResult) hold Longs
const ENUMS = new Set(['vbmsgboxresult', 'vbmsgboxstyle', 'vbvartype', 'vbcomparemethod', 'vbdayofweek', 'vbtristate', 'vbfileattribute', 'xldirection', 'xlcalculation', 'xlfileformat', 'xlsortorder', 'xlyesnoguess', 'xlpastetype', 'xlcelltype', 'xlbordersindex', 'xllinestyle', 'xlhalign', 'xlvalign', 'xlfixedformattype', 'xldeleteshiftdirection', 'xlinsertshiftdirection', 'xllookat', 'xlsearchorder', 'xlwindowstate']);
const enumAsLong = t => ENUMS.has(t) ? 'long' : t;
export function defaultOf(t) { t = enumAsLong((t || 'variant').toLowerCase()); if (['integer', 'long', 'longlong', 'single', 'double', 'currency', 'byte', 'decimal'].includes(t)) return 0; if (t === 'string') return ''; if (t === 'boolean') return false; if (t === 'date') return new VDate(0); if (t === 'variant') return EMPTY; return NOTHING; }
export const isNumType = t => ['integer', ...ENUMS, 'long', 'longlong', 'single', 'double', 'currency', 'byte', 'decimal'].includes(t);
const SCALAR = new Set(['integer', 'long', 'longlong', 'single', 'double', 'currency', 'byte', 'decimal', 'string', 'boolean', 'date', 'variant']);
const bankRound = x => { const f = Math.floor(x), d = x - f; if (Math.abs(d - 0.5) < 1e-12) return f % 2 === 0 ? f : f + 1; return Math.round(x); };
export function deref(v, rt) { if (isObj(v) && v.dflt) return deref(v.dflt([], rt), rt); return v; }
export function toNum(v, rt) {
  if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? -1 : 0; if (v === EMPTY || v === undefined || v === null || v === MISSING) return 0; if (v instanceof VDate) return v.n; if (v === NULLV) throw vbErr(94);
  if (typeof v === 'string') { const s = v.trim(); if (/^&h[0-9a-f]+$/i.test(s)) return parseInt(s.slice(2), 16); const t = s.replace(/,/g, ''); if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return Number(t); if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(t)) return Number(t.slice(0, -1)) / 100; throw vbErr(13); }
  if (isObj(v)) { if (v.dflt) return toNum(v.dflt([], rt), rt); throw vbErr(438); } if (v instanceof VArray) throw vbErr(13); if (v === NOTHING) throw vbErr(91); throw vbErr(13); }
export function numStr(n) { if (Number.isInteger(n)) return Math.abs(n) >= 1e15 ? n.toExponential().replace('e+', 'E+').replace(/^(-?\d)E/, '$1E') : String(n); let s = Number(n.toPrecision(15)).toString(); if (/e/.test(s)) s = s.replace('e', 'E'); return s; }
const two = n => String(n).padStart(2, '0');
export function dateStr(n) { const day = Math.floor(n), fr = n - day; const t = () => { const sec = Math.round(fr * 86400); return `${two(Math.floor(sec / 3600))}:${two(Math.floor(sec / 60) % 60)}:${two(sec % 60)}`; }; if (day === 0 && fr > 0) return t(); const d = serialToDate(day); const ds = `${two(d.d)}-${two(d.m)}-${d.y}`; return fr > 1e-9 ? ds + ' ' + t() : ds; }
export function toStr(v, rt) {
  if (typeof v === 'string') return v; if (typeof v === 'number') return numStr(v); if (typeof v === 'boolean') return v ? 'True' : 'False'; if (v === EMPTY || v === undefined || v === null || v === MISSING) return ''; if (v instanceof VDate) return dateStr(v.n);
  if (v === NULLV) throw vbErr(94); if (isObj(v)) { if (v.dflt) return toStr(v.dflt([], rt), rt); throw vbErr(438); } if (v === NOTHING) throw vbErr(91); if (v && v.code) return 'Error ' + ({ '#NULL!': 2000, '#DIV/0!': 2007, '#VALUE!': 2015, '#REF!': 2023, '#NAME?': 2029, '#NUM!': 2036, '#N/A': 2042 }[v.code] || 2015); throw vbErr(13); }
export function toBool(v, rt) { if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0; if (v === EMPTY || v === undefined) return false; if (typeof v === 'string') { const s = v.trim().toLowerCase(); if (s === 'true') return true; if (s === 'false') return false; return toNum(v) !== 0; } if (v instanceof VDate) return v.n !== 0; if (v === NULLV) throw vbErr(94); if (isObj(v) && v.dflt) return toBool(v.dflt([], rt), rt); throw vbErr(13); }
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
export function parseDate(s) { // CDate on an Indian (en-IN) computer: day before month
  s = String(s).trim(); let m; let t = 0; const tm = /\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(s); if (tm) { let h = +tm[1]; if (tm[4]) { if (/pm/i.test(tm[4]) && h < 12) h += 12; if (/am/i.test(tm[4]) && h === 12) h = 0; } t = (h * 3600 + +tm[2] * 60 + +(tm[3] || 0)) / 86400; s = s.slice(0, tm.index); }
  else if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(s))) { let h = +m[1]; if (m[4]) { if (/pm/i.test(m[4]) && h < 12) h += 12; if (/am/i.test(m[4]) && h === 12) h = 0; } return (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400; }
  const ok = (y, mo, d) => mo >= 1 && mo <= 12 && d >= 1 && d <= 31 && serialToDate(dateToSerial(y, mo, d)).m === mo ? dateToSerial(y, mo, d) + t : null; const yy = y => y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y;
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s))) return ok(+m[1], +m[2], +m[3]);
  if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(s))) return ok(yy(+m[3]), +m[2], +m[1]);
  const mi = w => MONTHS.findIndex(x => x.startsWith(w.toLowerCase()) && w.length >= 3) + 1;
  if ((m = /^(\d{1,2})[\s-]+([A-Za-z]+)[\s-,]+(\d{2,4})$/.exec(s)) && mi(m[2])) return ok(yy(+m[3]), mi(m[2]), +m[1]);
  if ((m = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{2,4})$/.exec(s)) && mi(m[1])) return ok(yy(+m[3]), mi(m[1]), +m[2]);
  if ((m = /^([A-Za-z]+)[\s-]+(\d{4})$/.exec(s)) && mi(m[1])) return ok(+m[2], mi(m[1]), 1);
  return null; }
export function toDate(v) { if (v instanceof VDate) return v; if (typeof v === 'number') return new VDate(v); if (typeof v === 'string') { const n = parseDate(v); if (n === null) { if (/^[+-]?\d+(\.\d+)?$/.test(v.trim())) return new VDate(+v); throw vbErr(13); } return new VDate(n); } if (v === EMPTY) return new VDate(0); if (typeof v === 'boolean') return new VDate(v ? -1 : 0); if (isObj(v) && v.dflt) return toDate(v.dflt([])); throw vbErr(13); }
export function coerce(type, v, rt) {
  type = enumAsLong((type || 'variant').toLowerCase()); if (v instanceof VArray) { if (type === 'variant') return v.clone(); throw vbErr(13); }
  if (!SCALAR.has(type)) { if (v === NOTHING || isObj(v)) return v; throw vbErr(424); }
  if (isObj(v)) v = deref(v, rt); if (v === NOTHING) throw vbErr(91);
  switch (type) {
    case 'variant': return v === undefined ? EMPTY : v;
    case 'string': return toStr(v, rt);
    case 'boolean': return toBool(v, rt);
    case 'date': return toDate(v);
    case 'integer': { const n = bankRound(toNum(v, rt)); if (n < -32768 || n > 32767) throw vbErr(6); return n; }
    case 'long': { const n = bankRound(toNum(v, rt)); if (n < -2147483648 || n > 2147483647) throw vbErr(6); return n; }
    case 'byte': { const n = bankRound(toNum(v, rt)); if (n < 0 || n > 255) throw vbErr(6); return n; }
    case 'currency': return Math.round(toNum(v, rt) * 10000) / 10000;
    case 'single': return Number(Math.fround(toNum(v, rt)).toPrecision(7));
    default: return toNum(v, rt); } }
export function typeName(v) { if (v === EMPTY || v === undefined) return 'Empty'; if (v === NULLV) return 'Null'; if (v === NOTHING) return 'Nothing'; if (typeof v === 'string') return 'String'; if (typeof v === 'boolean') return 'Boolean'; if (v instanceof VDate) return 'Date'; if (typeof v === 'number') return Number.isInteger(v) ? (Math.abs(v) <= 32767 ? 'Integer' : 'Long') : 'Double'; if (v instanceof VArray) return (v.type === 'variant' ? 'Variant' : v.type[0].toUpperCase() + v.type.slice(1)) + '()'; if (isObj(v)) return v.tn; if (v && v.code) return 'Error'; return 'Object'; }
const TYPE_NAME = { integer: 'Integer', long: 'Long', longlong: 'LongLong', single: 'Single', double: 'Double', currency: 'Currency', byte: 'Byte', string: 'String', boolean: 'Boolean', date: 'Date', variant: 'Variant', decimal: 'Decimal' };
export const typeLabel = t => TYPE_NAME[(t || 'variant').toLowerCase()] || (t[0].toUpperCase() + t.slice(1));

// ---------------------------------------------------------------- lexer
const OPS = ['<=', '>=', '<>', ':=', '=', '<', '>', '+', '-', '*', '/', '\\', '^', '&', '(', ')', ',', '.', ';', '!', ':'];
function lex(src, mod) {
  const T = []; let i = 0, line = 1, sp = true; const n = src.length;
  const err = m => { throw new CompileError(m, line, mod); };
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\t') { i++; sp = true; continue; }
    if (ch === '_' && sp) { let j = i + 1; while (src[j] === ' ' || src[j] === '\t') j++; if (src[j] === '\r' || src[j] === '\n' || j >= n) { i = j; if (src[i] === '\r') i++; if (src[i] === '\n') i++; line++; sp = true; continue; } }
    if (ch === '\r') { i++; continue; }
    if (ch === '\n') { T.push({ t: 'nl', line }); line++; i++; sp = true; continue; }
    if (ch === "'") { while (i < n && src[i] !== '\n') i++; continue; }
    if (ch === '"') { let s = ''; i++; for (;;) { if (i >= n || src[i] === '\n') err('Expected: end of statement'); if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; } s += src[i++]; } T.push({ t: 'str', v: s, line, sp }); sp = false; continue; }
    if (ch === '#') { const m = /^#([^#\n]+)#/.exec(src.slice(i, i + 40)); if (m && !(T.length && T[T.length - 1].t === 'id' && !sp)) { let d = m[1].trim(); const mm = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(.*)$/.exec(d); const v = mm ? parseDate(`${mm[3]}-${mm[1]}-${mm[2]}${mm[4]}`) : parseDate(d); if (v === null) err('Syntax error'); T.push({ t: 'date', v, line, sp }); i += m[0].length; sp = false; continue; } }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] || '') && sp)) { const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?[#!@%&^]?/.exec(src.slice(i)); T.push({ t: 'num', v: parseFloat(m[1] + (m[2] || '')), line, sp }); i += m[0].length; sp = false; continue; }
    if (ch === '&' && /[hH]/.test(src[i + 1] || '') && /[0-9a-fA-F]/.test(src[i + 2] || '')) { const m = /^&[hH]([0-9a-fA-F]+)&?/.exec(src.slice(i)); T.push({ t: 'num', v: parseInt(m[1], 16), line, sp }); i += m[0].length; sp = false; continue; }
    if (/[A-Za-z_\u00C0-\uFFFF]/.test(ch)) { const m = /^[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_\u00C0-\uFFFF]*/.exec(src.slice(i)); let w = m[0]; i += w.length; if ('$%&!#@'.includes(src[i] || ' ') && !/[A-Za-z0-9_]/.test(src[i + 1] || ' ') && src[i] !== '!' ) i++;
      if (w.toLowerCase() === 'rem' && (T.length === 0 || ['nl', 'op'].includes(T[T.length - 1].t) && (T[T.length - 1].t === 'nl' || T[T.length - 1].v === ':'))) { while (i < n && src[i] !== '\n') i++; continue; }
      T.push({ t: 'id', v: w, lc: w.toLowerCase(), line, sp }); sp = false; continue; }
    const op = OPS.find(o => src.startsWith(o, i)); if (!op) err('Invalid character'); T.push({ t: 'op', v: op, line, sp }); i += op.length; sp = false;
  }
  T.push({ t: 'nl', line }); T.push({ t: 'eof', line }); return T;
}

// ---------------------------------------------------------------- parser
const RESERVED = new Set(['and', 'as', 'byref', 'byval', 'call', 'case', 'const', 'dim', 'do', 'each', 'else', 'elseif', 'end', 'exit', 'false', 'for', 'function', 'goto', 'if', 'in', 'is', 'like', 'loop', 'mod', 'new', 'next', 'not', 'nothing', 'on', 'optional', 'or', 'paramarray', 'preserve', 'private', 'public', 'redim', 'resume', 'select', 'set', 'static', 'step', 'sub', 'then', 'to', 'true', 'until', 'wend', 'while', 'with', 'xor', 'let', 'global', 'eqv', 'imp', 'empty', 'null', 'me']);
class Parser {
  constructor(src, mod) { this.T = lex(src, mod); this.p = 0; this.mod = mod; }
  get tok() { return this.T[this.p]; } peek(k = 1) { return this.T[this.p + k]; }
  err(m, t = this.tok, word) { throw new CompileError(m, t.line, this.mod, word); }
  isW(w, t = this.tok) { return t.t === 'id' && t.lc === w; } isOp(o, t = this.tok) { return t.t === 'op' && t.v === o; }
  eatW(w) { if (this.isW(w)) { this.p++; return true; } return false; } eatOp(o) { if (this.isOp(o)) { this.p++; return true; } return false; }
  needW(w, m) { if (!this.eatW(w)) this.err(m || `Expected: ${w[0].toUpperCase() + w.slice(1)}`); } needOp(o, m) { if (!this.eatOp(o)) this.err(m || `Expected: ${o}`); }
  eos() { return this.tok.t === 'nl' || this.tok.t === 'eof' || this.isOp(':') || this.isW('else') && this.single; }
  endStmt() { if (this.isOp(':')) { this.p++; return; } if (this.tok.t === 'nl') { if (!this.single) this.p++; return; } if (this.tok.t === 'eof') return; if (this.isW('else') && this.single) return; this.err('Expected: end of statement'); }
  skipNl() { while (this.tok.t === 'nl' || this.isOp(':')) this.p++; }
  name(m = 'Expected: identifier') { const t = this.tok; if (t.t !== 'id') this.err(m); this.p++; return t.v; }
  // ---- module
  module() {
    const M = { options: {}, decls: [], procs: [], events: [] }; this.skipNl();
    while (this.tok.t !== 'eof') {
      const t = this.tok;
      if (this.isW('option')) { this.p++; const w = this.name(); const lw = w.toLowerCase(); if (lw === 'explicit') M.options.explicit = true; else if (lw === 'base') { M.options.base = this.tok.v; this.p++; } else if (lw === 'compare') { M.options.compare = this.name().toLowerCase(); } else if (lw === 'private') { this.name(); } this.endStmt(); this.skipNl(); continue; }
      let scope = null; if (['public', 'private', 'global', 'dim', 'static', 'friend'].includes(t.lc) && t.t === 'id') { scope = t.lc; this.p++; }
      if (this.isW('sub') || this.isW('function') || this.isW('property')) { M.procs.push(this.proc(scope || 'public')); this.skipNl(); continue; }
      if (this.isW('const')) { this.p++; for (const c of this.constList()) M.decls.push({ k: 'const', ...c, scope: scope || 'private', line: t.line }); this.endStmt(); this.skipNl(); continue; }
      if (scope) { if (this.isW('withevents')) this.p++; for (const v of this.varList()) M.decls.push({ k: 'dim', ...v, scope, line: t.line }); this.endStmt(); this.skipNl(); continue; }
      if (this.isW('declare') || this.isW('type') || this.isW('enum')) this.err('This kind of declaration isn’t part of this course.');
      this.err('Invalid outside procedure');
    }
    return M;
  }
  proc(scope) {
    const start = this.tok; let kind = this.name().toLowerCase(); if (kind === 'property') { this.err('Property procedures aren’t part of this course.'); }
    const name = this.name('Expected: identifier'); const params = []; if (this.eatOp('(')) { if (!this.eatOp(')')) { do { params.push(this.param()); } while (this.eatOp(',')); this.needOp(')', 'Expected: list separator or )'); } }
    let type = null; if (kind === 'function' && this.eatW('as')) type = this.typeName();
    this.endStmt(); const body = this.block(['end ' + kind]); const endTok = this.tok; this.p += 2; this.endStmt();
    return { kind, name, lname: name.toLowerCase(), params, type: type || 'variant', body, scope, line: start.line, endLine: endTok.line, mod: this.mod };
  }
  param() { const pr = { byval: false, optional: false, def: null, type: 'variant', array: false }; for (;;) { if (this.eatW('optional')) pr.optional = true; else if (this.eatW('byval')) pr.byval = true; else if (this.eatW('byref')) pr.byval = false; else if (this.eatW('paramarray')) pr.paramarray = true; else break; }
    pr.name = this.name(); if (this.eatOp('(')) { this.needOp(')'); pr.array = true; } if (this.eatW('as')) pr.type = this.typeName(); if (this.eatOp('=')) pr.def = this.expr(); return pr; }
  typeName() { let t = this.name('Expected: identifier'); while (this.eatOp('.')) t = this.name(); if (this.isOp('*')) { this.p++; this.p++; } return enumAsLong(t.toLowerCase()); }
  varList() { const out = []; do { const line = this.tok.line; const nameT = this.tok; const name = this.name(); const v = { name, line, dims: null, type: 'variant', isNew: false };
      if (RESERVED.has(nameT.lc)) this.err('Expected: identifier', nameT);
      if (this.eatOp('(')) { v.dims = []; if (!this.eatOp(')')) { do { const a = this.expr(); if (this.eatW('to')) v.dims.push([a, this.expr()]); else v.dims.push([null, a]); } while (this.eatOp(',')); this.needOp(')', 'Expected: list separator or )'); } }
      if (this.eatW('as')) { if (this.eatW('new')) v.isNew = true; v.type = this.typeName(); } out.push(v); } while (this.eatOp(',')); return out; }
  constList() { const out = []; do { const name = this.name(); let type = null; if (this.eatW('as')) type = this.typeName(); this.needOp('='); out.push({ name, type, expr: this.expr() }); } while (this.eatOp(',')); return out; }
  // ---- statements: parse until one of the terminators (e.g. 'end if', 'else', 'next'), which is left unconsumed
  atTerm(terms) { for (const t of terms) { const w = t.split(' '); if (w.every((x, k) => this.isW(x, this.peek(k)))) return t; } return null; }
  block(terms) { const out = []; const saved = this.terms; this.terms = terms; try { for (;;) { this.skipNl(); if (this.tok.t === 'eof') return this.missing(terms); if (this.atTerm(terms)) return out; const s = this.stmt(); if (s) out.push(s); } } finally { this.terms = saved; } }
  missing(terms, at) { const t = terms.includes('end if') ? 'end if' : terms.includes('end select') ? 'end select' : terms[0]; const m = { 'end if': 'Block If without End If', next: 'For without Next', loop: 'Do without Loop', wend: 'While without Wend', 'end select': 'Select Case without End Select', 'end with': 'Expected End With', 'end sub': 'Expected End Sub', 'end function': 'Expected End Function', case: 'Select Case without End Select' }[t] || 'Syntax error'; this.err(m, at || this.T[this.T.length - 1]); }
  stmt() {
    const t = this.tok; const line = t.line; const w = t.t === 'id' ? t.lc : null;
    const S = o => ({ ...o, line });
    // stray block ends
    if (w === 'end' && this.peek().t === 'id' && ['sub', 'function'].includes(this.peek().lc)) this.missing(this.terms || ['end sub'], t);
    if (w === 'end' && this.peek().t === 'id' && ['if', 'select', 'with'].includes(this.peek().lc)) this.err({ if: 'End If without block If', select: 'End Select without Select Case', with: 'End With without With' }[this.peek().lc]);
    if (w === 'next') this.err('Next without For'); if (w === 'loop') this.err('Loop without Do'); if (w === 'wend') this.err('Wend without While'); if (w === 'else' || w === 'elseif') this.err('Else without If'); if (w === 'case') this.err('Case without Select Case');
    if (w === 'sub' || w === 'function') this.err('Expected End Sub');
    // a label: Name: at the start of a line
    if (t.t === 'id' && this.isOp(':', this.peek()) && (this.p === 0 || this.T[this.p - 1].t === 'nl') && !RESERVED.has(w)) { this.p += 2; return S({ k: 'label', name: t.v, lname: w }); }
    if (w === 'dim' || w === 'static' || (w === 'private' || w === 'public')) { this.p++; if (this.isW('const')) { this.p++; const cs = this.constList(); this.endStmt(); return S({ k: 'consts', list: cs }); } const vars = this.varList(); this.endStmt(); return S({ k: 'dim', vars }); }
    if (w === 'const') { this.p++; const cs = this.constList(); this.endStmt(); return S({ k: 'consts', list: cs }); }
    if (w === 'redim') { this.p++; const preserve = this.eatW('preserve'); const vars = []; do { const name = this.name(); this.needOp('('); const dims = []; do { const a = this.expr(); if (this.eatW('to')) dims.push([a, this.expr()]); else dims.push([null, a]); } while (this.eatOp(',')); this.needOp(')'); let type = null; if (this.eatW('as')) type = this.typeName(); vars.push({ name, dims, type }); } while (this.eatOp(',')); this.endStmt(); return S({ k: 'redim', preserve, vars }); }
    if (w === 'erase') { this.p++; const names = []; do names.push(this.name()); while (this.eatOp(',')); this.endStmt(); return S({ k: 'erase', names }); }
    if (w === 'if') return this.ifStmt();
    if (w === 'select') { this.p++; this.needW('case'); const e = this.expr(); this.endStmt(); const cases = []; let els = null;
      this.skipNl(); if (!this.isW('case') && !this.isW('end')) this.err('Statements and labels invalid between Select Case and first Case');
      while (this.isW('case')) { this.p++; if (this.eatW('else')) { this.eatOp(':'); els = this.block(['case', 'end select']); continue; } const tests = []; do { if (this.eatW('is')) { const op = this.tok.v; this.p++; tests.push({ k: 'is', op, e: this.expr() }); } else { const a = this.expr(); if (this.eatW('to')) tests.push({ k: 'range', a, b: this.expr() }); else tests.push({ k: 'eq', e: a }); } } while (this.eatOp(','));
        if (this.isOp(':')) this.p++; cases.push({ tests, body: this.block(['case', 'end select']) }); }
      this.needW('end'); this.needW('select'); this.endStmt(); return S({ k: 'select', e, cases, els }); }
    if (w === 'for') { this.p++; if (this.eatW('each')) { const v = this.postfix(); this.needW('in'); const coll = this.expr(); this.endStmt(); const body = this.block(['next']); this.p++; if (this.tok.t === 'id' && !this.eos()) this.p++; this.endStmt(); return S({ k: 'foreach', v, coll, body }); }
      const v = this.postfix(); this.needOp('=', 'Expected: ='); const from = this.expr(); this.needW('to'); const to = this.expr(); let step = null; if (this.eatW('step')) step = this.expr(); this.endStmt(); const body = this.block(['next']); this.p++;
      if (this.tok.t === 'id' && !this.eos()) { const nv = this.tok; this.p++; if (v.k === 'id' && nv.lc !== v.name.toLowerCase()) this.err('Invalid Next control variable reference', nv); while (this.eatOp(',')) this.name(); } this.endStmt(); return S({ k: 'for', v, from, to, step, body }); }
    if (w === 'do') { this.p++; let pre = null; if (this.isW('while') || this.isW('until')) { const until = this.tok.lc === 'until'; this.p++; pre = { until, e: this.expr() }; } this.endStmt(); const body = this.block(['loop']); this.p++; let post = null; if (this.isW('while') || this.isW('until')) { const until = this.tok.lc === 'until'; this.p++; post = { until, e: this.expr() }; } this.endStmt(); return S({ k: 'do', pre, post, body }); }
    if (w === 'while') { this.p++; const e = this.expr(); this.endStmt(); const body = this.block(['wend']); this.p++; this.endStmt(); return S({ k: 'do', pre: { until: false, e }, post: null, body }); }
    if (w === 'with') { this.p++; const e = this.expr(); this.endStmt(); const body = this.block(['end with']); this.p += 2; this.endStmt(); return S({ k: 'with', e, body }); }
    if (w === 'exit') { this.p++; const what = this.name().toLowerCase(); if (!['for', 'do', 'sub', 'function'].includes(what)) this.err('Expected: Do or For or Sub or Function or Property'); this.endStmt(); return S({ k: 'exit', what }); }
    if (w === 'on' && this.isW('error', this.peek())) { this.p += 2; let s; if (this.eatW('resume')) { this.needW('next'); s = { k: 'onerror', mode: 'next' }; } else { this.needW('goto'); if (this.tok.t === 'num' && this.tok.v === 0) { this.p++; s = { k: 'onerror', mode: 'zero' }; } else if (this.isOp('-') && this.peek().v === 1) { this.p += 2; s = { k: 'onerror', mode: 'zero' }; } else s = { k: 'onerror', mode: 'goto', label: this.name().toLowerCase() }; } this.endStmt(); return S(s); }
    if (w === 'resume') { this.p++; let s = { k: 'resume', mode: 'here' }; if (this.eatW('next')) s = { k: 'resume', mode: 'next' }; else if (this.tok.t === 'id' && !this.eos()) s = { k: 'resume', mode: 'label', label: this.name().toLowerCase() }; else if (this.tok.t === 'num') { this.p++; } this.endStmt(); return S(s); }
    if (w === 'goto') { this.p++; const label = this.name().toLowerCase(); this.endStmt(); return S({ k: 'goto', label }); }
    if (w === 'end') { const nx = this.peek(); if (nx.t === 'nl' || nx.t === 'eof' || this.isOp(':', nx)) { this.p++; this.endStmt(); return S({ k: 'end' }); } }
    if (w === 'stop') { this.p++; this.endStmt(); return S({ k: 'stop' }); }
    if (w === 'set') { this.p++; const target = this.postfix(); this.needOp('=', 'Expected: ='); if (this.eatW('new')) { const cls = this.typeName(); this.endStmt(); return S({ k: 'assign', set: true, target, e: { k: 'new', cls } }); } const e = this.expr(); this.endStmt(); return S({ k: 'assign', set: true, target, e }); }
    if (w === 'let') { this.p++; }
    if (w === 'call') { this.p++; const e = this.postfix(); this.endStmt(); return S({ k: 'call', e: e.k === 'call' ? e : { k: 'call', fn: e, args: [] }, explicitCall: true }); }
    if (w === 'debug' && this.isOp('.', this.peek()) && this.isW('print', this.peek(2))) { this.p += 3; const items = []; while (!this.eos()) { if (this.isOp(';') || this.isOp(',')) { items.push({ sep: this.tok.v }); this.p++; continue; } items.push({ e: this.expr() }); } this.endStmt(); return S({ k: 'print', items }); }
    if (w === 'randomize' || w === 'beep' || w === 'doevents') { this.p++; if (!this.eos()) this.expr(); this.endStmt(); return null; }
    if (w === 'mid' && this.isOp('(', this.peek())) { /* Mid(s, i, n) = "x" */ const save = this.p; const target = this.postfix(); if (this.eatOp('=')) { const e = this.expr(); this.endStmt(); return S({ k: 'midassign', target, e }); } this.p = save; }
    if (t.t !== 'id' && !this.isOp('.')) this.err(t.t === 'eof' ? 'Syntax error' : 'Expected: statement');
    if (t.t === 'id' && RESERVED.has(w) && !['me', 'not'].includes(w)) this.err('Syntax error');
    // assignment or a call: target expression first, stopping at a spaced "(" (Foo (a), b passes (a) as an argument)
    const target = this.postfix(true);
    if (this.eatOp('=')) { const e = this.expr(); this.endStmt(); return S({ k: 'assign', set: false, target, e }); }
    if (this.eos()) { if (target.k === 'call' && target.args.length > 1 && !target.byParenSpace) this.err('Expected: ='); if (!['id', 'member', 'call'].includes(target.k)) this.err('Expected: statement'); this.endStmt(); return S({ k: 'call', e: target.k === 'call' ? target : { k: 'call', fn: target, args: [] } }); }
    const args = this.argList(true); this.endStmt();
    if (target.k === 'call' && !target.byParenSpace) this.err('Expected: end of statement');
    return S({ k: 'call', e: { k: 'call', fn: target, args } });
  }
  ifStmt() {
    const line = this.tok.line; this.p++; const cond = this.expr(); if (!this.eatW('then')) { if (this.isW('goto')) { this.p++; const label = this.name().toLowerCase(); this.endStmt(); return { k: 'if', line, branches: [{ cond, body: [{ k: 'goto', label, line }] }], els: null }; } this.err('Expected: Then or GoTo'); }
    if (!(this.tok.t === 'nl' || this.tok.t === 'eof')) { // single-line If
      const save = this.single; this.single = true; const body = []; const els = [];
      while (!(this.tok.t === 'nl' || this.tok.t === 'eof') && !this.isW('else')) { const s = this.stmt(); if (s) body.push(s); }
      if (this.eatW('else')) { while (!(this.tok.t === 'nl' || this.tok.t === 'eof')) { const s = this.stmt(); if (s) els.push(s); } }
      this.single = save; if (this.tok.t === 'nl' && !save) this.p++; return { k: 'if', line, branches: [{ cond, body }], els: els.length ? els : null }; } // a single-line If inside another leaves the end of the line to the outer one
    this.endStmt(); const branches = [{ cond, body: this.block(['elseif', 'else', 'end if']) }]; let els = null;
    for (;;) { if (this.eatW('elseif')) { const c = this.expr(); this.needW('then', 'Expected: Then or GoTo'); this.endStmt(); branches.push({ cond: c, body: this.block(['elseif', 'else', 'end if']) }); continue; }
      if (this.eatW('else')) { this.eatOp(':'); els = this.block(['end if']); }
      this.needW('end'); this.needW('if'); this.endStmt(); return { k: 'if', line, branches, els }; } }
  argList(stmtCtx) { const args = []; if (stmtCtx && this.eos()) return args; for (;;) { if (this.isOp(',')) { args.push({ e: null }); this.p++; continue; } if (this.isOp(')') && !stmtCtx) { break; }
      if (this.tok.t === 'id' && this.isOp(':=', this.peek())) { const nm = this.tok.lc; this.p += 2; args.push({ name: nm, e: this.expr() }); } else args.push({ e: this.expr() });
      if (!this.eatOp(',')) break; if ((stmtCtx && this.eos()) || (!stmtCtx && this.isOp(')'))) { args.push({ e: null }); break; } }
    return args; }
  // ---- expressions (VBA precedence: ^, unary -, * /, \, Mod, + -, &, comparisons, Not, And, Or, Xor, Eqv, Imp)
  expr() { return this.imp(); }
  imp() { let a = this.eqv(); while (this.isW('imp')) { this.p++; a = { k: 'bin', op: 'imp', a, b: this.eqv() }; } return a; }
  eqv() { let a = this.xor(); while (this.isW('eqv')) { this.p++; a = { k: 'bin', op: 'eqv', a, b: this.xor() }; } return a; }
  xor() { let a = this.or(); while (this.isW('xor')) { this.p++; a = { k: 'bin', op: 'xor', a, b: this.or() }; } return a; }
  or() { let a = this.and(); while (this.isW('or')) { this.p++; a = { k: 'bin', op: 'or', a, b: this.and() }; } return a; }
  and() { let a = this.not(); while (this.isW('and')) { this.p++; a = { k: 'bin', op: 'and', a, b: this.not() }; } return a; }
  not() { if (this.isW('not')) { this.p++; return { k: 'un', op: 'not', a: this.not() }; } return this.cmp(); }
  cmp() { let a = this.cat(); for (;;) { const t = this.tok; if (t.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) { this.p++; a = { k: 'bin', op: t.v, a, b: this.cat() }; } else if (this.isW('like')) { this.p++; a = { k: 'bin', op: 'like', a, b: this.cat() }; } else if (this.isW('is')) { this.p++; a = { k: 'bin', op: 'is', a, b: this.cat() }; } else return a; } }
  cat() { let a = this.add(); while (this.isOp('&')) { this.p++; a = { k: 'bin', op: '&', a, b: this.add() }; } return a; }
  add() { let a = this.modOp(); while (this.isOp('+') || this.isOp('-')) { const op = this.tok.v; this.p++; a = { k: 'bin', op, a, b: this.modOp() }; } return a; }
  modOp() { let a = this.idiv(); while (this.isW('mod')) { this.p++; a = { k: 'bin', op: 'mod', a, b: this.idiv() }; } return a; }
  idiv() { let a = this.mul(); while (this.isOp('\\')) { this.p++; a = { k: 'bin', op: '\\', a, b: this.mul() }; } return a; }
  mul() { let a = this.neg(); while (this.isOp('*') || this.isOp('/')) { const op = this.tok.v; this.p++; a = { k: 'bin', op, a, b: this.neg() }; } return a; }
  neg() { if (this.isOp('-')) { this.p++; return { k: 'un', op: '-', a: this.neg() }; } if (this.isOp('+')) { this.p++; return this.neg(); } return this.pow(); }
  pow() { let a = this.postfix(); while (this.isOp('^')) { this.p++; const b = this.isOp('-') ? (this.p++, { k: 'un', op: '-', a: this.postfix() }) : this.postfix(); a = { k: 'bin', op: '^', a, b }; } return a; }
  postfix(stmt) { let e = this.primary();
    for (;;) { if (this.isOp('.') && stmt && this.tok.sp && this.p > 0) return e; // c.Add .Cells(r, 1): a spaced dot in a statement starts a With-member argument
      if (this.isOp('.')) { this.p++; const nm = this.tok; if (nm.t !== 'id') this.err('Expected: identifier'); this.p++; e = { k: 'member', obj: e, name: nm.v, lname: nm.lc }; continue; }
      if (this.isOp('!')) { this.p++; const nm = this.name(); e = { k: 'call', fn: e, args: [{ e: { k: 'str', v: nm } }] }; continue; }
      if (this.isOp('(')) { if (stmt && this.tok.sp && e.k !== 'call' && this.p > 0) { return { ...e, stopAtParen: true }; } this.p++; const args = this.argList(false); this.needOp(')', 'Expected: list separator or )'); e = { k: 'call', fn: e, args }; continue; }
      return e; } }
  primary() { const t = this.tok;
    if (t.t === 'num') { this.p++; return { k: 'lit', v: t.v }; } if (t.t === 'str') { this.p++; return { k: 'lit', v: t.v }; } if (t.t === 'date') { this.p++; return { k: 'lit', v: new VDate(t.v) }; }
    if (this.isOp('(')) { this.p++; const e = this.expr(); this.needOp(')', 'Expected: )'); return { k: 'paren', e }; }
    if (this.isOp('.')) { this.p++; const nm = this.tok; if (nm.t !== 'id') this.err('Expected: identifier'); this.p++; return { k: 'member', obj: null, name: nm.v, lname: nm.lc }; }
    if (t.t === 'id') { if (t.lc === 'true') { this.p++; return { k: 'lit', v: true }; } if (t.lc === 'false') { this.p++; return { k: 'lit', v: false }; } if (t.lc === 'nothing') { this.p++; return { k: 'lit', v: NOTHING }; } if (t.lc === 'empty') { this.p++; return { k: 'lit', v: EMPTY }; } if (t.lc === 'null') { this.p++; return { k: 'lit', v: NULLV }; }
      if (t.lc === 'new') { this.p++; return { k: 'new', cls: this.typeName() }; } if (t.lc === 'me') { this.p++; return { k: 'me' }; }
      if (RESERVED.has(t.lc) && !['not'].includes(t.lc)) this.err('Expected: expression'); this.p++; return { k: 'id', name: t.v, lname: t.lc, line: t.line }; }
    this.err('Expected: expression'); }
}
// a statement call with a spaced paren: Foo (a), b → args [(a), b]
const _postfix = Parser.prototype.postfix;
Parser.prototype.postfix = function (stmt) { const e = _postfix.call(this, stmt); if (e && e.stopAtParen) { delete e.stopAtParen; const args = this.argList(true); return { k: 'call', fn: e, args, byParenSpace: true }; } return e; };
export function parseModule(code, mod = 'Module1') { return new Parser(code, mod).module(); }

// ---------------------------------------------------------------- built-in functions
const D = n => new VDate(n);
const ymdOf = v => serialToDate(Math.floor(toDate(v).n));
const MONTHNAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYNAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const addMonths = (n, k) => { const d = serialToDate(Math.floor(n)); const t = d.m - 1 + k; const y = d.y + Math.floor(t / 12), m = ((t % 12) + 12) % 12 + 1; const last = serialToDate(dateToSerial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1).d; return dateToSerial(y, m, Math.min(d.d, last)) + (n - Math.floor(n)); };
function likeRe(p, text) { let r = ''; for (let i = 0; i < p.length; i++) { const c = p[i]; if (c === '*') r += '[\\s\\S]*'; else if (c === '?') r += '[\\s\\S]'; else if (c === '#') r += '\\d'; else if (c === '[') { const j = p.indexOf(']', i); let cls = p.slice(i + 1, j); if (cls[0] === '!') cls = '^' + cls.slice(1); r += '[' + cls.replace(/\\/g, '\\\\') + ']'; i = j; } else r += c.replace(/[.+^${}()|\\]/g, '\\$&'); } return new RegExp('^' + r + '$', text ? 'i' : ''); }
export function vbaFormat(v, f) {
  if (f === undefined || f === MISSING || f === '') return toStr(v);
  const nf = String(f).toLowerCase(); const num = () => v instanceof VDate ? v.n : toNum(v);
  if (nf === 'general number') return numStr(num()); if (nf === 'fixed') return num().toFixed(2); if (nf === 'standard') return formatValue(num(), '#,##0.00'); if (nf === 'percent') return (num() * 100).toFixed(2) + '%'; if (nf === 'currency') return formatValue(num(), '₹#,##0.00');
  if (nf === 'yes/no') return toBool(v) ? 'Yes' : 'No'; if (nf === 'true/false') return toBool(v) ? 'True' : 'False'; if (nf === 'short date') return dateStr(Math.floor(toDate(v).n)); if (nf === 'long date') { const d = ymdOf(v); return `${d.d} ${MONTHNAMES[d.m - 1]} ${d.y}`; } if (nf === 'short time') { const n = toDate(v).n; const s = Math.round((n - Math.floor(n)) * 86400); return `${two(Math.floor(s / 3600))}:${two(Math.floor(s / 60) % 60)}`; } if (nf === 'long time') return dateStr(toDate(v).n - Math.floor(toDate(v).n) || 1e-9);
  if (f === '>' ) return toStr(v).toUpperCase(); if (f === '<') return toStr(v).toLowerCase();
  let x; if (v instanceof VDate) x = v.n; else if (typeof v === 'number') x = v; else if (typeof v === 'boolean') x = v ? -1 : 0; else { const s = toStr(v); const d = /[dmyhns]/i.test(f) && !/[0#]/.test(f) ? parseDate(s) : null; if (d !== null) x = d; else { try { x = toNum(s); } catch (e) { return s; } } }
  return formatValue(x, String(f).replace(/n/g, 'm').replace(/N/g, 'M'));
}
// each: (args, rt) → value; args are evaluated values (MISSING for an omitted one)
const has = a => a !== undefined && a !== MISSING;
export const BI = {
  len: ([s]) => s === NULLV ? NULLV : toStr(s).length, left: ([s, n]) => toStr(s).slice(0, Math.max(0, toNum(n))), right: ([s, n]) => { const t = toStr(s); const k = toNum(n); return k <= 0 ? '' : t.slice(-k); },
  mid: ([s, a, n]) => { const t = toStr(s); const st = toNum(a); if (st < 1) throw vbErr(5); return has(n) ? t.substr(st - 1, Math.max(0, toNum(n))) : t.slice(st - 1); },
  instr: (a, rt) => { let st = 1, s1, s2, cmp = rt && rt.compareText ? 1 : 0; if (a.length >= 3 || (a.length === 2 && typeof a[0] === 'number' && false)) { [st, s1, s2] = [toNum(a[0]), a[1], a[2]]; if (has(a[3])) cmp = toNum(a[3]); } else [s1, s2] = a; if (s1 === NULLV || s2 === NULLV) return NULLV; let x = toStr(s1), y = toStr(s2); if (cmp === 1) { x = x.toLowerCase(); y = y.toLowerCase(); } if (st < 1) throw vbErr(5); if (y === '') return st; const i = x.indexOf(y, st - 1); return i + 1; },
  instrrev: ([s1, s2, st, cmp]) => { let x = toStr(s1), y = toStr(s2); if (has(cmp) && toNum(cmp) === 1) { x = x.toLowerCase(); y = y.toLowerCase(); } const from = has(st) && toNum(st) > 0 ? toNum(st) - 1 : x.length; return x.lastIndexOf(y, from - y.length + 1 >= 0 ? from - y.length + 1 : 0) + 1 || 0; },
  replace: ([s, f, r, st, cnt, cmp]) => { let t = toStr(s); const a = toStr(f), b = toStr(r); if (a === '') return t; const start = has(st) ? toNum(st) : 1; t = t.slice(start - 1); let n = has(cnt) ? toNum(cnt) : -1; const ci = has(cmp) && toNum(cmp) === 1; let out = '', i = 0; for (;;) { if (n === 0) break; const j = ci ? t.toLowerCase().indexOf(a.toLowerCase(), i) : t.indexOf(a, i); if (j < 0) break; out += t.slice(i, j) + b; i = j + a.length; n--; } return out + t.slice(i); },
  ucase: ([s]) => toStr(s).toUpperCase(), lcase: ([s]) => toStr(s).toLowerCase(), trim: ([s]) => toStr(s).replace(/^ +| +$/g, ''), ltrim: ([s]) => toStr(s).replace(/^ +/, ''), rtrim: ([s]) => toStr(s).replace(/ +$/, ''),
  space: ([n]) => ' '.repeat(Math.max(0, toNum(n))), string: ([n, c]) => (typeof c === 'number' ? String.fromCharCode(c) : toStr(c)[0] || '').repeat(Math.max(0, toNum(n))), strreverse: ([s]) => [...toStr(s)].reverse().join(''),
  strcomp: ([a, b, c]) => { let x = toStr(a), y = toStr(b); if (has(c) && toNum(c) === 1) { x = x.toLowerCase(); y = y.toLowerCase(); } return x < y ? -1 : x > y ? 1 : 0; },
  strconv: ([s, m]) => { const k = toNum(m); const t = toStr(s); if (k === 1) return t.toUpperCase(); if (k === 2) return t.toLowerCase(); if (k === 3) return t.toLowerCase().replace(/(^|[\s])(\S)/g, (x, a, b) => a + b.toUpperCase()); return t; },
  chr: ([n]) => String.fromCharCode(toNum(n)), asc: ([s]) => { const t = toStr(s); if (!t) throw vbErr(5); return t.charCodeAt(0); }, chrw: ([n]) => String.fromCharCode(toNum(n)), ascw: ([s]) => toStr(s).charCodeAt(0),
  split: ([s, d, lim]) => { const t = toStr(s); const dl = has(d) ? toStr(d) : ' '; let parts = t === '' ? [] : dl === '' ? [t] : t.split(dl); if (has(lim) && toNum(lim) > 0 && parts.length > toNum(lim)) { const k = toNum(lim); parts = [...parts.slice(0, k - 1), parts.slice(k - 1).join(dl)]; } return new VArray([[0, parts.length - 1]], 'string', parts); },
  join: ([a, d]) => { if (!(a instanceof VArray)) throw vbErr(13); return a.d.map(x => toStr(x)).join(has(d) ? toStr(d) : ' '); },
  array: (args) => new VArray([[0, args.length - 1]], 'variant', args.map(x => x === MISSING ? EMPTY : x)),
  lbound: ([a, d]) => { if (!(a instanceof VArray)) throw vbErr(13); const k = has(d) ? toNum(d) : 1; if (!a.b[k - 1]) throw vbErr(9); return a.b[k - 1][0]; },
  ubound: ([a, d]) => { if (!(a instanceof VArray)) throw vbErr(13); const k = has(d) ? toNum(d) : 1; if (!a.b[k - 1]) throw vbErr(9); return a.b[k - 1][1]; },
  isarray: ([a]) => a instanceof VArray, isempty: ([a]) => a === EMPTY || a === undefined, isnull: ([a]) => a === NULLV, isobject: ([a]) => isObj(a) || a === NOTHING, ismissing: ([a]) => a === MISSING || a === undefined,
  isnumeric: ([a]) => { if (typeof a === 'number' || typeof a === 'boolean' || a === EMPTY) return true; if (isObj(a) && a.dflt) return BI.isnumeric([a.dflt([])]); if (typeof a !== 'string') return false; try { toNum(a); return a.trim() !== ''; } catch (e) { return false; } },
  isdate: ([a]) => { if (a instanceof VDate) return true; if (isObj(a) && a.dflt) return BI.isdate([a.dflt([])]); return typeof a === 'string' && parseDate(a) !== null; }, iserror: ([a]) => !!(a && a.code), cverr: ([n]) => ({ code: { 2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!', 2029: '#NAME?', 2036: '#NUM!', 2042: '#N/A' }[toNum(n)] || '#VALUE!', cv: toNum(n) }),
  cstr: ([a], rt) => toStr(a, rt), cint: ([a], rt) => coerce('integer', a, rt), clng: ([a], rt) => coerce('long', a, rt), cdbl: ([a], rt) => toNum(deref(a, rt), rt), csng: ([a], rt) => coerce('single', a, rt), ccur: ([a], rt) => coerce('currency', a, rt), cbool: ([a], rt) => toBool(a, rt), cdate: ([a], rt) => toDate(deref(a, rt)), cvar: ([a], rt) => deref(a, rt), cbyte: ([a], rt) => coerce('byte', a, rt), clnglng: ([a], rt) => coerce('long', a, rt), cdec: ([a], rt) => toNum(deref(a, rt), rt),
  val: ([s]) => { const m = /^\s*[+-]?(\d[\d\s]*\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(toStr(s)); return m ? Number(m[0].replace(/\s/g, '')) : 0; }, str: ([n]) => { const x = toNum(n); return (x >= 0 ? ' ' : '') + numStr(x); },
  int: ([n]) => Math.floor(toNum(n)), fix: ([n]) => Math.trunc(toNum(n)), abs: ([n]) => Math.abs(toNum(n)), sgn: ([n]) => Math.sign(toNum(n)), sqr: ([n]) => { const x = toNum(n); if (x < 0) throw vbErr(5); return Math.sqrt(x); }, exp: ([n]) => Math.exp(toNum(n)), log: ([n]) => { const x = toNum(n); if (x <= 0) throw vbErr(5); return Math.log(x); },
  round: ([n, d]) => { const k = has(d) ? toNum(d) : 0; const f = 10 ** k; return bankRound(Number((toNum(n) * f).toPrecision(15))) / f; }, rnd: (a, rt) => rt ? rt.rand() : Math.random(), atn: ([n]) => Math.atan(toNum(n)), sin: ([n]) => Math.sin(toNum(n)), cos: ([n]) => Math.cos(toNum(n)), tan: ([n]) => Math.tan(toNum(n)),
  hex: ([n]) => Math.round(toNum(n)).toString(16).toUpperCase(), oct: ([n]) => Math.round(toNum(n)).toString(8),
  iif: ([c, a, b]) => toBool(c) ? a : b, choose: ([i, ...xs]) => { const k = Math.trunc(toNum(i)); return k >= 1 && k <= xs.length ? xs[k - 1] : NULLV; }, switch: (xs) => { for (let i = 0; i + 1 < xs.length; i += 2) if (toBool(xs[i])) return xs[i + 1]; return NULLV; },
  now: (a, rt) => D(rt.now()), date: (a, rt) => D(Math.floor(rt.now())), time: (a, rt) => D(rt.now() - Math.floor(rt.now())), timer: (a, rt) => Math.round(((rt.now() - Math.floor(rt.now())) * 86400 + rt.clock) * 100) / 100,
  dateserial: ([y, m, d]) => { let yy = Math.trunc(toNum(y)), mm = Math.trunc(toNum(m)); yy += Math.floor((mm - 1) / 12); mm = ((mm - 1) % 12 + 12) % 12 + 1; return D(dateToSerial(yy, mm, 1) + Math.trunc(toNum(d)) - 1); },
  timeserial: ([h, m, s]) => D((toNum(h) * 3600 + toNum(m) * 60 + toNum(s)) / 86400), datevalue: ([s]) => D(Math.floor(toDate(s).n)), timevalue: ([s]) => { const n = toDate(s).n; return D(n - Math.floor(n)); },
  year: ([d]) => ymdOf(d).y, month: ([d]) => ymdOf(d).m, day: ([d]) => ymdOf(d).d, weekday: ([d, f]) => { const w = ((Math.floor(toDate(d).n) + 6) % 7 + 7) % 7; const first = has(f) ? toNum(f) : 1; return ((w - (first - 1)) % 7 + 7) % 7 + 1; },
  hour: ([d]) => Math.floor(Math.round((toDate(d).n % 1) * 86400) / 3600), minute: ([d]) => Math.floor(Math.round((toDate(d).n % 1) * 86400) / 60) % 60, second: ([d]) => Math.round((toDate(d).n % 1) * 86400) % 60,
  monthname: ([m, ab]) => { const k = toNum(m); if (k < 1 || k > 12) throw vbErr(5); const s = MONTHNAMES[k - 1]; return has(ab) && toBool(ab) ? s.slice(0, 3) : s; }, weekdayname: ([w, ab, f]) => { const first = has(f) ? toNum(f) : 1; const s = DAYNAMES[(toNum(w) - 1 + first - 1) % 7]; return has(ab) && toBool(ab) ? s.slice(0, 3) : s; },
  dateadd: ([iv, n, d]) => { const k = Math.trunc(toNum(n)), x = toDate(d).n, i = toStr(iv).toLowerCase(); const m = { yyyy: 12, q: 3, m: 1 }[i]; if (m) return D(addMonths(x, k * m)); const dd = { d: 1, y: 1, w: 1, ww: 7 }[i]; if (dd) return D(x + k * dd); const s = { h: 3600, n: 60, s: 1 }[i]; if (s) return D(x + k * s / 86400); throw vbErr(5); },
  datediff: ([iv, a, b]) => { const x = toDate(a).n, y = toDate(b).n, i = toStr(iv).toLowerCase(); const p = serialToDate(Math.floor(x)), q = serialToDate(Math.floor(y)); if (i === 'yyyy') return q.y - p.y; if (i === 'q') return (q.y - p.y) * 4 + Math.floor((q.m - 1) / 3) - Math.floor((p.m - 1) / 3); if (i === 'm') return (q.y - p.y) * 12 + q.m - p.m; if (i === 'd' || i === 'y') return Math.floor(y) - Math.floor(x); if (i === 'w') return Math.trunc((Math.floor(y) - Math.floor(x)) / 7); if (i === 'ww') return Math.floor((Math.floor(y) - ((Math.floor(y) + 6) % 7)) / 7) - Math.floor((Math.floor(x) - ((Math.floor(x) + 6) % 7)) / 7); const s = { h: 3600, n: 60, s: 1 }[i]; if (s) return Math.round((y - x) * 86400 / s); throw vbErr(5); },
  datepart: ([iv, d]) => { const i = toStr(iv).toLowerCase(); const x = ymdOf(d); if (i === 'yyyy') return x.y; if (i === 'm') return x.m; if (i === 'd') return x.d; if (i === 'q') return Math.floor((x.m - 1) / 3) + 1; if (i === 'w') return BI.weekday([d]); if (i === 'y') return Math.floor(toDate(d).n) - dateToSerial(x.y, 1, 1) + 1; return BI[{ h: 'hour', n: 'minute', s: 'second' }[i]]([d]); },
  format: ([v, f]) => vbaFormat(v, f), formatnumber: ([v, d]) => formatValue(toNum(v), '#,##0' + (has(d) && toNum(d) > 0 ? '.' + '0'.repeat(toNum(d)) : (has(d) ? '' : '.00'))), formatpercent: ([v, d]) => (toNum(v) * 100).toFixed(has(d) ? toNum(d) : 2) + '%', formatcurrency: ([v]) => formatValue(toNum(v), '₹#,##0.00'), formatdatetime: ([v]) => dateStr(toDate(v).n),
  rgb: ([r, g, b]) => (toNum(r) & 255) + (toNum(g) & 255) * 256 + (toNum(b) & 255) * 65536, qbcolor: ([n]) => [0, 8388608, 32768, 8421376, 128, 8388736, 32896, 12632256, 8421504, 16711680, 65280, 16776960, 255, 16711935, 65535, 16777215][toNum(n)] || 0,
  typename: ([a]) => typeName(a), vartype: ([a]) => { const t = typeName(a); return { Empty: 0, Null: 1, Integer: 2, Long: 3, Single: 4, Double: 5, Currency: 6, Date: 7, String: 8, Boolean: 11, Nothing: 9, Error: 10 }[t] ?? (a instanceof VArray ? 8204 : 9); },
  environ: ([s]) => ({ username: 'you', userprofile: 'C:\\Users\\you', computername: 'DHANUSH-PC', temp: 'C:\\Users\\you\\AppData\\Local\\Temp' }[toStr(s).toLowerCase()] || ''),
  doevents: () => 0, createobject: ([c], rt) => rt.newObj(toStr(c), true), nz: ([a, b]) => a === NULLV || a === EMPTY ? (has(b) ? b : 0) : a,
};
// MsgBox constants and the rest of the vb… constants
export const CONSTS = { vbcrlf: '\r\n', vbnewline: '\r\n', vbcr: '\r', vblf: '\n', vbtab: '\t', vbnullstring: '', vbnullchar: '\0', vbback: '\b',
  vbformcontrolmenu: 0, vbformcode: 1, vbmodal: 1, vbmodeless: 0,
  vbokonly: 0, vbokcancel: 1, vbabortretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5, vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64, vbdefaultbutton1: 0, vbdefaultbutton2: 256, vbdefaultbutton3: 512, vbapplicationmodal: 0, vbsystemmodal: 4096,
  vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5, vbyes: 6, vbno: 7,
  vbblack: 0, vbred: 255, vbgreen: 65280, vbyellow: 65535, vbblue: 16711680, vbmagenta: 16711935, vbcyan: 16776960, vbwhite: 16777215,
  vbsunday: 1, vbmonday: 2, vbtuesday: 3, vbwednesday: 4, vbthursday: 5, vbfriday: 6, vbsaturday: 7, vbbinarycompare: 0, vbtextcompare: 1, vbuppercase: 1, vblowercase: 2, vbpropercase: 3,
  vbnormal: 0, vbreadonly: 1, vbhidden: 2, vbdirectory: 16, vbarchive: 32, vbempty: 0, vbnull: 1, vbinteger: 2, vblong: 3, vbsingle: 4, vbdouble: 5, vbcurrency: 6, vbdate: 7, vbstring: 8, vbobject: 9, vberror: 10, vbboolean: 11, vbvariant: 12, vbarray: 8192, vbobjecterror: -2147221504, vbtrue: -1, vbfalse: 0, vbusedefault: -2 };

// ---------------------------------------------------------------- built-in objects: Collection, Dictionary, Err, Debug
export class VCollection extends VObject {
  constructor() { super(); this.items = []; } get tn() { return 'Collection'; }
  find(k) { if (typeof k === 'string') { const i = this.items.findIndex(x => x.key !== undefined && x.key.toLowerCase() === k.toLowerCase()); if (i < 0) throw vbErr(5); return i; } const n = Math.round(toNum(k)); if (n < 1 || n > this.items.length) throw vbErr(9); return n - 1; }
  get(n, a, rt) { const [x, y, z, w] = bind(a, n === 'add' ? ['item', 'key', 'before', 'after'] : n === 'remove' ? ['index'] : ['index']);
    if (n === 'count') return this.items.length; if (n === 'item') { if (!has(x)) throw vbErr(449); return this.items[this.find(deref(x, rt))].v; }
    if (n === 'add') { const key = has(y) ? toStr(y) : undefined; if (key !== undefined && this.items.some(i => i.key !== undefined && i.key.toLowerCase() === key.toLowerCase())) throw vbErr(457); const it = { v: x, key }; if (has(z)) this.items.splice(this.find(z), 0, it); else if (has(w)) this.items.splice(this.find(w) + 1, 0, it); else this.items.push(it); return EMPTY; }
    if (n === 'remove') { this.items.splice(this.find(deref(x, rt)), 1); return EMPTY; } throw vbErr(438); }
  dflt(a, rt) { if (!a.length) throw vbErr(450); return this.get('item', a, rt); } each() { return this.items.map(i => i.v); }
}
export class VDictionary extends VObject {
  constructor() { super(); this.m = new Map(); this.textCompare = false; } get tn() { return 'Dictionary'; }
  k(key) { key = deref(key); if (key instanceof VDate) return 'd' + key.n; if (typeof key === 'number') return 'n' + key; if (typeof key === 'string') return 's' + (this.textCompare ? key.toLowerCase() : key); if (typeof key === 'boolean') return 'b' + key; if (isObj(key)) { if (!this._ids) this._ids = new WeakMap(); if (!this._ids.has(key)) this._ids.set(key, 'o' + Math.random()); return this._ids.get(key); } return 'e'; }
  get(n, a, rt) { const [x, y] = bind(a, n === 'add' ? ['key', 'item'] : ['key']);
    if (n === 'count') return this.m.size; if (n === 'exists') return this.m.has(this.k(x)); if (n === 'keys') return new VArray([[0, this.m.size - 1]], 'variant', [...this.m.values()].map(e => e.key)); if (n === 'items') return new VArray([[0, this.m.size - 1]], 'variant', [...this.m.values()].map(e => e.v));
    if (n === 'add') { const kk = this.k(x); if (this.m.has(kk)) throw vbErr(457); this.m.set(kk, { key: deref(x), v: y }); return EMPTY; } if (n === 'item') { const kk = this.k(x); if (!this.m.has(kk)) this.m.set(kk, { key: deref(x), v: EMPTY }); return this.m.get(kk).v; }
    if (n === 'remove') { const kk = this.k(x); if (!this.m.has(kk)) throw vbErr(32811, 'Element not found'); this.m.delete(kk); return EMPTY; } if (n === 'removeall') { this.m.clear(); return EMPTY; } if (n === 'comparemode') return this.textCompare ? 1 : 0; throw vbErr(438); }
  put(n, a, v) { if (n === 'item') { const kk = this.k(a[0].v); const e = this.m.get(kk); if (e) e.v = v; else this.m.set(kk, { key: deref(a[0].v), v }); return; } if (n === 'comparemode') { this.textCompare = toNum(v) === 1; return; } if (n === 'key') { const e = this.m.get(this.k(a[0].v)); if (!e) throw vbErr(32811, 'Element not found'); this.m.delete(this.k(a[0].v)); this.m.set(this.k(v), { key: v, v: e.v }); return; } throw vbErr(438); }
  dflt(a, rt) { return this.get('item', a, rt); } putDflt(a, v) { this.put('item', a, v); } each() { return [...this.m.values()].map(e => e.key); }
}
class ErrObject extends VObject {
  constructor() { super(); this.clear(); } get tn() { return 'ErrObject'; } clear() { this.number = 0; this.description = ''; this.source = ''; }
  set(e) { this.number = e.num; this.description = e.desc; this.source = e.source || 'VBAProject'; }
  get(n, a) { if (n === 'number') return this.number; if (n === 'description') return this.description; if (n === 'source') return this.source; if (n === 'clear') { this.clear(); return EMPTY; } if (n === 'raise') { const [num, src, desc] = bind(a, ['number', 'source', 'description']); throw new VbaError(toNum(num), has(desc) ? toStr(desc) : (ERRS[toNum(num)] || 'Application-defined or object-defined error'), has(src) ? toStr(src) : undefined); } if (n === 'helpcontext' || n === 'lastdllerror') return 0; throw vbErr(438); }
  put(n, a, v) { if (n === 'number') this.number = toNum(v); else if (n === 'description') this.description = toStr(v); else if (n === 'source') this.source = toStr(v); else throw vbErr(438); }
  dflt() { return this.number; }
}
// bind [{name, v}] arguments to parameter names, positionally then by name
export function bind(args, names) { const out = names.map(() => MISSING); let i = 0; for (const a of args || []) { if (a.name) { const k = names.indexOf(a.name); if (k < 0) throw vbErr(448, 'Named argument not found'); out[k] = a.v; } else { if (i < out.length) out[i] = a.v; i++; } } return out; }

// ---------------------------------------------------------------- the program
const BUILTIN_SUBS = new Set(['msgbox', 'inputbox', 'createobject', 'getobject', 'dir', 'kill', 'filecopy', 'mkdir', 'rmdir', 'name', 'chdir', 'filelen', 'filedatetime', 'unload', 'load', 'sendkeys', 'shell', 'callbyname', 'curdir']);
export class Program {
  // modules: [{ name, kind: 'std' | 'sheet' | 'workbook' | 'form', code, sheet? }]; host: the Excel side (vbaxl.js)
  constructor(modules, host = {}) {
    this.host = host; this.mods = new Map(); this.procs = new Map(); this.errors = []; this.out = []; this.clock = 0; this.seed = 1; this.err = new ErrObject();
    for (const m of modules) { let ast; try { ast = parseModule(m.code || '', m.name); } catch (e) { if (e instanceof CompileError) { this.errors.push(e); ast = { options: {}, decls: [], procs: [] }; } else throw e; } this.mods.set(m.name.toLowerCase(), { ...m, ast, vars: new Map() }); }
    for (const M of this.mods.values()) { const seen = new Set(); for (const p of M.ast.procs) { if (seen.has(p.lname)) this.errors.push(new CompileError(`Ambiguous name detected: ${p.name}`, p.line, M.name, p.name)); seen.add(p.lname);
        if (M.kind !== 'std') continue; const prev = this.procs.get(p.lname); if (prev && p.scope !== 'private' && prev.proc.scope !== 'private') this.errors.push(new CompileError(`Ambiguous name detected: ${p.name}`, p.line, M.name, p.name)); if (!prev || prev.proc.scope === 'private') this.procs.set(p.lname, { proc: p, mod: M }); } }
    this.forms = new Map(); for (const M of this.mods.values()) if (M.kind === 'form') this.forms.set(M.name.toLowerCase(), Object.assign(new UserForm(M.form || {}, M), { prog: this })); this.formPending = [];
    if (!this.errors.length) this.checkAll(); this.initModules();
  }
  // ---- forms and events
  formRef(k, f) { if (f.mod.kind === 'form') { const F = this.forms.get(f.mod.name.toLowerCase()); const c = F && F.ctl(k); if (c) return c; } return this.forms ? this.forms.get(k) || null : null; }
  *fire(F, name, args = []) { const P = F.M.ast.procs.find(p => p.lname === name.toLowerCase()); if (!P) return false; yield* this.invoke(P, this.mods.get(F.M.name.toLowerCase()), args.map(v => ({ v })), F); return true; }
  // UserForm_QueryClose(Cancel, CloseMode) before a form closes: 0 = the X (vbFormControlMenu), 1 = Unload in code (vbFormCode). True if the code set Cancel.
  *queryClose(F, mode) { const P = F.M.ast.procs.find(p => p.lname === 'userform_queryclose'); if (!P) return false; if (F.closing) return false; F.closing = true;
    try { const slot = { type: (P.params[0] && P.params[0].type) || 'integer', v: 0 }; yield* this.invoke(P, this.mods.get(F.M.name.toLowerCase()), [{ v: 0, ref: slot }, { v: mode }], F); const c = slot.v; return !(c === 0 || c === false || c === EMPTY || c === ''); } finally { F.closing = false; } }
  *loadForm(F) { if (F.loaded) return; F.load(); yield* this.fire(F, 'UserForm_Initialize'); }
  *showForm(F, modeless) { yield* this.loadForm(F); if (!F.loaded) return EMPTY; F.visible = true; let first = true;
    while (F.visible && F.loaded) { const act = yield { ui: 'form', form: F, first, modeless }; first = false; yield TICK;
      if (!act || act.close) { if (yield* this.queryClose(F, 0)) continue; F.unload(); break; }
      if (act.set) for (const [n, v] of Object.entries(act.set)) { const c = F.ctl(n); if (c && c.enabled && c.setByUser(v)) { yield* this.fire(F, c.name + '_Change'); if (c.lt === 'checkbox' || c.lt === 'optionbutton' || c.lt === 'listbox' || c.lt === 'combobox') yield* this.fire(F, c.name + '_Click'); } }
      if (act.click) { const c = F.ctl(act.click); if (!c || c.enabled) yield* this.fire(F, (c ? c.name : act.click) + '_Click'); }
      if (act.dblclick) { const c = F.ctl(act.dblclick); if (c && c.enabled) yield* this.fire(F, c.name + '_DblClick', [false]); } }
    return EMPTY; }
  // code changed a control's value (txtPacks.Value = ""): its Change (and Click) events run after the statement, as VBA runs them during it
  *flushFormEvents() { const list = this.formPending.splice(0); for (const [F, ev] of list) if (F.loaded) yield* this.fire(F, ev); }
  *flushEvents() { const list = this.host.pending.splice(0); for (const ev of list) { if (this.host.events === false) break; const M = [...this.mods.values()].find(m => m.kind === 'sheet' && ev.sheet && m.sheet && m.sheet.toLowerCase() === ev.sheet.toLowerCase()); const P = M && M.ast.procs.find(p => p.lname === 'worksheet_change'); if (P) yield* this.invoke(P, M, [{ v: ev.target }]); } }
  rand() { this.seed = (this.seed * 1103515245 + 12345) % 2147483648; return this.seed / 2147483648; }
  now() { return this.host.now ? this.host.now() : dateToSerial(2026, 11, 16) + 10 / 24; }
  initModules() { for (const M of this.mods.values()) { M.vars = new Map(); for (const d of M.ast.decls) { if (d.k === 'dim') M.vars.set(d.name.toLowerCase(), this.newSlot(d)); else M.vars.set(d.name.toLowerCase(), { type: d.type || 'variant', v: this.constVal(d.expr, M), isConst: true }); } } this.err.clear(); }
  reset() { this.initModules(); }
  constVal(e, M) { const f = new Frame({ name: '(declarations)', body: [], params: [], lname: '' }, M); const it = this.ev(e, f); let r = it.next(); while (!r.done) r = it.next(); return r.value; }
  newSlot(d, f) { const t = d.type || 'variant'; if (d.dims) { if (!d.dims.length) return { type: t, v: new VArray([], t), dyn: true, arr: true }; const b = d.dims.map(([lo, hi]) => [lo ? toNum(this.constEval(lo, f)) : 0, toNum(this.constEval(hi, f))]); return { type: t, v: new VArray(b, t), arr: true, fixed: true }; }
    if (d.isNew) return { type: t, v: null, autoNew: t }; return { type: t, v: defaultOf(t) }; }
  constEval(e, f) { const it = this.ev(e, f || new Frame({ name: '', body: [], params: [], lname: '' }, [...this.mods.values()][0])); let r = it.next(); while (!r.done) r = it.next(); return r.value; }
  // ---- compile checks: Option Explicit (Variable not defined), unknown Subs, labels, duplicate declarations
  known(M) { const s = new Set([...this.procs.keys(), ...M.ast.procs.map(p => p.lname), ...Object.keys(BI), ...Object.keys(CONSTS), ...BUILTIN_SUBS, 'debug', 'err', 'me', ...(this.host.globalNames ? this.host.globalNames() : []), ...Object.keys(this.host.consts || {}), ...(this.host.fnNames ? this.host.fnNames() : []), ...(this.forms ? this.forms.keys() : []), ...(M.kind === 'form' && this.forms.get(M.name.toLowerCase()) ? this.forms.get(M.name.toLowerCase()).ctls.keys() : [])]); for (const X of this.mods.values()) { for (const d of X.ast.decls) if (X === M || (X.kind === 'std' && d.scope !== 'private' && d.scope !== 'dim')) s.add(d.name.toLowerCase()); if (X.kind !== 'std') s.add(X.name.toLowerCase()); } return s; }
  checkAll() { for (const M of this.mods.values()) { const G = this.known(M); for (const p of M.ast.procs) { try { this.checkProc(p, M, G); } catch (e) { if (e instanceof CompileError) { this.errors.push(e); break; } throw e; } } } }
  checkProc(p, M, G) {
    const L = new Set(p.params.map(x => x.name.toLowerCase())); const labels = new Set(); const declLines = new Map();
    const walk = (ss, fn) => { for (const s of ss || []) { fn(s); for (const k of ['body', 'els']) if (Array.isArray(s[k])) walk(s[k], fn); if (s.branches) for (const b of s.branches) walk(b.body, fn); if (s.cases) for (const c of s.cases) walk(c.body, fn); } };
    walk(p.body, s => { if (s.k === 'dim') for (const v of s.vars) { const k = v.name.toLowerCase(); if (L.has(k)) throw new CompileError('Duplicate declaration in current scope', s.line, M.name, v.name); L.add(k); declLines.set(k, s.line); } if (s.k === 'consts') for (const c of s.list) { const k = c.name.toLowerCase(); if (L.has(k)) throw new CompileError('Duplicate declaration in current scope', s.line, M.name, c.name); L.add(k); } if (s.k === 'label') labels.add(s.lname); });
    if (p.kind === 'function') L.add(p.lname);
    const explicit = !!M.ast.options.explicit; const seen = new Set(L);
    const chkId = (e, line, call) => { const k = e.lname; if (L.has(k) || G.has(k) || seen.has(k)) return; if (call) throw new CompileError('Sub or Function not defined', line, M.name, e.name); if (explicit) throw new CompileError('Variable not defined', line, M.name, e.name); seen.add(k); };
    const ex = (e, line) => { if (!e) return; switch (e.k) { case 'id': chkId(e, line, false); return; case 'member': ex(e.obj, line); return; case 'call': if (e.fn.k === 'id') { const k = e.fn.lname; if (!(L.has(k) || G.has(k) || seen.has(k))) chkId(e.fn, line, true); } else ex(e.fn, line); for (const a of e.args) ex(a.e, line); return; case 'bin': ex(e.a, line); ex(e.b, line); return; case 'un': case 'paren': ex(e.a || e.e, line); return; default: return; } };
    const top = p.body.filter(s => s.k === 'label').map(s => s.lname);
    walk(p.body, s => { const line = s.line;
      switch (s.k) { case 'assign': if (s.target.k === 'id') { if (!L.has(s.target.lname) && !G.has(s.target.lname) && !seen.has(s.target.lname)) { if (explicit) throw new CompileError('Variable not defined', line, M.name, s.target.name); seen.add(s.target.lname); } } else ex(s.target, line); ex(s.e, line); break;
        case 'call': if (s.e.fn.k === 'id' && !s.e.args.length) chkId(s.e.fn, line, true); else ex(s.e, line); break;
        case 'if': for (const b of s.branches) ex(b.cond, line); break; case 'select': ex(s.e, line); for (const c of s.cases) for (const t of c.tests) { ex(t.e || t.a, line); ex(t.b, line); } break;
        case 'for': if (s.v.k === 'id') chkId(s.v, line, false); else ex(s.v, line); ex(s.from, line); ex(s.to, line); ex(s.step, line); break; case 'foreach': if (s.v.k === 'id') chkId(s.v, line, false); ex(s.coll, line); break;
        case 'do': if (s.pre) ex(s.pre.e, line); if (s.post) ex(s.post.e, line); break; case 'with': ex(s.e, line); break; case 'print': for (const it of s.items) ex(it.e, line); break;
        case 'redim': for (const v of s.vars) { if (!L.has(v.name.toLowerCase()) && !G.has(v.name.toLowerCase())) { if (explicit) throw new CompileError('Variable not defined', line, M.name, v.name); L.add(v.name.toLowerCase()); } for (const [a, b] of v.dims) { ex(a, line); ex(b, line); } } break;
        case 'goto': if (!labels.has(s.label)) throw new CompileError('Label not defined', line, M.name, s.label); break; case 'onerror': if (s.mode === 'goto' && !labels.has(s.label)) throw new CompileError('Label not defined', line, M.name, s.label); break; case 'resume': if (s.mode === 'label' && !labels.has(s.label)) throw new CompileError('Label not defined', line, M.name, s.label); break;
        case 'dim': for (const v of s.vars) for (const [a, b] of v.dims || []) { ex(a, line); ex(b, line); } break; case 'consts': for (const c of s.list) ex(c.expr, line); break; default: } });
    p.topLabels = top;
  }
  firstError() { return this.errors[0] || null; }
  macros() { const out = []; for (const M of this.mods.values()) if (M.kind === 'std') for (const p of M.ast.procs) if (p.kind === 'sub' && !p.params.length && p.scope !== 'private') out.push({ name: p.name, mod: M.name }); return out; }
  procNames(mod) { const M = this.mods.get(String(mod).toLowerCase()); return M ? M.ast.procs.map(p => p.name) : []; }
  findProc(name) { return this.procs.get(String(name).toLowerCase()) || null; }
  // run a procedure: returns a generator (see Session)
  *call(name, args = [], opts = {}) { if (this.errors.length) throw this.errors[0]; const P = typeof name === 'object' ? name : this.findProc(name); if (!P) throw vbErr(1004, `Cannot run the macro '${name}'. The macro may not be available in this workbook or all macros may be disabled.`); this.stack = []; this.out = this.out || []; try { return yield* this.invoke(P.proc, P.mod, args.map(v => ({ v })), opts.me); } finally { this.flush(); } }
  flush() { if (this.pend) { this.print(this.pend.replace(/\s+$/, '')); this.pend = ''; } }
  // ---- invoking a user procedure
  *invoke(proc, M, args, me) {
    if ((this.stack || (this.stack = [])).length > 300) throw vbErr(28);
    const f = new Frame(proc, M, me); const pos = args.filter(a => !a.name); const named = args.filter(a => a.name);
    proc.params.forEach((pr, i) => { let a = pos[i]; const nm = named.find(x => x.name === pr.name.toLowerCase()); if (nm) a = nm; const k = pr.name.toLowerCase();
      if (pr.paramarray) { f.vars.set(k, { type: 'variant', v: new VArray([[0, pos.length - i - 1]], 'variant', pos.slice(i).map(x => x.v)) }); return; }
      if (!a || a.v === MISSING) { if (!pr.optional) throw vbErr(449); f.vars.set(k, { type: pr.type, v: pr.def ? coerce(pr.type, this.constEval(pr.def, f), this) : (pr.type === 'variant' ? MISSING : defaultOf(pr.type)) }); return; }
      if (a.ref && !pr.byval) { if (pr.type !== 'variant' && a.ref.type !== pr.type && !(a.ref.arr && pr.array) && SCALAR.has(pr.type)) throw new CompileError('ByRef argument type mismatch', this.line || proc.line, M.name, pr.name); f.vars.set(k, a.ref); return; }
      let v = a.v; if (pr.array) { if (!(v instanceof VArray)) throw vbErr(13); f.vars.set(k, { type: pr.type, v, arr: true }); return; }
      if (isObj(v) && !SCALAR.has(pr.type)) { f.vars.set(k, { type: pr.type, v }); return; } f.vars.set(k, { type: pr.type, v: coerce(pr.type, pr.type === 'variant' && isObj(v) ? v : v, this) }); });
    if (pos.length > proc.params.length && !proc.params.some(p => p.paramarray)) throw vbErr(450);
    // Dim is not executable: every local exists from the start of the procedure
    const hoist = ss => { for (const s of ss || []) { if (s.k === 'dim') for (const v of s.vars) f.vars.set(v.name.toLowerCase(), this.newSlot(v, f)); if (s.k === 'consts') for (const c of s.list) f.vars.set(c.name.toLowerCase(), { type: c.type || 'variant', v: this.constEval(c.expr, f), isConst: true }); for (const k of ['body', 'els']) if (Array.isArray(s[k])) hoist(s[k]); if (s.branches) for (const b of s.branches) hoist(b.body); if (s.cases) for (const c of s.cases) hoist(c.body); } };
    hoist(proc.body); if (proc.kind === 'function') f.vars.set(proc.lname, { type: proc.type, v: defaultOf(proc.type), ret: true });
    this.stack.push(f);
    try {
      yield { line: proc.line, frame: f, enter: true };
      let i = 0; const body = proc.body; const labelAt = l => body.findIndex(s => s.k === 'label' && s.lname === l);
      while (i < body.length) {
        const sig = yield* this.execTop(body, i, f);
        if (!sig) { i++; continue; }
        if (sig.exit === 'sub' || sig.exit === 'function') break; if (sig.end) throw new EndSignal();
        if (sig.gotoErr) { const j = labelAt(sig.gotoErr); f.resumeNext = i + 1; f.resumeHere = i; f.inHandler = true; i = j; continue; }
        if (sig.goto) { const j = labelAt(sig.goto); if (j < 0) throw new CompileError('Label not defined', body[i].line, M.name, sig.goto); i = j; continue; }
        if (sig.resume) { if (!f.inHandler) throw vbErr(20); f.inHandler = false; this.err.clear(); if (sig.resume === 'next') i = f.resumeNext; else if (sig.resume === 'here') i = f.resumeHere; else { const j = labelAt(sig.resume); i = j; } continue; }
        if (sig.exit) throw new CompileError(`Exit ${sig.exit[0].toUpperCase() + sig.exit.slice(1)} not within ${sig.exit[0].toUpperCase() + sig.exit.slice(1)}`, body[i].line, M.name);
        i++; }
    } finally { this.stack.pop(); }
    if (f.inHandler || f.onErr) this.err.clear(); // leaving a procedure that handles errors clears Err
    if (proc.kind === 'function') { const r = f.vars.get(proc.lname).v; return r; } return EMPTY;
  }
  // one top-level statement of a procedure: errors become On Error jumps here
  *execTop(body, i, f) { return yield* this.execStmt(body[i], f, true); }
  *execBlock(ss, f) { for (const s of ss) { const sig = yield* this.execStmt(s, f); if (sig) return sig; } return null; }
  *execStmt(s, f, top) {
    for (;;) {
      this.line = s.line; f.line = s.line; if (s.k !== 'label' && s.k !== 'dim' && s.k !== 'consts') yield { line: s.line, frame: f }; this.clock += 0.000002;
      try { const sig = yield* this.stmt(s, f); if (this.host.pending && this.host.pending.length) yield* this.flushEvents(); if (this.formPending && this.formPending.length) yield* this.flushFormEvents();
        if (this.assertFailed) { this.assertFailed = false; yield { line: s.line, frame: f, stop: true }; } // Debug.Assert False: stop here, like a breakpoint
        return sig; }
      catch (e) {
        if (!(e instanceof VbaError)) throw e;
        if (e.line === undefined) { e.line = s.line; e.mod = f.mod.name; e.proc = f.proc.name; e.frames = this.snapshot(); }
        if (f.inHandler) throw e; // an error inside the handler goes up to the caller
        if (f.onErr === 'next') { this.err.set(e); return null; }
        if (f.onErr && f.onErr.label) { this.err.set(e); const r = yield* this.runHandler(f, f.onErr.label);
          if (r && r.resume === 'next') return null; // Resume Next: the statement after the one that failed, even inside a loop
          if (r && r.resume === 'here') continue; // Resume: try the same statement again
          if (r && r.resume) return { goto: r.resume }; // Resume SomeLabel
          return r; }
        throw e; } } }
  // On Error GoTo Label: run the handler from the label, here and now, so Resume can come back to exactly where the error happened
  *runHandler(f, label) { const body = f.proc.body; const at = l => body.findIndex(x => x.k === 'label' && x.lname === l); let i = at(label) + 1; f.inHandler = true;
    while (i < body.length) { const sig = yield* this.execStmt(body[i], f); if (!sig) { i++; continue; }
      if (sig.resume) { f.inHandler = false; this.err.clear(); return sig; }
      if (sig.goto) { const j = at(sig.goto); if (j < 0) throw new CompileError('Label not defined', body[i].line, f.mod.name, sig.goto); i = j; continue; }
      return sig; }
    return { exit: f.proc.kind === 'function' ? 'function' : 'sub', fromHandler: true }; }
  snapshot() { return (this.stack || []).map(fr => ({ proc: fr.proc.name, mod: fr.mod.name, line: fr.line, vars: fr.locals() })); }
  *stmt(s, f) {
    switch (s.k) {
      case 'label': case 'dim': case 'consts': return null;
      case 'assign': { const v = yield* this.ev(s.e, f, s.set); yield* this.assign(s.target, v, f, s.set); return null; }
      case 'call': { yield* this.callStmt(s.e, f); return null; }
      case 'print': { let line = this.pend || ''; this.pend = ''; for (const it of s.items) { if (it.sep === ',') line += ' '.repeat(14 - (line.length % 14)); else if (it.sep !== ';') line += printStr(deref(yield* this.ev(it.e, f), this)); } if (s.items.length && s.items[s.items.length - 1].sep) this.pend = line; else this.print(line.replace(/\s+$/, '')); return null; }
      case 'if': { for (const b of s.branches) { const c = yield* this.ev(b.cond, f); if (c !== NULLV && toBool(c, this)) return yield* this.execBlock(b.body, f); } if (s.els) return yield* this.execBlock(s.els, f); return null; }
      case 'select': { const x = deref(yield* this.ev(s.e, f), this); for (const c of s.cases) { for (const t of c.tests) { let hit;
            if (t.k === 'eq') hit = toBool(this.compare('=', x, deref(yield* this.ev(t.e, f), this)), this); else if (t.k === 'range') hit = toBool(this.compare('>=', x, deref(yield* this.ev(t.a, f), this))) && toBool(this.compare('<=', x, deref(yield* this.ev(t.b, f), this))); else hit = toBool(this.compare(t.op, x, deref(yield* this.ev(t.e, f), this)));
            if (hit) return yield* this.execBlock(c.body, f); } } if (s.els) return yield* this.execBlock(s.els, f); return null; }
      case 'for': { const from = deref(yield* this.ev(s.from, f), this), to = toNum(deref(yield* this.ev(s.to, f), this), this), step = s.step ? toNum(deref(yield* this.ev(s.step, f), this), this) : 1;
        yield* this.assign(s.v, from, f, false); for (;;) { const cur = toNum(yield* this.ev(s.v, f), this); if (step >= 0 ? cur > to : cur < to) break; yield TICK; const sig = yield* this.execBlock(s.body, f); if (sig) { if (sig.exit === 'for') break; return sig; } const now = toNum(yield* this.ev(s.v, f), this); yield* this.assign(s.v, now + step, f, false); } return null; }
      case 'foreach': { const c = yield* this.ev(s.coll, f); let items; if (c instanceof VArray) items = c.d.slice(); else if (isObj(c) && c.each) items = c.each(this); else if (c === NOTHING) throw vbErr(91); else throw vbErr(c === EMPTY ? 92 : 451, c === EMPTY ? undefined : 'Object not a collection');
        for (const it of items) { yield* this.assign(s.v, it, f, isObj(it)); const sig = yield* this.execBlock(s.body, f); if (sig) { if (sig.exit === 'for') break; return sig; } } return null; }
      case 'do': { for (;;) { yield TICK; if (s.pre) { const c = yield* this.ev(s.pre.e, f); const b = c !== NULLV && toBool(c, this); if (s.pre.until ? b : !b) break; } const sig = yield* this.execBlock(s.body, f); if (sig) { if (sig.exit === 'do') break; return sig; } if (s.post) { const c = yield* this.ev(s.post.e, f); const b = c !== NULLV && toBool(c, this); if (s.post.until ? b : !b) break; } } return null; }
      case 'with': { const o = yield* this.ev(s.e, f, true); if (o === NOTHING) throw vbErr(91); f.with.push(o); try { return yield* this.execBlock(s.body, f); } finally { f.with.pop(); } }
      case 'exit': return { exit: s.what };
      case 'end': return { end: true };
      case 'stop': yield { line: s.line, frame: f, stop: true }; return null;
      case 'onerror': if (s.mode === 'next') f.onErr = 'next'; else if (s.mode === 'zero') { f.onErr = null; this.err.clear(); } else f.onErr = { label: s.label }; return null;
      case 'resume': return { resume: s.mode === 'label' ? s.label : s.mode };
      case 'goto': return { goto: s.label };
      case 'redim': { for (const v of s.vars) { const slot = this.slot(v.name, f, true); if (slot.fixed) throw vbErr(10); const b = []; for (const [lo, hi] of v.dims) b.push([lo ? toNum(yield* this.ev(lo, f)) : 0, toNum(yield* this.ev(hi, f))]); const t = v.type || slot.type || 'variant'; const nv = new VArray(b, t);
          if (s.preserve && slot.v instanceof VArray && slot.v.b.length) { const old = slot.v; if (old.b.length !== b.length || old.b.slice(0, -1).some((x, i) => x[0] !== b[i][0] || x[1] !== b[i][1])) throw vbErr(9); const last = b.length - 1; const each = (bb, fn, pre = []) => { if (pre.length === bb.length) return fn(pre); for (let k = bb[pre.length][0]; k <= bb[pre.length][1]; k++) each(bb, fn, [...pre, k]); }; each(old.b, ix => { if (ix[last] <= b[last][1]) nv.set1(ix, old.get1(ix)); }); }
          slot.v = nv; slot.arr = true; if (v.type) slot.type = v.type; } return null; }
      case 'erase': for (const n of s.names) { const slot = this.slot(n, f); if (slot.v instanceof VArray) { if (slot.fixed) slot.v = new VArray(slot.v.b, slot.v.type); else slot.v = new VArray([], slot.type); } } return null;
      case 'midassign': { const t = s.target; const sv = toStr(yield* this.ev(t.args[0].e, f)); const st = toNum(yield* this.ev(t.args[1].e, f)); const r = toStr(yield* this.ev(s.e, f)); const n = t.args[2] ? Math.min(toNum(yield* this.ev(t.args[2].e, f)), r.length) : r.length; const out = sv.slice(0, st - 1) + r.slice(0, Math.min(n, sv.length - st + 1)) + sv.slice(st - 1 + Math.min(n, sv.length - st + 1)); yield* this.assign(t.args[0].e, out, f, false); return null; }
    }
    throw vbErr(1004, 'Unknown statement'); }
  print(line) { this.out.push(line); if (this.host.print) this.host.print(line); }
  // ---- variables
  slot(name, f, create) { const k = name.toLowerCase(); if (f.vars.has(k)) return f.vars.get(k); if (f.mod.vars.has(k)) return f.mod.vars.get(k); for (const M of this.mods.values()) if (M !== f.mod && M.kind === 'std' && M.vars.has(k)) { const d = M.ast.decls.find(x => x.name.toLowerCase() === k); if (d && d.scope !== 'private' && d.scope !== 'dim') return M.vars.get(k); }
    if (create) { const s = { type: 'variant', v: EMPTY }; f.vars.set(k, s); return s; } return null; }
  *assign(t, v, f, isSet) {
    if (t.k === 'paren') t = t.e;
    if (t.k === 'id') { let slot = this.slot(t.name, f, false);
      if (!slot) { const fc = this.formRef(t.lname, f); if (fc && fc.putDflt) { fc.putDflt([], deref(v, this), this); return; } if (this.host.globalPut && this.host.globalPut(t.lname, [], v, this)) return; slot = this.slot(t.name, f, true); }
      if (slot.isConst) throw new CompileError('Assignment to constant not permitted', this.line, f.mod.name, t.name);
      if (isSet) { if (!(isObj(v) || v === NOTHING)) throw vbErr(424); if (SCALAR.has(slot.type) && slot.type !== 'variant') throw vbErr(13); slot.v = v; return; }
      if (slot.arr && slot.fixed) throw new CompileError("Can't assign to array", this.line, f.mod.name, t.name);
      if (!SCALAR.has(slot.type)) { if (isObj(v) && !v.dflt) throw new CompileError('Invalid use of property', this.line, f.mod.name); if (slot.v === NOTHING || slot.v === null) throw vbErr(91); if (isObj(slot.v) && slot.v.putDflt) { slot.v.putDflt([], v, this); return; } throw vbErr(91); }
      if (v instanceof VArray) { slot.v = slot.type === 'variant' || slot.arr ? v.clone() : (() => { throw vbErr(13); })(); return; }
      if (isObj(v)) v = deref(v, this); if (v === NOTHING) throw vbErr(91); slot.v = coerce(slot.type, v, this); return; }
    if (t.k === 'member') { const o = t.obj ? yield* this.ev(t.obj, f, true) : this.withObj(f); if (!isObj(o)) throw vbErr(o === NOTHING ? 91 : 424); if (!isSet && isObj(v) && !(o.acceptsObj && o.acceptsObj(t.lname))) v = deref(v, this); o.put(t.lname, [], v, this, isSet); this.clock += o.cost ? o.cost('put', t.lname, this) : 0; return; }
    if (t.k === 'call') {
      const args = yield* this.args(t.args, f);
      if (t.fn.k === 'id') { const slot = this.slot(t.fn.name, f, false); if (slot) { if (slot.v instanceof VArray) { if (isObj(v) && !isSet) v = deref(v, this); slot.v.set1(args.map(a => a.v), slot.v.type === 'variant' ? v : coerce(slot.v.type, v, this)); return; } if (isObj(slot.v)) { if (slot.v.putDflt) { slot.v.putDflt(args, isSet ? v : deref(v, this), this); return; } throw vbErr(438); } throw vbErr(13); }
        const o = this.host.global ? this.host.global(t.fn.lname, args, this) : undefined; if (isObj(o)) { if (!o.putDflt) throw vbErr(438); o.putDflt([], isSet ? v : deref(v, this), this); return; } throw new CompileError('Sub or Function not defined', this.line, f.mod.name, t.fn.name); }
      if (t.fn.k === 'member') { const o = t.fn.obj ? yield* this.ev(t.fn.obj, f, true) : this.withObj(f); if (!isObj(o)) throw vbErr(o === NOTHING ? 91 : 424); if (!isSet && isObj(v)) v = deref(v, this);
        if (o.hasArgProp && o.hasArgProp(t.fn.lname)) { o.put(t.fn.lname, args, v, this, isSet); return; } const sub = o.get(t.fn.lname, args, this); if (isObj(sub) && sub.putDflt) { sub.putDflt([], v, this); return; } o.put(t.fn.lname, args, v, this, isSet); return; }
      const o = yield* this.ev(t.fn, f, true); if (isObj(o) && o.putDflt) { o.putDflt(args, isSet ? v : deref(v, this), this); return; } throw vbErr(424); }
    throw new CompileError('Invalid assignment target', this.line, f.mod.name);
  }
  withObj(f) { if (!f.with.length) throw new CompileError('Invalid or unqualified reference', this.line, f.mod.name); return f.with[f.with.length - 1]; }
  *args(list, f) { const out = []; for (const a of list) { if (!a.e) { out.push({ name: a.name, v: MISSING }); continue; } const v = yield* this.ev(a.e, f, true); out.push({ name: a.name, v }); } return out; }
  // arguments for a user procedure: a plain variable goes by reference
  *refArgs(list, f) { const out = []; for (const a of list) { if (!a.e) { out.push({ name: a.name, v: MISSING }); continue; } if (a.e.k === 'id') { const slot = this.slot(a.e.name, f, false); if (slot && !slot.isConst) { out.push({ name: a.name, v: slot.v, ref: slot }); continue; } } out.push({ name: a.name, v: yield* this.ev(a.e, f, true) }); } return out; }
  *callStmt(e, f) { const fn = e.fn; if (fn.k === 'id') { const P = this.userProc(fn.lname, f); if (P) { const args = yield* this.refArgs(e.args, f); return yield* this.invoke(P.proc, P.mod, args); } } return yield* this.ev(e, f, true, true); }
  userProc(k, f) { const own = f.mod.ast.procs.find(p => p.lname === k); if (own) return { proc: own, mod: f.mod }; const P = this.procs.get(k); if (!P || (P.proc.scope === 'private' && P.mod !== f.mod)) return null; return P; }
  // ---- expressions: keepObj returns objects as they are; otherwise the caller derefs where it needs a value
  *ev(e, f, keepObj, isStmt) {
    switch (e.k) {
      case 'lit': return e.v;
      case 'paren': { const v = yield* this.ev(e.e, f); return isObj(v) ? deref(v, this) : v; }
      case 'me': { if (f.me) return f.me; if (f.mod.kind === 'form' && this.forms.get(f.mod.name.toLowerCase())) return this.forms.get(f.mod.name.toLowerCase()); const m = this.host.me ? this.host.me(f.mod, this) : null; if (!m) throw new CompileError('Invalid use of Me keyword', this.line, f.mod.name); return m; }
      case 'new': return this.newObj(e.cls);
      case 'id': return yield* this.evId(e, f, isStmt);
      case 'member': { const o = e.obj ? yield* this.ev(e.obj, f, true) : this.withObj(f); return yield* this.member(o, e.lname, [], f, e.name, isStmt); }
      case 'call': return yield* this.evCall(e, f, isStmt);
      case 'un': { const a = deref(yield* this.ev(e.a, f), this); if (a === NULLV) return NULLV; if (e.op === '-') { if (a instanceof VDate) return new VDate(-a.n); const n = toNum(a, this); return n === 0 ? 0 : -n; } if (typeof a === 'boolean') return !a; return ~Math.round(toNum(a, this)); }
      case 'bin': { if (e.op === 'is') { const a = yield* this.ev(e.a, f, true), b = yield* this.ev(e.b, f, true); if (!(isObj(a) || a === NOTHING) || !(isObj(b) || b === NOTHING)) throw vbErr(424); return a === b || (a.same && a.same(b)); }
        const a = deref(yield* this.ev(e.a, f), this), b = deref(yield* this.ev(e.b, f), this); return this.binop(e.op, a, b); }
    }
    throw vbErr(1004); }
  *evId(e, f, isStmt) {
    const k = e.lname; const slot = this.slot(e.name, f, false);
    if (slot) { if (slot.autoNew && (slot.v === null || slot.v === NOTHING)) slot.v = this.newObj(slot.autoNew); if (slot.v === MISSING && !isStmt) return MISSING; return slot.v; }
    if (k in CONSTS) return CONSTS[k]; if (this.host.consts && k in this.host.consts) return this.host.consts[k];
    if (k === 'err') return this.err; if (k === 'debug') return this.debugObj || (this.debugObj = new DebugObj(this));
    const P = this.userProc(k, f); if (P) return yield* this.invoke(P.proc, P.mod, []);
    if (k === 'msgbox' || k === 'inputbox') return yield* this.ui(k, [], f);
    if (BI[k]) return BI[k]([], this);
    const fo = this.formRef(k, f); if (fo) { if (fo instanceof UserForm) yield* this.loadForm(fo); return fo; }
    if (this.host.global) { const g = this.host.global(k, [], this); if (g !== undefined) return g; }
    if (this.host.fn && this.host.fn(k)) return this.host.fn(k)([], this);
    // an undeclared variable (no Option Explicit): a new Variant
    return this.slot(e.name, f, true).v; }
  *evCall(e, f, isStmt) {
    const fn = e.fn;
    if (fn.k === 'id') { const k = fn.lname; const slot = this.slot(fn.name, f, false);
      if (slot) { const v = slot.autoNew && !slot.v ? (slot.v = this.newObj(slot.autoNew)) : slot.v; const args = yield* this.args(e.args, f); if (v instanceof VArray) return v.get1(args.map(a => deref(a.v, this))); if (isObj(v)) { if (!v.dflt) throw vbErr(438); return v.dflt(args, this); } if (slot.ret) { const P = this.userProc(k, f); if (P) return yield* this.invoke(P.proc, P.mod, yield* this.refArgs(e.args, f)); } throw vbErr(13); }
      const P = this.userProc(k, f); if (P) return yield* this.invoke(P.proc, P.mod, yield* this.refArgs(e.args, f));
      if (k === 'msgbox' || k === 'inputbox') return yield* this.ui(k, yield* this.args(e.args, f), f);
      if (k === 'unload' || k === 'load') { const a = yield* this.args(e.args, f); const o = a[0] && a[0].v; if (!(o instanceof UserForm)) throw vbErr(361, "Can't load or unload this object"); if (k === 'unload') { if (o.loaded && !(yield* this.queryClose(o, 1))) o.unload(); } else yield* this.loadForm(o); return EMPTY; }
      if (k === 'typename' || k === 'vartype') { const a = e.args[0] && e.args[0].e; if (a && a.k === 'id') { const s = this.slot(a.name, f, false); if (s && s.type !== 'variant' && SCALAR.has(s.type) && !(s.v instanceof VArray)) return k === 'typename' ? typeLabel(s.type) : { integer: 2, long: 3, single: 4, double: 5, currency: 6, date: 7, string: 8, boolean: 11, byte: 17 }[s.type] ?? 0; } }
      if (BI[k]) { const args = yield* this.args(e.args, f); return BI[k](bindBI(k, args).map(v => k === 'iif' || k === 'array' || k === 'typename' || k === 'isobject' || k === 'ismissing' || k === 'vartype' ? v : deref(v, this)), this); }
      if (this.host.fn && this.host.fn(k)) { const args = yield* this.args(e.args, f); return yield* wrapUi(this.host.fn(k)(args, this)); }
      if (this.host.global) { const args = yield* this.args(e.args, f); const g = this.host.global(k, args, this); if (g !== undefined) return g; }
      throw new CompileError('Sub or Function not defined', this.line, f.mod.name, fn.name); }
    if (fn.k === 'member') { const o = fn.obj ? yield* this.ev(fn.obj, f, true) : this.withObj(f); const args = yield* this.args(e.args, f); return yield* this.member(o, fn.lname, args, f, fn.name, isStmt); }
    const o = yield* this.ev(fn, f, true); const args = yield* this.args(e.args, f); if (o instanceof VArray) return o.get1(args.map(a => deref(a.v, this))); if (isObj(o) && o.dflt) return o.dflt(args, this); throw vbErr(o === NOTHING ? 91 : 13); }
  *member(o, name, args, f, display, isStmt) {
    if (o === NOTHING || o === EMPTY) throw vbErr(91); if (!isObj(o)) throw vbErr(424);
    this.clock += o.cost ? o.cost('get', name, this) : 0;
    if (o instanceof UserForm && name === 'show') { const [mode] = bind(args, ['modal']); return yield* this.showForm(o, has(mode) && toNum(deref(mode, this)) === 0); }
    if (o.uiMember && o.uiMember(name)) { const req = o.uiRequest(name, args, this); const r = yield req; return o.uiResult ? o.uiResult(name, r, this) : r; }
    const v = o.get(name, args, this); if (v === undefined) throw vbErr(438); if (v && v.vbaRun) { const P = this.findProc(v.name.replace(/^.*!/, '')); if (!P) throw vbErr(1004, `Cannot run the macro '${v.name}'. The macro may not be available in this workbook or all macros may be disabled.`); return yield* this.invoke(P.proc, P.mod, v.args.map(x => ({ v: x }))); } return yield* wrapUi(v); }
  *ui(k, args, f) { const [p, b, t, d] = bind(args, k === 'msgbox' ? ['prompt', 'buttons', 'title', 'helpfile'] : ['prompt', 'title', 'default', 'xpos']);
    if (k === 'msgbox') { const req = { ui: 'msgbox', prompt: toStr(deref(p, this)), buttons: has(b) ? toNum(b) : 0, title: has(t) ? toStr(t) : 'Microsoft Excel' }; const r = yield req; return r === undefined ? defaultButton(req.buttons) : r; }
    const req = { ui: 'inputbox', prompt: toStr(deref(p, this)), title: has(b) ? toStr(b) : 'Microsoft Excel', def: has(t) ? toStr(t) : '' }; const r = yield req; return r === undefined || r === null ? '' : String(r); }
  newObj(cls, byProgId) { cls = cls.toLowerCase(); if (cls === 'collection' && !byProgId) return new VCollection(); if (cls === 'dictionary' || cls === 'scripting.dictionary') return new VDictionary(); if (this.host.newObj) { const o = this.host.newObj(cls, this, byProgId); if (o) return o; } if (byProgId) throw vbErr(429); throw new CompileError('User-defined type not defined', this.line, '', cls); }
  // ---- operators
  binop(op, a, b) {
    if (op === '&') return (a === NULLV ? '' : toStr(a, this)) + (b === NULLV ? '' : toStr(b, this));
    if ((a === NULLV || b === NULLV) && op !== 'and' && op !== 'or') return NULLV;
    switch (op) {
      case '+': if (typeof a === 'string' && typeof b === 'string') return a + b; if (a === EMPTY && typeof b === 'string') return b; if (b === EMPTY && typeof a === 'string') return a; if (a instanceof VDate || b instanceof VDate) return new VDate(toNum(a, this) + toNum(b, this)); return fix(toNum(a, this) + toNum(b, this));
      case '-': if (a instanceof VDate && b instanceof VDate) return a.n - b.n; if (a instanceof VDate) return new VDate(a.n - toNum(b, this)); return fix(toNum(a, this) - toNum(b, this));
      case '*': return fix(toNum(a, this) * toNum(b, this));
      case '/': { const d = toNum(b, this); if (d === 0) throw vbErr(toNum(a, this) === 0 ? 6 : 11); return toNum(a, this) / d; }
      case '\\': { const x = bankRound(toNum(a, this)), y = bankRound(toNum(b, this)); if (y === 0) throw vbErr(11); return Math.trunc(x / y); }
      case 'mod': { const x = bankRound(toNum(a, this)), y = bankRound(toNum(b, this)); if (y === 0) throw vbErr(11); const r = x % y; return r === 0 ? 0 : r; }
      case '^': return Math.pow(toNum(a, this), toNum(b, this));
      case 'and': case 'or': case 'xor': case 'eqv': case 'imp': { if (a === NULLV || b === NULLV) { if (op === 'and' && (a === false || b === false)) return false; if (op === 'or' && (a === true || b === true)) return true; return NULLV; } if (typeof a === 'boolean' && typeof b === 'boolean') return { and: a && b, or: a || b, xor: a !== b, eqv: a === b, imp: !a || b }[op]; const x = Math.round(toNum(a, this)), y = Math.round(toNum(b, this)); const r = { and: x & y, or: x | y, xor: x ^ y, eqv: ~(x ^ y), imp: ~x | y }[op]; return (typeof a === 'boolean' || typeof b === 'boolean') && (r === -1 || r === 0) ? r === -1 : r; }
      case 'like': return likeRe(toStr(b, this), this.compareText).test(toStr(a, this));
      default: return this.compare(op, a, b);
    } }
  compare(op, a, b) {
    if (a === NULLV || b === NULLV) return NULLV; let c;
    const numLike = v => typeof v === 'number' || typeof v === 'boolean' || v instanceof VDate;
    if (a === EMPTY) a = typeof b === 'string' ? '' : 0; if (b === EMPTY) b = typeof a === 'string' ? '' : 0;
    if (numLike(a) && numLike(b)) { const x = toNum(a), y = toNum(b); c = x < y ? -1 : x > y ? 1 : 0; }
    else if (typeof a === 'string' && typeof b === 'string') { let x = a, y = b; if (this.compareText) { x = x.toLowerCase(); y = y.toLowerCase(); } c = x < y ? -1 : x > y ? 1 : 0; }
    else { const s = numLike(a) ? b : a; let n = null; try { n = toNum(s); } catch (e) { n = null; } if (n === null) c = numLike(a) ? -1 : 1; else { const x = numLike(a) ? toNum(a) : n, y = numLike(b) ? toNum(b) : n; c = x < y ? -1 : x > y ? 1 : 0; } }
    return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op]; }
  get compareText() { const f = this.stack && this.stack[this.stack.length - 1]; return !!(f && f.mod.ast.options.compare === 'text'); }
  // ---- the Immediate window: ? expression, or a statement
  *immediate(text, frame) {
    const t = text.trim(); if (!t) return null; const M = frame ? frame.mod : [...this.mods.values()].find(m => m.kind === 'std') || [...this.mods.values()][0];
    const f = frame || new Frame({ name: '(Immediate)', body: [], params: [], lname: '' }, M || { name: 'Immediate', ast: { options: {}, procs: [] }, vars: new Map(), kind: 'std' });
    this.stack = this.stack && this.stack.length ? this.stack : [f];
    if (/^(\?|print\b)/i.test(t)) { const src = t.replace(/^(\?|print)\s*/i, ''); const P = new Parser('Debug.Print ' + src + '\n', 'Immediate'); const s = P.stmt(); return yield* this.stmt(s, f); }
    const P = new Parser(t + '\n', 'Immediate'); const s = P.stmt(); if (!s) return null; return yield* this.stmt(s, f);
  }
}
const TICK = { tick: true };
const fix = n => (Number.isInteger(n) ? n : Number(n.toPrecision(15)));
function defaultButton(b) { const k = b & 7; return [1, 1, 3, 6, 6, 4][k] || 1; }
const BI_PARAMS = { msgbox: ['prompt', 'buttons', 'title'], inputbox: ['prompt', 'title', 'default'], format: ['expression', 'format'], replace: ['expression', 'find', 'replace', 'start', 'count', 'compare'], split: ['expression', 'delimiter', 'limit', 'compare'], dateadd: ['interval', 'number', 'date'], datediff: ['interval', 'date1', 'date2'], instr: ['start', 'string1', 'string2', 'compare'], round: ['number', 'numdigitsafterdecimal'], mid: ['string', 'start', 'length'], left: ['string', 'length'], right: ['string', 'length'], dateserial: ['year', 'month', 'day'], weekday: ['date', 'firstdayofweek'] };
function bindBI(k, args) { if (!args.some(a => a.name)) return args.map(a => a.v); return bind(args, BI_PARAMS[k] || []); }
// an object member may ask the lab for something (a form shown modally): it returns { ui } and we yield it
function* wrapUi(v) { if (v && typeof v === 'object' && v.ui && !v.vbObj) { const r = yield v; return v.after ? v.after(r) : r; } return v; }
export function printStr(v) { if (typeof v === 'number') return (v >= 0 ? ' ' : '') + numStr(v) + ' '; if (v === NULLV) return 'Null'; if (v instanceof VArray) throw vbErr(13); if (v === NOTHING) throw vbErr(91); return toStr(v); }
class DebugObj extends VObject { constructor(rt) { super(); this.rt = rt; } get tn() { return 'Debug'; } get(n, a) { if (n === 'print') { this.rt.print(a.map(x => printStr(deref(x.v, this.rt))).join('')); return EMPTY; } if (n === 'assert') { if (!toBool(a[0].v)) this.rt.assertFailed = true; return EMPTY; } throw vbErr(438); } }
export class Frame { constructor(proc, mod, me) { this.proc = proc; this.mod = mod; this.me = me || null; this.vars = new Map(); this.with = []; this.onErr = null; this.inHandler = false; this.line = proc.line; }
  locals() { return [...this.vars.entries()].filter(([k, s]) => !s.isConst).map(([k, s]) => ({ name: displayName(this, k), value: s.v, type: s.arr ? typeLabel(s.type) + '()' : typeLabel(s.type), ret: !!s.ret })); } }
function displayName(f, k) { const fromParam = f.proc.params.find(p => p.name.toLowerCase() === k); if (fromParam) return fromParam.name; const walk = ss => { for (const s of ss || []) { if (s.k === 'dim') for (const v of s.vars) if (v.name.toLowerCase() === k) return v.name; for (const kk of ['body', 'els']) if (Array.isArray(s[kk])) { const r = walk(s[kk]); if (r) return r; } if (s.branches) for (const b of s.branches) { const r = walk(b.body); if (r) return r; } if (s.cases) for (const c of s.cases) { const r = walk(c.body); if (r) return r; } } return null; }; return walk(f.proc.body) || (f.proc.lname === k ? f.proc.name : k); }

// ---------------------------------------------------------------- running: a session drives the generator
// mode 'run' stops only at breakpoints and Stop; 'into' at the next statement; 'over' at the next statement in this or an outer procedure; 'out' when this procedure returns.
export class Session {
  constructor(prog, opts = {}) { this.prog = prog; this.o = { maxSteps: 3e6, breakpoints: new Set(), ui: null, ...opts }; this.it = null; this.state = 'idle'; this.steps = 0; this.cur = null; this.result = undefined; }
  start(name, args = []) { this.prog.stack = []; this.it = this.prog.call(name, args); this.state = 'ready'; this.steps = 0; this.cur = null; this.pending = undefined; return this; }
  startImmediate(text, frame) { const saved = this.prog.stack; this.it = this.prog.immediate(text, frame); this.state = 'ready'; this.cur = null; this.pending = undefined; this.restoreStack = saved; return this; }
  depth() { return (this.prog.stack || []).length; }
  // run synchronously (checks): ui requests are answered at once by o.ui; returns { done, value } | { error } | { paused }
  runSync(mode = 'run') { return this._go(mode, false); }
  async run(mode = 'run') { return this._go(mode, true); }
  _go(mode, async) {
    const startDepth = this.depth();
    const step = (input) => {
      for (;;) {
        let r; try { r = this.it.next(input); } catch (e) { return this.fail(e); } input = undefined;
        if (r.done) { this.state = 'done'; this.result = r.value; this.cur = null; if (this.restoreStack) { this.prog.stack = this.restoreStack; this.restoreStack = null; } return { done: true, value: r.value }; }
        const y = r.value;
        if (y === TICK) { if (++this.steps > this.o.maxSteps) return this.fail(Object.assign(vbErr(1004, 'The macro ran for too long and was stopped: is there a loop that never ends?'), { line: this.prog.line, mod: this.prog.stack && this.prog.stack.length ? this.prog.stack[this.prog.stack.length - 1].mod.name : undefined, runaway: true })); continue; }
        if (y.ui) { const ans = this.o.ui ? this.o.ui(y) : undefined; if (ans && typeof ans.then === 'function') { if (!async) return this.fail(vbErr(1004, 'This run can’t wait for an answer.')); this.state = 'waiting'; return ans.then(v => { this.state = 'running'; return step(v); }); } input = ans; continue; }
        this.steps++; if (this.steps > this.o.maxSteps) return this.fail(Object.assign(vbErr(1004, 'The macro ran for too long and was stopped: is there a loop that never ends?'), { line: y.line, mod: y.frame.mod.name, runaway: true }));
        const d = this.depth(); const bp = this.o.breakpoints.has(`${y.frame.mod.name.toLowerCase()}:${y.line}`) && !y.enter;
        let pause = y.stop || bp; if (!pause) { if (mode === 'into') pause = true; else if (mode === 'over') pause = d <= startDepth && !y.enter; else if (mode === 'out') pause = d < startDepth && !y.enter; }
        if (mode === 'into' && y.enter) pause = true;
        if (pause) { this.state = 'paused'; this.cur = y; return { paused: true, line: y.line, mod: y.frame.mod.name, proc: y.frame.proc.name, stop: !!y.stop }; }
      } };
    this.state = 'running'; return step(this.pending);
  }
  fail(e) { this.state = 'error'; if (this.restoreStack) { this.prog.stack = this.restoreStack; this.restoreStack = null; } if (e instanceof EndSignal) { this.state = 'done'; return { done: true, ended: true }; } if (e instanceof VbaError || e instanceof CompileError) { this.error = e; return { error: e }; } this.error = e; return { error: Object.assign(vbErr(1004, 'Sheetah couldn’t run this line: ' + (e && e.message)), { line: this.prog.line, internal: e }) }; }
  locals() { const y = this.cur; return y ? y.frame.locals() : []; }
  frame() { return this.cur ? this.cur.frame : null; }
}
// run a macro to the end, synchronously: answers MsgBox with its default button and InputBox from a list
export function runMacro(prog, name, opts = {}) { const inputs = (opts.inputs || []).slice(); const answers = (opts.answers || []).slice(); const log = [];
  const s = new Session(prog, { ui: req => { log.push(req); if (req.ui === 'inputbox') return inputs.length ? inputs.shift() : ''; if (req.ui === 'msgbox') return answers.length ? answers.shift() : undefined; return opts.ui ? opts.ui(req) : undefined; }, maxSteps: opts.maxSteps || 3e6 });
  s.start(name, opts.args || []); const r = s.runSync('run'); if (r.paused) { let rr = r; while (rr.paused) rr = s.runSync('run'); return { ...rr, ui: log, out: prog.out }; } return { ...r, ui: log, out: prog.out }; }

// ---------------------------------------------------------------- UserForms (MSForms): a form's default instance and its controls
const FM_DEFAULT = { textbox: 'value', combobox: 'value', listbox: 'value', checkbox: 'value', optionbutton: 'value', togglebutton: 'value', label: 'caption', commandbutton: 'caption', frame: 'caption', image: 'picture' };
const FM_NAMES = { textbox: 'TextBox', combobox: 'ComboBox', listbox: 'ListBox', checkbox: 'CheckBox', optionbutton: 'OptionButton', togglebutton: 'ToggleButton', label: 'Label', commandbutton: 'CommandButton', frame: 'Frame', image: 'Image' };
export class FormCtl extends VObject {
  constructor(d, F) { super(); this.F = F; this.d = d; this.lt = String(d.type || 'TextBox').toLowerCase(); this.name = d.name; this.caption = d.caption ?? ''; this.enabled = d.enabled !== false; this.visible = d.visible !== false; this.list = (d.items || []).map(String); this.locked = !!d.locked; this.tag = d.tag || ''; this.group = d.group || '';
    this.val = d.value !== undefined ? d.value : (this.lt === 'checkbox' || this.lt === 'optionbutton' || this.lt === 'togglebutton' ? false : this.lt === 'listbox' ? NULLV : ''); this.backColor = d.backColor; this.foreColor = d.foreColor; this.geo = { left: d.left || 0, top: d.top || 0, width: d.width || 72, height: d.height || 18 }; }
  get tn() { return FM_NAMES[this.lt] || 'Control'; }
  get isList() { return this.lt === 'combobox' || this.lt === 'listbox'; }
  get listIndex() { if (!this.isList) return -1; const v = this.val; return v === NULLV || v === '' || v === null ? -1 : this.list.findIndex(x => x === String(v)); }
  coerceVal(v) { v = deref(v); if (this.lt === 'checkbox' || this.lt === 'optionbutton' || this.lt === 'togglebutton') return v === NULLV ? NULLV : toBool(v); if (this.lt === 'listbox') return v === NULLV || v === EMPTY ? NULLV : toStr(v); return v === NULLV ? '' : toStr(v); }
  setVal(v) { const nv = this.coerceVal(v); const changed = nv !== this.val; this.val = nv; if (this.lt === 'optionbutton' && nv === true) for (const c of this.F.order) if (c !== this && c.lt === 'optionbutton' && (c.group || '') === (this.group || '')) c.val = false; return changed; }
  setByUser(v) { if (this.locked) return false; if (this.lt === 'combobox' && this.d.style !== 0 && v !== '' && !this.list.includes(String(v))) return false; return this.setVal(v); }
  get(n, a, rt) { switch (n) {
    case 'value': return this.val; case 'text': return this.lt === 'listbox' && this.val === NULLV ? '' : toStr(this.val === NULLV ? '' : this.val); case 'caption': return this.caption; case 'name': return this.name;
    case 'enabled': return this.enabled; case 'visible': return this.visible; case 'locked': return this.locked; case 'tag': return this.tag;
    case 'listindex': return this.listIndex; case 'listcount': return this.list.length;
    case 'list': { if (a.length) { const i = Math.round(toNum(deref(a[0].v))); if (i < 0 || i >= this.list.length) throw vbErr(381, 'Could not get the List property. Invalid property array index.'); return this.list[i]; } return new VArray([[0, Math.max(0, this.list.length - 1)]], 'variant', this.list.length ? this.list.slice() : [EMPTY]); }
    case 'additem': { const [x, i] = bind(a, ['pvargitem', 'pvargindex']); const item = toStr(deref(x)); if (has(i)) this.list.splice(Math.round(toNum(i)), 0, item); else this.list.push(item); return EMPTY; }
    case 'removeitem': { const i = Math.round(toNum(deref(a[0].v))); if (i < 0 || i >= this.list.length) throw vbErr(-2147024809, 'Invalid argument.'); const [gone] = this.list.splice(i, 1); if (String(this.val) === gone) this.val = this.lt === 'listbox' ? NULLV : ''; return EMPTY; }
    case 'clear': this.list = []; this.val = this.lt === 'listbox' ? NULLV : ''; return EMPTY;
    case 'setfocus': this.F.focus = this.name; return EMPTY;
    case 'left': case 'top': case 'width': case 'height': return this.geo[n];
    case 'backcolor': return this.backColor ?? 16777215; case 'forecolor': return this.foreColor ?? 0; case 'controltiptext': return this.d.tip || ''; case 'maxlength': return this.d.maxLength || 0; case 'style': return this.d.style ?? (this.lt === 'combobox' ? 2 : 0);
    case 'parent': return this.F; case 'object': return this; }
    return undefined; }
  queue(changed) { if (!changed || !this.F.loaded || !this.F.prog) return; this.F.prog.formPending.push([this.F, this.name + '_Change']); if (this.isList || this.lt === 'checkbox' || this.lt === 'optionbutton') this.F.prog.formPending.push([this.F, this.name + '_Click']); }
  put(n, a, v, rt) { v = deref(v); switch (n) {
    case 'value': this.queue(this.setVal(v)); return; case 'text': this.queue(this.setVal(v)); return; case 'caption': this.caption = toStr(v); return; case 'enabled': this.enabled = toBool(v); return; case 'visible': this.visible = toBool(v); return; case 'locked': this.locked = toBool(v); return; case 'tag': this.tag = toStr(v); return;
    case 'listindex': { const i = Math.round(toNum(v)); if (i === -1) { this.queue(this.setVal(this.lt === 'listbox' ? NULLV : '')); return; } if (i < -1 || i >= this.list.length) throw vbErr(380, 'Could not set the ListIndex property. Invalid property value.'); this.queue(this.setVal(this.list[i])); return; }
    case 'list': { if (a.length) { const i = Math.round(toNum(deref(a[0].v))); this.list[i] = toStr(v); return; } this.list = v instanceof VArray ? v.d.map(x => toStr(x)) : [toStr(v)]; return; }
    case 'left': case 'top': case 'width': case 'height': this.geo[n] = toNum(v); return; case 'backcolor': this.backColor = toNum(v); return; case 'forecolor': this.foreColor = toNum(v); return; case 'controltiptext': this.d.tip = toStr(v); return; case 'maxlength': this.d.maxLength = toNum(v); return; case 'style': this.d.style = toNum(v); return; }
    throw vbErr(438); }
  dflt(a, rt) { return this.get(FM_DEFAULT[this.lt] || 'value', a, rt); } putDflt(a, v, rt) { this.put(FM_DEFAULT[this.lt] || 'value', a, v, rt); }
  snap() { const o = { type: this.tn, value: this.val === NULLV ? null : this.val, caption: this.caption, enabled: this.enabled, visible: this.visible }; if (this.isList) o.list = this.list.slice(); return o; } }
class FormControls extends VObject { constructor(F) { super(); this.F = F; } get tn() { return 'Controls'; } each() { return this.F.order.slice(); }
  get(n, a) { if (n === 'count') return this.F.order.length; if (n === 'item') return this.dflt(a); return undefined; }
  dflt(a) { const k = deref(a[0].v); const c = typeof k === 'string' ? this.F.ctl(k) : this.F.order[Math.round(toNum(k))]; if (!c) throw vbErr(-2147024809, 'Could not find the specified object.'); return c; } }
export class UserForm extends VObject {
  constructor(def, M) { super(); this.def = def; this.M = M; this.name = M.name; this.loaded = false; this.visible = false; this.reset(); }
  get tn() { return this.name; }
  reset() { this.caption = this.def.caption || this.name; this.ctls = new Map(); this.order = []; this.focus = null; for (const c of this.def.controls || []) { const k = new FormCtl(c, this); this.ctls.set(c.name.toLowerCase(), k); this.order.push(k); } }
  load() { this.reset(); this.loaded = true; } unload() { this.loaded = false; this.visible = false; }
  ctl(n) { return this.ctls.get(String(n).toLowerCase()); }
  get(n, a, rt) { switch (n) { case 'caption': return this.caption; case 'name': return this.name; case 'hide': this.visible = false; return EMPTY; case 'repaint': return EMPTY; case 'controls': return a.length ? new FormControls(this).dflt(a) : new FormControls(this);
    case 'width': return this.def.width || 240; case 'height': return this.def.height || 180; case 'activecontrol': return this.focus ? this.ctl(this.focus) : NOTHING; case 'visible': return this.visible; }
    const c = this.ctl(n); if (c) return a.length ? c.dflt(a, rt) : c; return undefined; }
  put(n, a, v) { v = deref(v); if (n === 'caption') { this.caption = toStr(v); return; } if (n === 'width' || n === 'height') { this.def = { ...this.def, [n]: toNum(v) }; return; } const c = this.ctl(n); if (c) { c.putDflt(a, v); return; } throw vbErr(438); }
  snap() { return { caption: this.caption, controls: Object.fromEntries(this.order.map(c => [c.name, c.snap()])) }; } }
