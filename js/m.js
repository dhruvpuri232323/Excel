// M, the Power Query formula language: a parser and a lazy evaluator for the part of M that Excel's Power Query
// Editor writes, and that people type in Custom Column and the Advanced Editor.
// Values: null, true/false, numbers, text, MDate / MDateTime / MTime / MDuration, lists (arrays), MRecord, MTable,
// MFn, MType and MBinary. A cell that failed inside a table holds an ErrCell; reading it raises the error again.

export class MError extends Error {
  constructor(reason, message, detail) { super(message); this.reason = reason || 'Expression.Error'; this.detail = detail || null; }
  get text() { return `${this.reason}: ${this.message}`; }
}
const err = (m, reason = 'Expression.Error') => new MError(reason, m);
export class ErrCell { constructor(e) { this.e = e; } }
export class MDate { constructor(n) { this.n = Math.floor(n); } }
export class MDateTime { constructor(n) { this.n = n; } }
export class MTime { constructor(n) { this.n = n; } }
export class MDuration { constructor(n) { this.n = n; } }
export class MBinary { constructor(name, text, extra) { this.name = name; this.text = text; this.extra = extra || null; } }
export class MType { constructor(name, o = {}) { this.name = name; Object.assign(this, o); } }
export class MRecord {
  constructor(names, vals) { this.names = names; this.vals = vals; }
  idx(n) { return this.names.indexOf(n); }
  get(n, opt) { const i = this.idx(n); if (i < 0) { if (opt) return null; throw err(`The field '${n}' of the record wasn't found.`); } return cell(this.vals[i]); }
  static of(o) { return new MRecord(Object.keys(o), Object.values(o)); }
}
export class MTable {
  constructor(cols, rows, types) { this.cols = cols; this.rows = rows; this.types = types || cols.map(() => null); }
  ci(n, opt) { const i = this.cols.indexOf(n); if (i < 0 && !opt) throw err(`The column '${n}' of the table wasn't found.`); return i; }
  rec(i) { return new MRecord(this.cols, this.rows[i]); }
  col(n) { const i = this.ci(n); return this.rows.map(r => r[i]); }
}
export class MFn { constructor(params, call, name) { this.params = params; this.call = call; this.name = name || null; } }
const cell = v => { if (v instanceof ErrCell) throw v.e; if (v && v.th) return force(v); return v; };

// ---------------------------------------------------------------- lexer
const KW = new Set(['and', 'as', 'each', 'else', 'error', 'false', 'if', 'in', 'is', 'let', 'meta', 'not', 'null', 'or', 'otherwise', 'section', 'shared', 'then', 'true', 'try', 'type']);
const isIdStart = c => /[A-Za-z_À-￿]/.test(c);
const isIdPart = c => /[A-Za-z0-9_À-￿]/.test(c);
class Lexer {
  constructor(src) { this.s = src; this.p = 0; }
  ws() { const s = this.s; for (;;) { while (this.p < s.length && /\s/.test(s[this.p])) this.p++; if (s.startsWith('//', this.p)) { while (this.p < s.length && s[this.p] !== '\n') this.p++; continue; } if (s.startsWith('/*', this.p)) { const e = s.indexOf('*/', this.p + 2); this.p = e < 0 ? s.length : e + 2; continue; } break; } }
  str(at) { const s = this.s; let i = at + 1, out = ''; for (;;) { if (i >= s.length) throw this.fail('Token Literal expected.', at); const c = s[i];
      if (c === '"') { if (s[i + 1] === '"') { out += '"'; i += 2; continue; } return { v: out, end: i + 1 }; }
      if (c === '#' && s[i + 1] === '(') { const e = s.indexOf(')', i); const inner = s.slice(i + 2, e); out += inner.split(',').map(x => x.trim()).map(x => ({ lf: '\n', cr: '\r', tab: '\t', '#': '#' }[x] ?? String.fromCharCode(parseInt(x, 16)))).join(''); i = e + 1; continue; }
      out += c; i++; } }
  next() { this.ws(); const s = this.s, p = this.p, c = s[p];
    if (p >= s.length) return { t: 'eof', p };
    if (c === '"') { const r = this.str(p); this.p = r.end; return { t: 'str', v: r.v, p }; }
    if (c === '#') { if (s[p + 1] === '"') { const r = this.str(p + 1); this.p = r.end; return { t: 'id', v: r.v, p, q: true }; } let i = p + 1; while (i < s.length && isIdPart(s[i])) i++; this.p = i; return { t: 'hash', v: s.slice(p, i), p }; }
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(s[p + 1] || ''))) { const m = /^(0[xX][0-9a-fA-F]+|(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)/.exec(s.slice(p)); this.p = p + m[0].length; return { t: 'num', v: Number(m[0]), p }; }
    if (isIdStart(c)) { let i = p; while (i < s.length && (isIdPart(s[i]) || (s[i] === '.' && isIdStart(s[i + 1] || '')))) i++; const v = s.slice(p, i); this.p = i; return KW.has(v) ? { t: 'kw', v, p } : { t: 'id', v, p }; }
    for (const op of ['=>', '<=', '>=', '<>', '??', '...', '..', '=', '<', '>', '+', '-', '*', '/', '&', '(', ')', '[', ']', '{', '}', ',', ';', '?', '@', '!']) if (s.startsWith(op, p)) { this.p = p + op.length; return { t: 'op', v: op, p }; }
    throw this.fail(`Invalid character '${c}'.`, p); }
  fail(m, at) { const e = err(m, 'Expression.SyntaxError'); e.pos = at ?? this.p; return e; }
}

// ---------------------------------------------------------------- parser
export function parseM(src) { const P = new Parser(src); const e = P.expr(); const t = P.peek(); if (t.t !== 'eof') throw P.lx.fail(`Token Eof expected.`, t.p); return e; }
class Parser {
  constructor(src) { this.lx = new Lexer(src); this.src = src; this.buf = null; }
  peek() { if (!this.buf) { const save = this.lx.p; this.buf = this.lx.next(); this.buf.start = save; } return this.buf; }
  take() { const t = this.peek(); this.buf = null; return t; }
  is(t, v) { const k = this.peek(); return k.t === t && (v === undefined || k.v === v); }
  eat(t, v) { if (this.is(t, v)) return this.take(); return null; }
  want(t, v) { const k = this.peek(); if (k.t === t && (v === undefined || k.v === v)) return this.take(); throw this.lx.fail(`Token ${v || t} expected.`, k.p); }
  mark() { return { p: this.lx.p, buf: this.buf }; }
  reset(m) { this.lx.p = m.p; this.buf = m.buf; }
  pos() { return this.buf ? this.buf.p : (this.lx.ws(), this.lx.p); }
  expr() {
    const k = this.peek();
    if (k.t === 'kw') {
      if (k.v === 'let') { this.take(); const binds = []; do { const ns = this.peek().p; const n = this.name(); this.want('op', '='); const s = this.pos(); const e = this.expr(); binds.push({ name: n, e, ns, s, end: this.peek().p }); } while (this.eat('op', ','));
        this.want('kw', 'in'); return { k: 'let', binds, body: this.expr() }; }
      if (k.v === 'if') { this.take(); const c = this.expr(); this.want('kw', 'then'); const a = this.expr(); this.want('kw', 'else'); return { k: 'if', c, a, b: this.expr() }; }
      if (k.v === 'each') { this.take(); return { k: 'fn', params: [{ name: '_' }], body: this.expr(), each: true }; }
      if (k.v === 'try') { this.take(); const e = this.expr(); if (this.eat('kw', 'otherwise')) return { k: 'try', e, other: this.expr() }; return { k: 'try', e }; }
      if (k.v === 'error') { this.take(); return { k: 'error', e: this.expr() }; }
    }
    if (this.is('op', '(')) { const m = this.mark(); const fn = this.tryLambda(); if (fn) return fn; this.reset(m); }
    return this.coalesce();
  }
  name() { const t = this.take(); if (t.t === 'id') return t.v; if (t.t === 'kw') return t.v; throw this.lx.fail('Token Identifier expected.', t.p); }
  tryLambda() { try { this.want('op', '('); const params = []; if (!this.is('op', ')')) do { const opt = !!(this.is('id', 'optional') && this.take()); const t = this.take(); if (t.t !== 'id') return null; let ty = null; if (this.eat('kw', 'as')) ty = this.typeExpr(true); params.push({ name: t.v, optional: opt, type: ty }); } while (this.eat('op', ','));
      if (!this.eat('op', ')')) return null; let ret = null; if (this.eat('kw', 'as')) ret = this.typeExpr(true); if (!this.eat('op', '=>')) return null; return { k: 'fn', params, ret, body: this.expr() }; } catch (e) { return null; } }
  coalesce() { let a = this.or(); while (this.eat('op', '??')) a = { k: 'bin', op: '??', a, b: this.or() }; return a; }
  or() { let a = this.and(); while (this.eat('kw', 'or')) a = { k: 'bin', op: 'or', a, b: this.and() }; return a; }
  and() { let a = this.isas(); while (this.eat('kw', 'and')) a = { k: 'bin', op: 'and', a, b: this.isas() }; return a; }
  isas() { let a = this.eq(); for (;;) { if (this.eat('kw', 'is')) a = { k: 'is', e: a, t: this.typeExpr(true) }; else if (this.eat('kw', 'as')) a = { k: 'as', e: a, t: this.typeExpr(true) }; else return a; } }
  eq() { let a = this.rel(); for (;;) { const t = this.peek(); if (t.t === 'op' && (t.v === '=' || t.v === '<>')) { this.take(); a = { k: 'bin', op: t.v, a, b: this.rel() }; } else return a; } }
  rel() { let a = this.add(); for (;;) { const t = this.peek(); if (t.t === 'op' && ['<', '>', '<=', '>='].includes(t.v)) { this.take(); a = { k: 'bin', op: t.v, a, b: this.add() }; } else return a; } }
  add() { let a = this.mul(); for (;;) { const t = this.peek(); if (t.t === 'op' && ['+', '-', '&'].includes(t.v)) { this.take(); a = { k: 'bin', op: t.v, a, b: this.mul() }; } else return a; } }
  mul() { let a = this.meta(); for (;;) { const t = this.peek(); if (t.t === 'op' && (t.v === '*' || t.v === '/')) { this.take(); a = { k: 'bin', op: t.v, a, b: this.meta() }; } else return a; } }
  meta() { const a = this.unary(); if (this.eat('kw', 'meta')) { const m = this.unary(); return { k: 'meta', e: a, m }; } return a; }
  unary() { const t = this.peek(); if (t.t === 'op' && (t.v === '-' || t.v === '+')) { this.take(); return { k: 'un', op: t.v, a: this.unary() }; } if (t.t === 'kw' && t.v === 'not') { this.take(); return { k: 'un', op: 'not', a: this.unary() }; } if (t.t === 'kw' && t.v === 'type') { this.take(); return { k: 'type', t: this.typeExpr(false) }; } return this.postfix(this.primary()); }
  postfix(e) { for (;;) {
      if (this.is('op', '[')) { const p = this.peek(); if (this.looksRecord(p.p)) return e; this.take(); if (this.is('op', '[')) { const names = []; do { this.want('op', '['); names.push(this.genName()); this.want('op', ']'); } while (this.eat('op', ',')); this.want('op', ']'); const opt = !!this.eat('op', '?'); e = { k: 'proj', e, names, opt }; continue; }
        const n = this.genName(); this.want('op', ']'); const opt = !!this.eat('op', '?'); e = { k: 'field', e, name: n, opt }; continue; }
      if (this.is('op', '{')) { this.take(); const i = this.expr(); this.want('op', '}'); const opt = !!this.eat('op', '?'); e = { k: 'item', e, i, opt }; continue; }
      if (this.is('op', '(')) { this.take(); const args = []; if (!this.is('op', ')')) do args.push(this.expr()); while (this.eat('op', ',')); this.want('op', ')'); e = { k: 'call', f: e, args }; continue; }
      return e; } }
  // [A=1] is a record; [A] is a field of _ (or of the expression before it)
  looksRecord(at) { const s = this.src; let i = at + 1, d = 0; while (i < s.length) { const c = s[i];
      if (c === '"') { i++; while (i < s.length && !(s[i] === '"' && s[i + 1] !== '"')) i += s[i] === '"' ? 2 : 1; i++; continue; }
      if ('([{'.includes(c)) d++; else if (')]}'.includes(c)) { if (d === 0) return false; d--; }
      else if (c === '=' && d === 0 && !'<>='.includes(s[i - 1]) && s[i + 1] !== '>') return true;
      else if (c === ',' && d === 0) return false;
      i++; } return false; }
  genName() { if (this.buf) { this.lx.p = this.buf.start; this.buf = null; } this.lx.ws(); const s = this.src; if (s.startsWith('#"', this.lx.p)) { const r = this.lx.str(this.lx.p + 1); this.lx.p = r.end; return r.v; }
    let i = this.lx.p; while (i < s.length && !']=,'.includes(s[i])) i++; const v = s.slice(this.lx.p, i).trim(); if (!v) throw this.lx.fail('Token Identifier expected.'); this.lx.p = i; return v; }
  primary() { const t = this.peek();
    if (t.t === 'num') { this.take(); return { k: 'lit', v: t.v }; }
    if (t.t === 'str') { this.take(); return { k: 'lit', v: t.v }; }
    if (t.t === 'kw' && (t.v === 'true' || t.v === 'false')) { this.take(); return { k: 'lit', v: t.v === 'true' }; }
    if (t.t === 'kw' && t.v === 'null') { this.take(); return { k: 'lit', v: null }; }
    if (t.t === 'id') { this.take(); return { k: 'id', name: t.v, p: t.p }; }
    if (t.t === 'hash') { this.take(); if (t.v === '#infinity') return { k: 'lit', v: Infinity }; if (t.v === '#nan') return { k: 'lit', v: NaN }; if (t.v === '#shared' || t.v === '#sections') return { k: 'id', name: t.v };
      return { k: 'id', name: t.v, p: t.p }; }
    if (t.t === 'op' && t.v === '(') { this.take(); const e = this.expr(); this.want('op', ')'); return e; }
    if (t.t === 'op' && t.v === '{') { this.take(); const items = []; if (!this.is('op', '}')) do { const a = this.expr(); if (this.eat('op', '..')) items.push({ range: [a, this.expr()] }); else items.push(a); } while (this.eat('op', ',')); this.want('op', '}'); return { k: 'list', items }; }
    if (t.t === 'op' && t.v === '[') { if (this.looksRecord(t.p)) { this.take(); const fields = []; if (!this.is('op', ']')) do { const n = this.genName(); this.want('op', '='); fields.push([n, this.expr()]); } while (this.eat('op', ',')); this.want('op', ']'); return { k: 'rec', fields }; }
      this.take(); if (this.eat('op', ']')) return { k: 'rec', fields: [] }; const n = this.genName(); this.want('op', ']'); const opt = !!this.eat('op', '?'); return { k: 'field', e: { k: 'id', name: '_' }, name: n, opt }; }
    if (t.t === 'op' && t.v === '...') { this.take(); return { k: 'error', e: { k: 'lit', v: 'Not Implemented' } }; }
    throw this.lx.fail(t.t === 'eof' ? 'Token Literal expected.' : `Token Literal expected, but found '${t.v ?? t.t}'.`, t.p); }
  typeExpr(primaryOnly) { const t = this.peek();
    if (t.t === 'id' && t.v === 'nullable') { this.take(); const x = this.typeExpr(true); return { ...x, nullable: true }; }
    if (t.t === 'kw' && t.v === 'type') { this.take(); return this.typeExpr(primaryOnly); }
    if (t.t === 'id' && t.v === 'table') { this.take(); if (this.is('op', '[')) { this.take(); const cols = []; if (!this.is('op', ']')) do { const n = this.genName(); this.want('op', '='); cols.push([n, this.typeExpr(true)]); } while (this.eat('op', ',')); this.want('op', ']'); return { name: 'table', cols }; } return { name: 'table' }; }
    if (t.t === 'op' && t.v === '{') { this.take(); const it = this.typeExpr(true); this.want('op', '}'); return { name: 'list', item: it }; }
    if (t.t === 'op' && t.v === '[') { this.take(); const cols = []; if (!this.is('op', ']')) do { const n = this.genName(); this.want('op', '='); cols.push([n, this.typeExpr(true)]); } while (this.eat('op', ',')); this.want('op', ']'); return { name: 'record', cols }; }
    if (t.t === 'id' && t.v === 'function') { this.take(); this.want('op', '('); let d = 1; while (d) { const x = this.take(); if (x.t === 'op' && x.v === '(') d++; if (x.t === 'op' && x.v === ')') d--; if (x.t === 'eof') break; } if (this.eat('kw', 'as')) this.typeExpr(true); return { name: 'function' }; }
    if (t.t === 'id' || (t.t === 'kw' && (t.v === 'null'))) { this.take(); if (/\.Type$/.test(t.v)) return { expr: { k: 'id', name: t.v } }; return { name: t.v }; }
    if (!primaryOnly) return { expr: this.postfix(this.primary()) };
    throw this.lx.fail('Token Identifier expected.', t.p); }
}

// a section document (Formulas/Section1.m inside an Excel workbook's DataMashup) → {query name: M text}
// a workbook's Section1.m; Excel puts attribute records such as [ Description = … ] before some members
export function parseSection(src) { const P = new Parser(src); const out = {}; if (P.eat('kw', 'section')) { P.name(); P.want('op', ';'); }
  while (!P.is('eof')) { while (P.is('op', '[')) P.expr(); P.eat('kw', 'shared'); const n = P.name(); P.want('op', '='); const s = P.pos(); P.expr(); const e = P.peek().p; out[n] = src.slice(s, e).trim(); P.want('op', ';'); } return out; }

// ---------------------------------------------------------------- values
const force = t => { if (!t || !t.th) return t; if (t.busy) throw err('A cyclic reference was encountered during evaluation.'); if (!t.done) { t.busy = true; try { t.v = t.f(); t.ok = true; } catch (e) { t.e = e; } t.busy = false; t.done = true; } if (t.e) throw t.e; return t.v; };
const thunk = f => ({ th: true, f, done: false });
export const isTable = v => v instanceof MTable, isRec = v => v instanceof MRecord, isList = Array.isArray, isFn = v => v instanceof MFn;
const isNum = v => typeof v === 'number', isText = v => typeof v === 'string', isLog = v => typeof v === 'boolean';
const isDate = v => v instanceof MDate, isDT = v => v instanceof MDateTime, isTime = v => v instanceof MTime, isDur = v => v instanceof MDuration;
export function typeName(v) { if (v === null || v === undefined) return 'Null'; if (isNum(v)) return 'Number'; if (isText(v)) return 'Text'; if (isLog(v)) return 'Logical'; if (isDate(v)) return 'Date'; if (isDT(v)) return 'DateTime'; if (isTime(v)) return 'Time'; if (isDur(v)) return 'Duration'; if (isList(v)) return 'List'; if (isRec(v)) return 'Record'; if (isTable(v)) return 'Table'; if (isFn(v)) return 'Function'; if (v instanceof MType) return 'Type'; if (v instanceof MBinary) return 'Binary'; return 'Any'; }
const show = v => isText(v) ? v : isNum(v) || isLog(v) ? String(v) : v === null ? 'null' : `[${typeName(v)}]`;
const cantConvert = (v, to) => { const e = err(`We cannot convert the value ${isText(v) ? `"${v}"` : show(v)} to type ${to}.`); e.detail = { Value: v, Type: to }; return e; };
const needNum = v => { v = cell(v); if (!isNum(v)) throw cantConvert(v, 'Number'); return v; };
const needText = v => { v = cell(v); if (!isText(v)) throw cantConvert(v, 'Text'); return v; };
const needList = v => { v = cell(v); if (!isList(v)) throw cantConvert(v, 'List'); return v; };
const needTable = v => { v = cell(v); if (!isTable(v)) throw cantConvert(v, 'Table'); return v; };
const needFn = v => { v = cell(v); if (!isFn(v)) throw cantConvert(v, 'Function'); return v; };
const asCols = v => { v = cell(v); return isText(v) ? [v] : needList(v).map(needText); };
const asDateish = v => { v = cell(v); if (isDate(v) || isDT(v)) return v; throw cantConvert(v, 'Date'); };

// days ↔ y/m/d, on the same serials as the grid (0 = 30 Dec 1899)
const EPOCH = Date.UTC(1899, 11, 30);
export const ser = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 864e5;
export const ymd = n => { const dt = new Date(EPOCH + Math.floor(n) * 864e5); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), wd: dt.getUTCDay() }; };
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const CULT = { 'en-IN': { dmy: true, short: 'dd-MM-yyyy', dec: '.', grp: ',' }, 'en-GB': { dmy: true, short: 'dd/MM/yyyy', dec: '.', grp: ',' }, 'en-US': { dmy: false, short: 'M/d/yyyy', dec: '.', grp: ',' },
  'de-DE': { dmy: true, short: 'dd.MM.yyyy', dec: ',', grp: '.' }, 'fr-FR': { dmy: true, short: 'dd/MM/yyyy', dec: ',', grp: ' ' }, 'en-AU': { dmy: true, short: 'd/MM/yyyy', dec: '.', grp: ',' } };
const cult = c => CULT[c] || CULT['en-IN'];
const pad = (n, k = 2) => String(n).padStart(k, '0');
function fmtDate(n, f, c) { const { y, m, d, wd } = ymd(n); const frac = n - Math.floor(n); let secs = Math.round(frac * 86400); const H = Math.floor(secs / 3600), Mi = Math.floor(secs / 60) % 60, S = secs % 60;
  return f.replace(/yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|ss|tt|"[^"]*"|'[^']*'/g, t => ({ yyyy: y, yy: pad(y % 100), MMMM: MONTHS[m - 1], MMM: MONTHS[m - 1].slice(0, 3), MM: pad(m), M: m, dddd: DAYS[wd], ddd: DAYS[wd].slice(0, 3), dd: pad(d), d,
    HH: pad(H), H, hh: pad(H % 12 || 12), h: H % 12 || 12, mm: pad(Mi), ss: pad(S), tt: H < 12 ? 'AM' : 'PM' }[t] ?? t.slice(1, -1))); }
const MON3 = MONTHS.map(x => x.slice(0, 3).toLowerCase());
// text → serial (with a time fraction), or null when it can't be read; c is the culture
export function parseDateText(s, c) { s = String(s).trim(); const C = cult(c); let m, y, mo, d, rest = '';
  const yr = v => { v = +v; return String(v).length <= 2 && v < 100 ? (v <= 49 ? 2000 + v : 1900 + v) : v; };
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](.*))?$/.exec(s))) { y = +m[1]; mo = +m[2]; d = +m[3]; rest = m[4] || ''; }
  else if ((m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?:\s+(.*))?$/.exec(s))) { if (C.dmy) { d = +m[1]; mo = +m[2]; } else { mo = +m[1]; d = +m[2]; } y = yr(m[3]); rest = m[4] || ''; }
  else if ((m = /^(\d{1,2})[\s\-]([A-Za-z]{3,9})[\s\-,]+(\d{2,4})(?:\s+(.*))?$/.exec(s))) { d = +m[1]; mo = MON3.indexOf(m[2].slice(0, 3).toLowerCase()) + 1; y = yr(m[3]); rest = m[4] || ''; }
  else if ((m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{2,4})(?:\s+(.*))?$/.exec(s))) { mo = MON3.indexOf(m[1].slice(0, 3).toLowerCase()) + 1; d = +m[2]; y = yr(m[3]); rest = m[4] || ''; }
  else if ((m = /^([A-Za-z]{3,9})[\s\-,]+(\d{4})$/.exec(s))) { mo = MON3.indexOf(m[1].slice(0, 3).toLowerCase()) + 1; d = 1; y = +m[2]; }  // a month and a four-digit year, “Apr 2026”, is the 1st of that month
  else return null;
  if (!(mo >= 1 && mo <= 12 && d >= 1 && d <= 31)) return null; const n = ser(y, mo, d); if (ymd(n).d !== d) return null;
  if (!rest) return n; const t = parseTimeText(rest); return t === null ? null : n + t; }
function parseTimeText(s) { const m = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([AaPp][Mm])?$/.exec(String(s).trim()); if (!m) return null; let h = +m[1]; if (m[4]) { const pm = /p/i.test(m[4]); if (h === 12) h = 0; if (pm) h += 12; } if (h > 23 || +m[2] > 59) return null; return (h * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400; }
export function parseNumText(s, c) { const C = cult(c); let t = String(s).trim(); if (t === '') return null; let pct = false; if (t.endsWith('%')) { pct = true; t = t.slice(0, -1).trim(); }
  let neg = false; if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  t = t.split(C.grp).join('').replace(/ /g, ''); if (C.dec !== '.') t = t.replace(C.dec, '.');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return undefined; let n = Number(t); if (neg) n = -n; return pct ? n / 100 : n; }
const sig = n => Number.isInteger(n) ? n : Number(n.toPrecision(15));
export function numText(n, c) { if (Number.isNaN(n)) return 'NaN'; if (n === Infinity) return 'Infinity'; if (n === -Infinity) return '-Infinity'; let s = String(sig(n)); if (/e/.test(s)) s = sig(n).toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 20 }); const C = cult(c); return C.dec === '.' ? s : s.replace('.', C.dec); }
function durText(n) { const neg = n < 0; n = Math.abs(n); const d = Math.floor(n + 1e-9); let s = Math.round((n - d) * 86400); return `${neg ? '-' : ''}${d}.${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`; }
export function toText(v, c) { v = cell(v); if (v === null) return null; if (isText(v)) return v; if (isNum(v)) return numText(v, c); if (isLog(v)) return v ? 'true' : 'false';
  if (isDate(v)) return fmtDate(v.n, cult(c).short, c); if (isDT(v)) return fmtDate(v.n, cult(c).short + ' HH:mm:ss', c); if (isTime(v)) return fmtDate(v.n, 'HH:mm:ss', c); if (isDur(v)) return durText(v.n); throw cantConvert(v, 'Text'); }

// ---------------------------------------------------------------- operators
export function mEquals(a, b) { a = cell(a); b = cell(b);
  if (a === null || b === null) return a === b; if (isNum(a) && isNum(b)) return a === b || (Number.isNaN(a) && Number.isNaN(b)); if (isText(a) || isLog(a)) return a === b;
  if ((isDate(a) && isDate(b)) || (isDT(a) && isDT(b)) || (isTime(a) && isTime(b)) || (isDur(a) && isDur(b))) return Math.abs(a.n - b.n) < 1e-9;
  if (isList(a) && isList(b)) return a.length === b.length && a.every((x, i) => mEquals(x, b[i]));
  if (isRec(a) && isRec(b)) return a.names.length === b.names.length && a.names.every(n => b.idx(n) >= 0 && mEquals(a.get(n), b.get(n)));
  if (isTable(a) && isTable(b)) return a.cols.length === b.cols.length && a.cols.every((c, i) => c === b.cols[i]) && a.rows.length === b.rows.length && a.rows.every((r, i) => r.every((x, j) => mEquals(x, b.rows[i][j])));
  return a === b; }
const kindOf = v => isNum(v) ? 0 : isText(v) ? 1 : isLog(v) ? 2 : isDate(v) ? 3 : isDT(v) ? 4 : isTime(v) ? 5 : isDur(v) ? 6 : -1;
export function mCompare(a, b) { a = cell(a); b = cell(b); if (a === null && b === null) return 0; if (a === null) return -1; if (b === null) return 1;
  const ka = kindOf(a), kb = kindOf(b); if (ka !== kb || ka < 0) throw err(`We cannot apply operator < to types ${typeName(a)} and ${typeName(b)}.`);
  if (ka === 0) return a < b ? -1 : a > b ? 1 : 0; if (ka === 1) return a < b ? -1 : a > b ? 1 : 0; if (ka === 2) return (a ? 1 : 0) - (b ? 1 : 0); return a.n < b.n ? -1 : a.n > b.n ? 1 : 0; }
const opErr = (op, a, b) => err(`We cannot apply operator ${op} to types ${typeName(a)} and ${typeName(b)}.`);
function binop(op, a, b) {
  if (op === '=') return mEquals(a, b); if (op === '<>') return !mEquals(a, b);
  if (op === '<' || op === '>' || op === '<=' || op === '>=') { if (a === null || b === null) return null; const k = kindOf(a), k2 = kindOf(b); if (k !== k2 || k < 0) throw opErr(op, a, b); const c = mCompare(a, b); return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0; }
  if (a === null || b === null) { if (op === '&' && (isList(a) || isList(b) || isRec(a) || isRec(b))) throw opErr(op, a, b); return null; }
  if (op === '&') { if (isText(a) && isText(b)) return a + b; if (isList(a) && isList(b)) return a.concat(b); if (isRec(a) && isRec(b)) { const r = new MRecord(a.names.slice(), a.vals.slice()); b.names.forEach((n, i) => { const j = r.idx(n); if (j >= 0) r.vals[j] = b.vals[i]; else { r.names.push(n); r.vals.push(b.vals[i]); } }); return r; }
    if (isTable(a) && isTable(b)) return combine([a, b]); if (isDate(a) && isTime(b)) return new MDateTime(a.n + b.n); throw opErr(op, a, b); }
  if (isNum(a) && isNum(b)) return op === '+' ? a + b : op === '-' ? a - b : op === '*' ? a * b : a / b;
  if (op === '+' || op === '-') { const s = op === '+' ? 1 : -1;
    if (isDate(a) && isDur(b)) return Number.isInteger(b.n) ? new MDate(a.n + s * b.n) : new MDateTime(a.n + s * b.n); if (isDT(a) && isDur(b)) return new MDateTime(a.n + s * b.n); if (isTime(a) && isDur(b)) return new MTime(((a.n + s * b.n) % 1 + 1) % 1);
    if (isDur(a) && isDur(b)) return new MDuration(a.n + s * b.n); if (op === '+' && isDur(a) && (isDate(b) || isDT(b))) return binop('+', b, a);
    if (op === '-' && ((isDate(a) && isDate(b)) || (isDT(a) && isDT(b)) || (isTime(a) && isTime(b)))) return new MDuration(a.n - b.n); }
  if (op === '*' && isDur(a) && isNum(b)) return new MDuration(a.n * b); if (op === '*' && isNum(a) && isDur(b)) return new MDuration(a * b.n); if (op === '/' && isDur(a) && isNum(b)) return new MDuration(a.n / b);
  throw opErr(op, a, b); }

// ---------------------------------------------------------------- evaluator
class Env { constructor(vars, parent) { this.vars = vars; this.parent = parent; } }
function lookup(env, name) { for (let e = env; e; e = e.parent) { if (e.vars.has(name)) return force(e.vars.get(name)); if (e.resolve) { const v = e.resolve(name); if (v !== undefined) return v; } } throw err(`The name '${name}' wasn't recognized. Make sure it's spelled correctly.`); }
export function evaluate(n, env) {
  switch (n.k) {
    case 'lit': return n.v;
    case 'id': return lookup(env, n.name);
    case 'let': { const vars = new Map(); const e2 = new Env(vars, env); for (const b of n.binds) vars.set(b.name, thunk(() => evaluate(b.e, e2))); return evaluate(n.body, e2); }
    case 'if': { const c = cell(evaluate(n.c, env)); if (!isLog(c)) throw cantConvert(c, 'Logical'); return evaluate(c ? n.a : n.b, env); }
    case 'fn': { const f = new MFn(n.params, args => { const vars = new Map(); n.params.forEach((p, i) => vars.set(p.name, args[i] === undefined ? null : args[i])); return evaluate(n.body, new Env(vars, env)); }, null); f.ret = n.ret || null; return f; }
    case 'try': { try { const v = evaluate(n.e, env); force(v); if (n.other) return v; return MRecord.of({ HasError: false, Value: v }); } catch (e) { if (!(e instanceof MError)) throw e; if (n.other) return evaluate(n.other, env); return MRecord.of({ HasError: true, Error: MRecord.of({ Reason: e.reason, Message: e.message, Detail: e.detail }) }); } }
    case 'error': { const v = cell(evaluate(n.e, env)); if (isText(v)) throw err(v, 'Expression.Error'); if (isRec(v)) throw new MError(v.get('Reason', true) || 'Expression.Error', v.get('Message', true) || '', v.get('Detail', true)); throw err(show(v)); }
    case 'bin': {
      if (n.op === 'and') { const a = cell(evaluate(n.a, env)); if (a === false) return false; const b = cell(evaluate(n.b, env)); if (a === null) return b === false ? false : (isLog(b) || b === null ? null : (() => { throw opErr('and', a, b); })()); if (!isLog(a) || !(isLog(b) || b === null)) throw opErr('and', a, b); return b; }
      if (n.op === 'or') { const a = cell(evaluate(n.a, env)); if (a === true) return true; const b = cell(evaluate(n.b, env)); if (a === null) return b === true ? true : (isLog(b) || b === null ? null : (() => { throw opErr('or', a, b); })()); if (!isLog(a) || !(isLog(b) || b === null)) throw opErr('or', a, b); return b; }
      if (n.op === '??') { const a = cell(evaluate(n.a, env)); return a === null ? evaluate(n.b, env) : a; }
      return binop(n.op, cell(evaluate(n.a, env)), cell(evaluate(n.b, env))); }
    case 'un': { const a = cell(evaluate(n.a, env)); if (n.op === 'not') { if (a === null) return null; if (!isLog(a)) throw cantConvert(a, 'Logical'); return !a; } if (a === null) return null; if (isNum(a)) return n.op === '-' ? -a : a; if (isDur(a)) return new MDuration(n.op === '-' ? -a.n : a.n); throw err(`We cannot apply operator ${n.op} to type ${typeName(a)}.`); }
    case 'list': { const out = []; for (const it of n.items) { if (it.range) { const a = cell(evaluate(it.range[0], env)), b = cell(evaluate(it.range[1], env)); if (isNum(a) && isNum(b)) { for (let k = a; k <= b; k++) out.push(k); } else if (isText(a) && isText(b)) { for (let k = a.charCodeAt(0); k <= b.charCodeAt(0); k++) out.push(String.fromCharCode(k)); } else throw opErr('..', a, b); } else out.push(evaluate(it, env)); } return out; }
    case 'rec': { const names = [], vals = []; const e2 = new Env(new Map(), env); for (const [k, e] of n.fields) { names.push(k); const t = thunk(() => evaluate(e, e2)); vals.push(t); e2.vars.set(k, t); } return new MRecord(names, vals.map(t => { try { return force(t); } catch (e) { if (e instanceof MError) return new ErrCell(e); throw e; } })); }
    case 'field': { const v = cell(evaluate(n.e, env)); if (v === null && n.opt) return null; if (isRec(v)) return v.get(n.name, n.opt); if (isTable(v)) { const i = v.ci(n.name, n.opt); return i < 0 ? null : v.rows.map(r => r[i]); } if (n.opt) return null; throw err(`We cannot apply field access to the type ${typeName(v)}.`); }
    case 'proj': { const v = cell(evaluate(n.e, env)); if (isRec(v)) return new MRecord(n.names.slice(), n.names.map(x => v.get(x, n.opt))); if (isTable(v)) return selectColumns(v, n.names, n.opt ? 2 : 0); throw err(`We cannot apply field access to the type ${typeName(v)}.`); }
    case 'item': { const v = cell(evaluate(n.e, env)); const i = cell(evaluate(n.i, env)); if (v === null && n.opt) return null;
      if (isList(v)) { if (!isNum(i)) throw cantConvert(i, 'Number'); if (i < 0 || i >= v.length) { if (n.opt) return null; throw err('There weren\'t enough elements in the enumeration to complete the operation.'); } return cell(v[i]); }
      if (isTable(v)) { if (isNum(i)) { if (i < 0 || i >= v.rows.length) { if (n.opt) return null; throw err('There weren\'t enough elements in the enumeration to complete the operation.'); } return v.rec(i); }
        if (isRec(i)) { const hits = v.rows.map((r, k) => k).filter(k => i.names.every(f => mEquals(v.rows[k][v.ci(f)], i.get(f)))); if (hits.length === 1) return v.rec(hits[0]); if (n.opt && !hits.length) return null; throw err(hits.length ? 'The key matched more than one row in the table.' : 'The key didn\'t match any rows in the table.'); } }
      throw err(`We cannot apply indexing to the type ${typeName(v)}.`); }
    case 'call': { const f = cell(evaluate(n.f, env)); if (!isFn(f)) throw err(`We cannot convert a value of type ${typeName(f)} to type Function.`); const args = n.args.map(a => evaluate(a, env));
      const req = f.params ? f.params.filter(p => !p.optional).length : 0; if (f.params && !f.native && (args.length < req || args.length > f.params.length)) throw err(`${args.length} arguments were passed to a function which expects ${req === f.params.length ? req : `between ${req} and ${f.params.length}`}.`);
      return f.call(args); }
    case 'meta': return evaluate(n.e, env);
    case 'type': return mkType(n.t, env);
    case 'is': { const v = cell(evaluate(n.e, env)); return isType(v, mkType(n.t, env)); }
    case 'as': { const v = cell(evaluate(n.e, env)); if (!isType(v, mkType(n.t, env))) throw cantConvert(v, mkType(n.t, env).name); return v; }
  }
  throw err('Unsupported expression.');
}
const TYPE_ALIAS = { 'Int64.Type': 'Int64', 'Int32.Type': 'Int64', 'Int16.Type': 'Int64', 'Int8.Type': 'Int64', 'Currency.Type': 'Currency', 'Percentage.Type': 'Percentage', 'Decimal.Type': 'number', 'Double.Type': 'number', 'Single.Type': 'number', 'Byte.Type': 'Int64' };
function mkType(t, env) { if (t.expr) { const v = cell(evaluate(t.expr, env)); if (v instanceof MType) return t.nullable ? new MType(v.name, { ...v, nullable: true }) : v; throw cantConvert(v, 'Type'); } return new MType(t.name, { nullable: !!t.nullable, cols: t.cols ? t.cols.map(([n, x]) => [n, mkType(x, env)]) : undefined }); }
function isType(v, t) { const n = t.name; if (n === 'any') return true; if (n === 'anynonnull') return v !== null; if (v === null) return n === 'null' || !!t.nullable;
  return { number: isNum(v), Int64: isNum(v), Currency: isNum(v), Percentage: isNum(v), text: isText(v), logical: isLog(v), date: isDate(v), datetime: isDT(v), time: isTime(v), duration: isDur(v), list: isList(v), record: isRec(v), table: isTable(v), function: isFn(v), type: v instanceof MType, binary: v instanceof MBinary }[n] ?? false; }
const typeKey = t => t instanceof MType ? t.name : null;   // what a table column remembers: 'text', 'number', 'Int64' …

// ---------------------------------------------------------------- conversions (type changes, with a culture)
const roundEven = (x, d = 0) => { const f = 10 ** d; const y = Number((x * f).toPrecision(12)); const r = Math.round(y); const v = Math.abs(y % 1) === 0.5 ? (r % 2 === 0 ? r : r - 1) : r; return v / f; };
const roundAway = (x, d = 0) => { const f = 10 ** d; return Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f; };
export function convert(v, to, c) { v = cell(v); if (v === null) return null; if (!to || to === 'any') return v;
  if (to === 'text') return toText(v, c);
  if (to === 'number' || to === 'Int64' || to === 'Currency' || to === 'Percentage') { let n;
    if (isNum(v)) n = v; else if (isText(v)) { if (v.trim() === '') return null; n = parseNumText(v, c); if (n === undefined || n === null) throw dataErr(v, 'Number'); } else if (isLog(v)) n = v ? 1 : 0; else if (isDate(v) || isDT(v) || isTime(v) || isDur(v)) n = v.n; else throw dataErr(v, 'Number');
    return to === 'Int64' ? roundEven(n) : to === 'Currency' ? roundEven(n, 4) : n; }
  if (to === 'date') { if (isDate(v)) return v; if (isDT(v)) return new MDate(v.n); if (isNum(v)) return new MDate(v); if (isText(v)) { if (v.trim() === '') return null; const n = parseDateText(v, c); if (n === null) throw dataErr(v, 'Date'); return new MDate(n); } throw dataErr(v, 'Date'); }
  if (to === 'datetime') { if (isDT(v)) return v; if (isDate(v)) return new MDateTime(v.n); if (isNum(v)) return new MDateTime(v); if (isText(v)) { if (v.trim() === '') return null; const n = parseDateText(v, c); if (n === null) throw dataErr(v, 'DateTime'); return new MDateTime(n); } throw dataErr(v, 'DateTime'); }
  if (to === 'time') { if (isTime(v)) return v; if (isDT(v)) return new MTime(v.n % 1); if (isNum(v)) return new MTime(v % 1); if (isText(v)) { const n = parseTimeText(v); if (n === null) throw dataErr(v, 'Time'); return new MTime(n); } throw dataErr(v, 'Time'); }
  if (to === 'duration') { if (isDur(v)) return v; if (isNum(v)) return new MDuration(v); if (isText(v)) { const m = /^(-)?(?:(\d+)\.)?(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(v.trim()); if (!m) throw dataErr(v, 'Duration'); return new MDuration((m[1] ? -1 : 1) * (+(m[2] || 0) + (+m[3] * 3600 + +m[4] * 60 + +(m[5] || 0)) / 86400)); } throw dataErr(v, 'Duration'); }
  if (to === 'logical') { if (isLog(v)) return v; if (isNum(v)) return v !== 0; if (isText(v)) { const t = v.trim().toLowerCase(); if (t === 'true') return true; if (t === 'false') return false; } throw dataErr(v, 'Logical'); }
  return v; }
const dataErr = (v, to) => { const e = new MError('DataFormat.Error', `We couldn't convert to ${to}.`); e.detail = v; return e; };

// ---------------------------------------------------------------- table helpers
const wrap = f => { try { return f(); } catch (e) { if (e instanceof MError) return new ErrCell(e); throw e; } };
const call = (f, ...a) => needFn(f).call(a);
const missingMode = m => { m = cell(m); return m === null || m === undefined ? 0 : m; };
function selectColumns(t, cols, mode = 0) { cols = asCols(cols); const idx = []; const names = []; for (const c of cols) { const i = t.cols.indexOf(c); if (i < 0) { if (mode === 1) continue; if (mode === 2) { idx.push(-1); names.push(c); continue; } t.ci(c); } idx.push(i); names.push(c); }
  return new MTable(names, t.rows.map(r => idx.map(i => i < 0 ? null : r[i])), idx.map(i => i < 0 ? null : t.types[i])); }
function combine(ts) { const cols = []; for (const t of ts) for (const c of t.cols) if (!cols.includes(c)) cols.push(c);
  const types = cols.map(c => { const ks = ts.filter(t => t.cols.includes(c)).map(t => t.types[t.cols.indexOf(c)]); return ks.every(k => k === ks[0]) && ts.every(t => t.cols.includes(c)) ? ks[0] : null; });
  return new MTable(cols, ts.flatMap(t => t.rows.map(r => cols.map(c => { const i = t.cols.indexOf(c); return i < 0 ? null : r[i]; }))), types); }
const keyOf = (row, idx, ci) => idx.map(i => { const v = row[i]; if (v instanceof ErrCell) return '\u0000err'; const x = cell(v); return x === null ? '\u0000null' : (ci && isText(x) ? 't:' + x.toLowerCase() : kindOf(x) + ':' + (x && x.n !== undefined ? x.n : x)); }).join('\u0001');
const isCI = c => { c = cell(c); return !!(c && c.ci); };
function sortRows(t, crit) { crit = cell(crit); let list = isList(crit) && crit.length && !isList(cell(crit[0])) && isText(cell(crit[0])) && crit.length === 2 && isNum(cell(crit[1])) ? [crit] : isList(crit) ? crit : [crit];
  const keys = list.map(k => { k = cell(k); if (isText(k)) return { i: t.ci(k), d: 1 }; if (isList(k)) { const a = cell(k[0]); if (isFn(a)) return { f: a, d: cell(k[1]) === 1 ? -1 : 1 }; return { i: t.ci(needText(a)), d: cell(k[1]) === 1 ? -1 : 1 }; } if (isFn(k)) return { f: k, d: 1 }; throw cantConvert(k, 'List'); });
  const rows = t.rows.map((r, n) => ({ r, n })); rows.sort((x, y) => { for (const k of keys) { let a, b; if (k.f) { a = k.f.call([new MRecord(t.cols, x.r)]); b = k.f.call([new MRecord(t.cols, y.r)]); } else { a = x.r[k.i]; b = y.r[k.i]; } a = a instanceof ErrCell ? null : cell(a); b = b instanceof ErrCell ? null : cell(b); const c = mCompare(a, b); if (c) return c * k.d; } return x.n - y.n; });
  return new MTable(t.cols, rows.map(x => x.r), t.types); }
function csvParse(text, delim, quoteLines) { const rows = []; let row = [], f = '', q = false, i = 0; const s = String(text).replace(/^﻿/, '');
  while (i < s.length) { const c = s[i];
    if (q) { if (c === '"') { if (s[i + 1] === '"') { f += '"'; i += 2; continue; } q = false; i++; continue; } if ((c === '\n' || c === '\r') && !quoteLines) { q = false; continue; } f += c; i++; continue; }
    if (c === '"' && f === '') { q = true; i++; continue; }
    if (s.startsWith(delim, i)) { row.push(f); f = ''; i += delim.length; continue; }
    if (c === '\r' || c === '\n') { row.push(f); rows.push(row); row = []; f = ''; if (c === '\r' && s[i + 1] === '\n') i++; i++; continue; }
    f += c; i++; }
  if (f !== '' || row.length) { row.push(f); rows.push(row); } return rows; }
function splitText(s, delims, qs, fromEnd, once) { if (s === null) return [null]; s = needText(s); const parts = []; let cur = '', q = false, i = 0, used = 0;
  if (once) { const d = delims[0]; const k = fromEnd ? s.lastIndexOf(d) : s.indexOf(d); return k < 0 ? [s] : [s.slice(0, k), s.slice(k + d.length)]; }
  if (delims.length > 1 && !once) { for (let j = 0; j < delims.length; j++) { const k = s.indexOf(delims[j], i); if (k < 0) { parts.push(s.slice(i)); return parts; } parts.push(s.slice(i, k)); i = k + delims[j].length; } parts.push(s.slice(i)); return parts; }
  const d = delims[0]; if (d === '') return [s];
  while (i < s.length) { const c = s[i]; if (qs && c === '"' && cur === '') { q = !q; i++; continue; } if (q) { if (c === '"') { if (s[i + 1] === '"') { cur += '"'; i += 2; continue; } q = false; i++; continue; } cur += c; i++; continue; }
    if (s.startsWith(d, i)) { parts.push(cur); cur = ''; i += d.length; used++; continue; } cur += c; i++; }
  parts.push(cur); return parts; }

// ---------------------------------------------------------------- the standard library
const N = (name, f, arity) => { const fn = new MFn(null, args => f(...args), name); fn.native = true; fn.arity = arity; return fn; };
const enumRec = o => MRecord.of(o);
function buildLib(ctx) {
  const L = new Map(); const def = (n, f) => L.set(n, N(n, f)); const C = c => { c = cell(c); return isText(c) ? c : (isRec(c) ? c.get('Culture', true) : null) || ctx.culture; };
  const consts = { 'Order.Ascending': 0, 'Order.Descending': 1, 'JoinKind.Inner': 0, 'JoinKind.LeftOuter': 1, 'JoinKind.RightOuter': 2, 'JoinKind.FullOuter': 3, 'JoinKind.LeftAnti': 4, 'JoinKind.RightAnti': 5,
    'QuoteStyle.None': 0, 'QuoteStyle.Csv': 1, 'MissingField.Error': 0, 'MissingField.Ignore': 1, 'MissingField.UseNull': 2, 'Day.Sunday': 0, 'Day.Monday': 1, 'Day.Tuesday': 2, 'Day.Wednesday': 3, 'Day.Thursday': 4, 'Day.Friday': 5, 'Day.Saturday': 6,
    'RoundingMode.Up': 0, 'RoundingMode.Down': 1, 'RoundingMode.AwayFromZero': 2, 'RoundingMode.TowardZero': 3, 'RoundingMode.ToEven': 4, 'GroupKind.Global': 1, 'GroupKind.Local': 0, 'ExtraValues.List': 0, 'ExtraValues.Error': 1, 'ExtraValues.Ignore': 2,
    'Occurrence.First': 0, 'Occurrence.Last': 1, 'Occurrence.All': 2, 'RelativePosition.FromStart': 0, 'RelativePosition.FromEnd': 1, 'TextEncoding.Utf8': 65001, 'TextEncoding.Windows': 1252, 'JoinAlgorithm.Dynamic': 0, 'Precision.Double': 0, 'Precision.Decimal': 1 };
  for (const [k, v] of Object.entries(consts)) L.set(k, v);
  for (const [k, v] of Object.entries(TYPE_ALIAS)) L.set(k, new MType(v));
  const ci = N('Comparer.OrdinalIgnoreCase', (a, b) => { a = toText(a) ?? ''; b = toText(b) ?? ''; const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }); ci.ci = true; L.set('Comparer.OrdinalIgnoreCase', ci);
  L.set('Comparer.Ordinal', N('Comparer.Ordinal', (a, b) => mCompare(a, b))); def('Comparer.FromCulture', () => ci);
  const eqf = (c) => isCI(c) ? (a, b) => { a = cell(a); b = cell(b); return isText(a) && isText(b) ? a.toLowerCase() === b.toLowerCase() : mEquals(a, b); } : mEquals;
  // sources
  def('Excel.CurrentWorkbook', () => { const names = Object.keys(ctx.tables || {}); return new MTable(['Content', 'Name'], names.map(n => [toMTable(ctx.tables[n]), n]), ['table', 'text']); });
  def('File.Contents', p => { p = needText(p); const f = ctx.file(p); if (!f) throw new MError('DataSource.Error', `Could not find file '${p}'.`); return f; });
  def('Folder.Files', p => { p = needText(p); const list = ctx.folder(p); if (!list) throw new MError('DataSource.Error', `Could not find a part of the path '${p}'.`);
    return new MTable(['Content', 'Name', 'Extension', 'Date accessed', 'Date modified', 'Date created', 'Attributes', 'Folder Path'], list.map(f => [f.bin, f.name, (/\.[^.]+$/.exec(f.name) || [''])[0].toLowerCase(), new MDateTime(f.modified), new MDateTime(f.modified), new MDateTime(f.modified), MRecord.of({ 'Content Type': 'text/csv', Kind: 'Text File', Size: f.bin.text ? f.bin.text.length : 0, ReadOnly: false, Hidden: !!f.hidden, System: false, Directory: false, Archive: true, Device: false, Normal: false, Temporary: false }), f.folder]),
      ['binary', 'text', 'text', 'datetime', 'datetime', 'datetime', 'record', 'text']); });
  def('Folder.Contents', p => L.get('Folder.Files').call([p]));
  def('Csv.Document', (src, opts, delim, extra, enc) => { src = cell(src); const text = src instanceof MBinary ? src.text : isText(src) ? src : (() => { throw cantConvert(src, 'Binary'); })(); opts = cell(opts);
    let d = ',', cols = null, qs = 0; if (isRec(opts)) { d = opts.get('Delimiter', true) ?? ','; cols = opts.get('Columns', true); qs = opts.get('QuoteStyle', true) ?? 0; } else if (isNum(opts)) cols = opts; else if (isList(opts)) cols = opts; if (cell(delim)) d = cell(delim);
    const rows = csvParse(text, d, qs === 1); const n = isList(cols) ? cols.length : isNum(cols) ? cols : Math.max(0, ...rows.map(r => r.length)); const names = isList(cols) ? cols.map(cell) : Array.from({ length: n }, (_, i) => `Column${i + 1}`);
    return new MTable(names, rows.map(r => names.map((_, i) => r[i] ?? null)), names.map(() => null)); });
  def('Text.FromBinary', b => { b = cell(b); return b instanceof MBinary ? b.text : needText(b); }); def('Lines.FromBinary', b => String(cell(b).text).split(/\r?\n/));
  def('Excel.Workbook', (b) => { b = cell(b); const w = b && b.extra && b.extra.workbook; if (!w) throw new MError('DataFormat.Error', 'External table is not in the expected format.'); const rows = [];
    for (const [n, s] of Object.entries(w.sheets || {})) rows.push([n, toMTable(s, true), n, 'Sheet', false]); for (const [n, s] of Object.entries(w.tables || {})) rows.push([n, toMTable(s), n, 'Table', false]); return new MTable(['Name', 'Data', 'Item', 'Kind', 'Hidden'], rows, ['text', 'table', 'text', 'text', 'logical']); });
  def('Web.Contents', () => { throw new MError('DataSource.Error', 'Web sources aren’t available in this practice editor.'); });
  // building tables
  L.set('#table', N('#table', (cols, rows) => { cols = cell(cols); rows = needList(rows); let names, types; if (cols instanceof MType) { names = (cols.cols || []).map(x => x[0]); types = (cols.cols || []).map(x => typeKey(x[1])); } else if (isNum(cols)) { names = Array.from({ length: cols }, (_, i) => `Column${i + 1}`); } else names = asCols(cols); return new MTable(names, rows.map(r => { r = needList(r); return names.map((_, i) => r[i] ?? null); }), types); }));
  def('Table.FromRows', (rows, cols) => L.get('#table').call([cell(cols) ?? Math.max(0, ...needList(rows).map(r => needList(r).length)), rows]));
  def('Table.FromRecords', (recs, cols) => { recs = needList(recs).map(cell); let names = cell(cols) ? asCols(cols) : []; if (!names.length) for (const r of recs) for (const n of r.names) if (!names.includes(n)) names.push(n); return new MTable(names, recs.map(r => names.map(n => r.idx(n) < 0 ? null : r.vals[r.idx(n)]))); });
  def('Table.FromList', (list, splitter, cols) => { list = needList(list); splitter = cell(splitter); const parts = list.map(x => splitter && !splitter.nothing ? needList(splitter.call([x])) : [x]); cols = cell(cols); const n = isList(cols) ? cols.length : isNum(cols) ? cols : Math.max(1, ...parts.map(p => p.length)); const names = isList(cols) ? cols.map(cell) : Array.from({ length: n }, (_, i) => `Column${i + 1}`); return new MTable(names, parts.map(p => names.map((_, i) => p[i] ?? null))); });
  const nothing = N('Splitter.SplitByNothing', x => [x]); nothing.nothing = true; def('Splitter.SplitByNothing', () => nothing);
  def('Table.ToRows', t => needTable(t).rows.map(r => r.slice())); def('Table.ToRecords', t => { t = needTable(t); return t.rows.map((r, i) => t.rec(i)); }); def('Table.ToList', (t, comb) => { t = needTable(t); return t.rows.map(r => comb ? call(comb, r) : (t.cols.length === 1 ? r[0] : r.map(x => toText(x)).join(','))); });
  def('Table.Column', (t, c) => needTable(t).col(needText(c))); def('Table.ColumnNames', t => needTable(t).cols.slice()); def('Table.RowCount', t => needTable(t).rows.length); def('Table.ColumnCount', t => needTable(t).cols.length);
  def('Table.IsEmpty', t => needTable(t).rows.length === 0); def('Table.Buffer', t => needTable(t)); def('Table.HasColumns', (t, c) => asCols(c).every(x => needTable(t).cols.includes(x)));
  def('Table.First', (t, d) => { t = needTable(t); return t.rows.length ? t.rec(0) : (cell(d) ?? null); }); def('Table.Last', (t, d) => { t = needTable(t); return t.rows.length ? t.rec(t.rows.length - 1) : (cell(d) ?? null); });
  def('Table.FirstValue', (t, d) => { t = needTable(t); return t.rows.length && t.cols.length ? cell(t.rows[0][0]) : (cell(d) ?? null); });
  def('Table.Max', (t, c, d) => { t = needTable(t); const i = t.ci(needText(c)); let best = -1; t.rows.forEach((r, k) => { if (cell(r[i]) !== null && (best < 0 || mCompare(r[i], t.rows[best][i]) > 0)) best = k; }); return best < 0 ? (cell(d) ?? null) : t.rec(best); });
  def('Table.Min', (t, c, d) => { t = needTable(t); const i = t.ci(needText(c)); let best = -1; t.rows.forEach((r, k) => { if (cell(r[i]) !== null && (best < 0 || mCompare(r[i], t.rows[best][i]) < 0)) best = k; }); return best < 0 ? (cell(d) ?? null) : t.rec(best); });
  // headers
  def('Table.PromoteHeaders', (t, o) => { t = needTable(t); if (!t.rows.length) return t; const seen = {}; const names = t.rows[0].map((v, i) => { let n = v instanceof ErrCell ? `Column${i + 1}` : toText(cell(v)) ?? ''; if (n === '') n = `Column${i + 1}`; if (seen[n]) { let k = 1; while (seen[n + '_' + k]) k++; n = n + '_' + k; } seen[n] = 1; return n; }); return new MTable(names, t.rows.slice(1), names.map(() => null)); });
  def('Table.DemoteHeaders', t => { t = needTable(t); const names = t.cols.map((_, i) => `Column${i + 1}`); return new MTable(names, [t.cols.slice(), ...t.rows], names.map(() => null)); });
  def('Table.Transpose', (t, cols) => { t = needTable(t); const n = t.rows.length; const names = cell(cols) ? asCols(cols) : Array.from({ length: n }, (_, i) => `Column${i + 1}`); return new MTable(names, t.cols.map((_, j) => names.map((_, i) => t.rows[i] ? t.rows[i][j] : null))); });
  // rows
  const countOrCond = (t, x, fromStart) => { x = cell(x); if (isFn(x)) { let k = 0; const rows = fromStart ? t.rows : t.rows.slice().reverse(); while (k < rows.length && x.call([new MRecord(t.cols, rows[k])]) === true) k++; return k; } return x === null || x === undefined ? 1 : needNum(x); };
  def('Table.Skip', (t, x) => { t = needTable(t); const k = countOrCond(t, x, true); return new MTable(t.cols, t.rows.slice(k), t.types); });
  def('Table.RemoveFirstN', (t, x) => L.get('Table.Skip').call([t, x]));
  def('Table.FirstN', (t, x) => { t = needTable(t); const k = countOrCond(t, x, true); return new MTable(t.cols, t.rows.slice(0, k), t.types); });
  def('Table.LastN', (t, x) => { t = needTable(t); const k = countOrCond(t, x, false); return new MTable(t.cols, k ? t.rows.slice(-k) : [], t.types); });
  def('Table.RemoveLastN', (t, x) => { t = needTable(t); const k = countOrCond(t, x, false); return new MTable(t.cols, t.rows.slice(0, Math.max(0, t.rows.length - k)), t.types); });
  def('Table.Range', (t, o, n) => { t = needTable(t); o = needNum(o); n = cell(n); return new MTable(t.cols, t.rows.slice(o, n === null || n === undefined ? undefined : o + n), t.types); });
  def('Table.RemoveRows', (t, o, n) => { t = needTable(t); o = needNum(o); n = cell(n) ?? 1; const rows = t.rows.slice(); rows.splice(o, n); return new MTable(t.cols, rows, t.types); });
  def('Table.ReverseRows', t => { t = needTable(t); return new MTable(t.cols, t.rows.slice().reverse(), t.types); });
  def('Table.SelectRows', (t, f) => { t = needTable(t); f = needFn(f); return new MTable(t.cols, t.rows.filter(r => { const v = cell(f.call([new MRecord(t.cols, r)])); if (v === null) return false; if (!isLog(v)) throw cantConvert(v, 'Logical'); return v; }), t.types); });
  def('Table.Distinct', (t, cols) => { t = needTable(t); cols = cell(cols); let idx = t.cols.map((_, i) => i), cmp = null; if (isText(cols) || (isList(cols) && cols.length && isText(cell(cols[0])))) idx = asCols(cols).map(c => t.ci(c)); else if (isList(cols) && cols.length === 2 && isFn(cell(cols[1]))) { idx = asCols(cols[0]).map(c => t.ci(c)); cmp = cols[1]; } else if (isFn(cols)) cmp = cols;
    const seen = new Set(); return new MTable(t.cols, t.rows.filter(r => { const k = keyOf(r, idx, isCI(cmp)); if (seen.has(k)) return false; seen.add(k); return true; }), t.types); });
  const errRows = (t, cols, keep) => { t = needTable(t); const idx = cell(cols) ? asCols(cols).map(c => t.ci(c)) : t.cols.map((_, i) => i); return new MTable(t.cols, t.rows.filter(r => idx.some(i => r[i] instanceof ErrCell) === keep), t.types); };
  def('Table.RemoveRowsWithErrors', (t, cols) => errRows(t, cols, false)); def('Table.SelectRowsWithErrors', (t, cols) => errRows(t, cols, true));
  def('Table.ReplaceErrorValues', (t, pairs) => { t = needTable(t); const ps = needList(pairs).map(p => { p = needList(p); return [t.ci(needText(p[0])), cell(p[1])]; }); return new MTable(t.cols, t.rows.map(r => { const x = r.slice(); for (const [i, v] of ps) if (x[i] instanceof ErrCell) x[i] = v; return x; }), t.types); });
  def('Table.Sort', (t, crit) => sortRows(needTable(t), crit));
  const fill = (t, cols, up) => { t = needTable(t); const idx = asCols(cols).map(c => t.ci(c)); const rows = t.rows.map(r => r.slice()); const order = up ? rows.slice().reverse() : rows; for (const i of idx) { let last = null; for (const r of order) { if (r[i] === null) r[i] = last; else last = r[i]; } } return new MTable(t.cols, rows, t.types); };
  def('Table.FillDown', (t, c) => fill(t, c, false)); def('Table.FillUp', (t, c) => fill(t, c, true));
  def('Table.InsertRows', (t, o, recs) => { t = needTable(t); const rows = t.rows.slice(); rows.splice(needNum(o), 0, ...needList(recs).map(r => { r = cell(r); return t.cols.map(c => r.get(c, true)); })); return new MTable(t.cols, rows, t.types); });
  def('Table.Repeat', (t, n) => { t = needTable(t); const rows = []; for (let k = 0; k < needNum(n); k++) rows.push(...t.rows); return new MTable(t.cols, rows, t.types); });
  // columns
  def('Table.SelectColumns', (t, c, m) => selectColumns(needTable(t), c, missingMode(m)));
  def('Table.RemoveColumns', (t, c, m) => { t = needTable(t); const drop = asCols(c); if (!missingMode(m)) drop.forEach(x => t.ci(x)); return selectColumns(t, t.cols.filter(x => !drop.includes(x))); });
  def('Table.ReorderColumns', (t, c, m) => { t = needTable(t); const want = asCols(c).filter(x => !missingMode(m) || t.cols.includes(x)); want.forEach(x => t.ci(x)); const rest = t.cols.slice(); const slots = rest.map((x, i) => want.includes(x) ? i : -1).filter(i => i >= 0); const out = rest.slice(); slots.forEach((s, k) => { out[s] = want[k]; }); return selectColumns(t, out); });
  def('Table.RenameColumns', (t, pairs, m) => { t = needTable(t); let ps = needList(pairs); if (ps.length === 2 && isText(cell(ps[0]))) ps = [ps]; const names = t.cols.slice(); for (let p of ps) { p = needList(p); const a = needText(p[0]), b = needText(p[1]); const i = names.indexOf(a); if (i < 0) { if (missingMode(m)) continue; t.ci(a); } if (names.includes(b) && names.indexOf(b) !== i) throw err(`The field '${b}' already exists in the record.`); names[i] = b; } return new MTable(names, t.rows, t.types); });
  def('Table.TransformColumnNames', (t, f) => { t = needTable(t); return new MTable(t.cols.map(c => needText(call(f, c))), t.rows, t.types); });
  def('Table.PrefixColumns', (t, p) => { t = needTable(t); return new MTable(t.cols.map(c => needText(p) + '.' + c), t.rows, t.types); });
  def('Table.DuplicateColumn', (t, c, n, ty) => { t = needTable(t); const i = t.ci(needText(c)); return new MTable([...t.cols, needText(n)], t.rows.map(r => [...r, r[i]]), [...t.types, cell(ty) ? typeKey(cell(ty)) : t.types[i]]); });
  def('Table.AddColumn', (t, n, f, ty) => { t = needTable(t); n = needText(n); if (t.cols.includes(n)) throw err(`The field '${n}' already exists in the record.`); f = needFn(f); return new MTable([...t.cols, n], t.rows.map(r => [...r, wrap(() => { const v = f.call([new MRecord(t.cols, r)]); return cell(v); })]), [...t.types, cell(ty) ? typeKey(cell(ty)) : null]); });
  def('Table.AddIndexColumn', (t, n, s, st, ty) => { t = needTable(t); s = cell(s) ?? 0; st = cell(st) ?? 1; return new MTable([...t.cols, needText(n)], t.rows.map((r, i) => [...r, s + i * st]), [...t.types, cell(ty) ? typeKey(cell(ty)) : 'Int64']); });
  def('Table.TransformColumns', (t, ops, dflt, m) => { t = needTable(t); let list = needList(ops); if (list.length >= 2 && isText(cell(list[0]))) list = [list]; const rows = t.rows.map(r => r.slice()); const types = t.types.slice(); const done = new Set();
    for (let op of list) { op = needList(op); const c = needText(op[0]); const i = t.cols.indexOf(c); if (i < 0) { if (missingMode(m)) continue; t.ci(c); } const f = needFn(op[1]); done.add(i); for (const r of rows) r[i] = wrap(() => cell(f.call([cell(r[i])]))); if (op[2] !== undefined && cell(op[2])) types[i] = typeKey(cell(op[2])); }
    if (cell(dflt)) { const f = needFn(dflt); t.cols.forEach((_, i) => { if (!done.has(i)) for (const r of rows) r[i] = wrap(() => cell(f.call([cell(r[i])]))); }); }
    return new MTable(t.cols, rows, types); });
  def('Table.TransformColumnTypes', (t, ops, cu) => { t = needTable(t); let list = needList(ops); if (list.length === 2 && isText(cell(list[0]))) list = [list]; const c = C(cu); const rows = t.rows.map(r => r.slice()); const types = t.types.slice();
    for (let op of list) { op = needList(op); const i = t.ci(needText(op[0])); const ty = cell(op[1]); const to = typeKey(ty); for (const r of rows) { if (r[i] instanceof ErrCell) continue; r[i] = wrap(() => convert(r[i], to, c)); } types[i] = to; }
    return new MTable(t.cols, rows, types); });
  def('Table.ReplaceValue', (t, old, nw, rep, cols) => { t = needTable(t); const idx = asCols(cols).map(c => t.ci(c)); rep = needFn(rep); return new MTable(t.cols, t.rows.map(r => { const x = r.slice(); for (const i of idx) { if (x[i] instanceof ErrCell) continue; const row = new MRecord(t.cols, r); x[i] = wrap(() => { const o = isFn(cell(old)) ? cell(old).call([row]) : old; const w = isFn(cell(nw)) ? cell(nw).call([row]) : nw; return rep.call([x[i], o, w]); }); } return x; }), t.types); });
  def('Replacer.ReplaceText', (v, o, w) => { v = cell(v); if (v === null) return null; if (!isText(v)) return v; o = needText(o); if (o === '') return v; return v.split(o).join(needText(w)); });
  def('Replacer.ReplaceValue', (v, o, w) => mEquals(v, o) ? cell(w) : cell(v));
  def('Table.SplitColumn', (t, c, sp, names, dflt, extra) => { t = needTable(t); c = needText(c); const i = t.ci(c); sp = needFn(sp); const parts = t.rows.map(r => r[i] instanceof ErrCell ? [r[i]] : wrap(() => needList(sp.call([cell(r[i])])))); names = cell(names);
    const n = isList(names) ? names.length : isNum(names) ? names : Math.max(1, ...parts.map(p => p instanceof ErrCell ? 1 : p.length)); const newNames = isList(names) ? names.map(cell) : Array.from({ length: n }, (_, k) => `${c}.${k + 1}`);
    const cols = [...t.cols.slice(0, i), ...newNames, ...t.cols.slice(i + 1)]; const types = [...t.types.slice(0, i), ...newNames.map(() => null), ...t.types.slice(i + 1)];
    return new MTable(cols, t.rows.map((r, k) => { const p = parts[k]; const vals = p instanceof ErrCell ? newNames.map(() => p) : newNames.map((_, j) => p[j] ?? null); return [...r.slice(0, i), ...vals, ...r.slice(i + 1)]; }), types); });
  def('Splitter.SplitTextByDelimiter', (d, qs) => N('Splitter', s => splitText(s, [needText(d)], cell(qs) === 1)));
  def('Splitter.SplitTextByEachDelimiter', (ds, qs, fromEnd) => N('Splitter', s => { const list = needList(ds).map(needText); return list.length === 1 ? splitText(s, list, cell(qs) === 1, cell(fromEnd) === true, true) : splitText(s, list, false); }));
  def('Splitter.SplitTextByAnyDelimiter', (ds) => N('Splitter', s => { if (cell(s) === null) return [null]; s = needText(s); const list = needList(ds).map(needText); const re = new RegExp(list.map(x => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')); return s.split(re); }));
  def('Splitter.SplitTextByPositions', (ps, fromEnd) => N('Splitter', s => { if (cell(s) === null) return [null]; s = needText(s); let p = needList(ps).map(needNum); if (cell(fromEnd) === true) { const b = p.map(x => Math.max(0, s.length - x)); return b.map((x, k) => s.slice(k + 1 < b.length ? b[k + 1] : 0, x)).reverse(); } return p.map((x, k) => s.slice(x, p[k + 1])); }));
  def('Splitter.SplitTextByRepeatedLengths', (n) => N('Splitter', s => { if (cell(s) === null) return [null]; s = needText(s); n = needNum(n); const out = []; for (let k = 0; k < s.length; k += n) out.push(s.slice(k, k + n)); return out.length ? out : ['']; }));
  def('Splitter.SplitTextByLengths', (ls) => N('Splitter', s => { if (cell(s) === null) return [null]; s = needText(s); let o = 0; return needList(ls).map(needNum).map(n => { const x = s.slice(o, o + n); o += n; return x; }); }));
  def('Splitter.SplitTextByCharacterTransition', (a, b) => N('Splitter', s => { if (cell(s) === null) return [null]; s = needText(s); const test = (x, ch) => { x = cell(x); return isFn(x) ? x.call([ch]) === true : needList(x).map(cell).includes(ch); }; const out = []; let cur = ''; for (let k = 0; k < s.length; k++) { if (k > 0 && test(a, s[k - 1]) && test(b, s[k])) { out.push(cur); cur = ''; } cur += s[k]; } out.push(cur); return out; }));
  def('Table.CombineColumns', (t, cols, comb, n) => { t = needTable(t); const cs = asCols(cols); const idx = cs.map(c => t.ci(c)); comb = needFn(comb); const first = Math.min(...idx); const keep = t.cols.map((_, i) => i).filter(i => !idx.includes(i)); const at = keep.filter(i => i < first).length;
    const names = keep.map(i => t.cols[i]); names.splice(at, 0, needText(n)); const types = keep.map(i => t.types[i]); types.splice(at, 0, 'text');
    return new MTable(names, t.rows.map(r => { const x = keep.map(i => r[i]); x.splice(at, 0, wrap(() => comb.call([idx.map(i => cell(r[i]))]))); return x; }), types); });
  def('Combiner.CombineTextByDelimiter', (d, qs) => N('Combiner', list => needList(list).map(v => cell(v)).filter(v => v !== null).map(v => toText(v)).join(needText(d))));
  def('Table.Unpivot', (t, cols, a, v) => { t = needTable(t); const un = asCols(cols); un.forEach(c => t.ci(c)); return unpivot(t, t.cols.filter(c => !un.includes(c)), needText(a), needText(v)); });
  def('Table.UnpivotOtherColumns', (t, keep, a, v) => { t = needTable(t); const k = asCols(keep); k.forEach(c => t.ci(c)); return unpivot(t, k, needText(a), needText(v)); });
  function unpivot(t, keep, a, v) { const ki = keep.map(c => t.ci(c)); const ui = t.cols.map((_, i) => i).filter(i => !ki.includes(i)); const ut = ui.map(i => t.types[i]); const rows = [];
    for (const r of t.rows) for (const i of ui) { if (cell0(r[i]) === null) continue; rows.push([...ki.map(j => r[j]), t.cols[i], r[i]]); }
    return new MTable([...keep, a, v], rows, [...ki.map(j => t.types[j]), 'text', ut.every(x => x === ut[0]) ? ut[0] : null]); }
  def('Table.Pivot', (t, vals, ac, vc, agg) => { t = needTable(t); const ai = t.ci(needText(ac)), vi = t.ci(needText(vc)); const names = needList(vals).map(x => toText(x)); const keepIdx = t.cols.map((_, i) => i).filter(i => i !== ai && i !== vi); agg = cell(agg);
    const groups = new Map(); for (const r of t.rows) { const k = keyOf(r, keepIdx); if (!groups.has(k)) groups.set(k, { key: keepIdx.map(i => r[i]), cells: names.map(() => []) }); const j = names.indexOf(toText(r[ai])); if (j >= 0) groups.get(k).cells[j].push(r[vi]); }
    const rows = [...groups.values()].map(g => [...g.key, ...g.cells.map(list => { if (agg) return wrap(() => cell(agg.call([list.map(cell)]))); if (list.length > 1) return new ErrCell(err('There were too many elements in the enumeration to complete the operation.')); return list.length ? list[0] : null; })]);
    return new MTable([...keepIdx.map(i => t.cols[i]), ...names], rows, [...keepIdx.map(i => t.types[i]), ...names.map(() => agg ? null : t.types[vi])]); });
  def('Table.Group', (t, keys, aggs, kind, cmp) => { t = needTable(t); const ks = asCols(keys); const ki = ks.map(c => t.ci(c)); let al = needList(aggs); if (al.length >= 2 && isText(cell(al[0]))) al = [al]; al = al.map(a => needList(a)); const local = cell(kind) === 0; const ic = isCI(cmp);
    const groups = []; const map = new Map(); let lastKey = null; for (const r of t.rows) { const k = keyOf(r, ki, ic); if (local) { if (k !== lastKey) groups.push({ key: ki.map(i => r[i]), rows: [] }); groups[groups.length - 1].rows.push(r); lastKey = k; } else { if (!map.has(k)) { const g = { key: ki.map(i => r[i]), rows: [] }; map.set(k, g); groups.push(g); } map.get(k).rows.push(r); } }
    return new MTable([...ks, ...al.map(a => needText(a[0]))], groups.map(g => [...g.key, ...al.map(a => wrap(() => cell(needFn(a[1]).call([new MTable(t.cols, g.rows, t.types)]))))]), [...ki.map(i => t.types[i]), ...al.map(a => a[2] !== undefined && cell(a[2]) ? typeKey(cell(a[2])) : null)]); });
  def('Table.Combine', (ts, cols) => { const list = needList(ts).map(needTable); const out = combine(list); return cell(cols) ? selectColumns(out, cols, 2) : out; });
  const joinRows = (a, ka, b, kb, kind, ic) => { const ai = asCols(ka).map(c => a.ci(c)), bi = asCols(kb).map(c => b.ci(c)); const index = new Map(); b.rows.forEach((r, k) => { const key = keyOf(r, bi, ic); if (!index.has(key)) index.set(key, []); index.get(key).push(k); }); const used = new Set(); const out = [];
    a.rows.forEach(r => { const key = ai.some(i => cell0(r[i]) === null) ? null : keyOf(r, ai, ic); const hits = key === null ? [] : (index.get(key) || []); hits.forEach(h => used.add(h)); out.push({ r, hits }); }); return { out, used }; };
  def('Table.NestedJoin', (a, ka, b, kb, n, kind, cmp) => { a = needTable(a); b = needTable(b); kind = cell(kind) ?? 1; const { out, used } = joinRows(a, ka, b, kb, kind, isCI(cmp)); const sub = hits => new MTable(b.cols, hits.map(h => b.rows[h]), b.types); let rows;
    if (kind === 0) rows = out.filter(x => x.hits.length).map(x => [...x.r, sub(x.hits)]); else if (kind === 1 || kind === 3) rows = out.map(x => [...x.r, sub(x.hits)]); else if (kind === 4) rows = out.filter(x => !x.hits.length).map(x => [...x.r, sub([])]); else rows = [];
    if (kind === 2 || kind === 3 || kind === 5) { const right = b.rows.map((r, k) => k).filter(k => (kind === 5 ? !used.has(k) : true)); if (kind === 2) rows = out.filter(x => x.hits.length).map(x => [...x.r, sub(x.hits)]); for (const k of right) if (!used.has(k) || kind === 5) rows.push([...a.cols.map(() => null), sub([k])]); if (kind === 5) rows = rows.filter(r => r[a.cols.length].rows.length); }
    return new MTable([...a.cols, needText(n)], rows, [...a.types, 'table']); });
  def('Table.Join', (a, ka, b, kb, kind) => { a = needTable(a); b = needTable(b); kind = cell(kind) ?? 0; const { out, used } = joinRows(a, ka, b, kb, kind); const dup = b.cols.find(c => a.cols.includes(c)); if (dup) throw err(`A join operation cannot result in a table with duplicate column names ("${dup}").`); const rows = [];
    for (const x of out) { if (x.hits.length) for (const h of x.hits) rows.push([...x.r, ...b.rows[h]]); else if (kind === 1 || kind === 3) rows.push([...x.r, ...b.cols.map(() => null)]); }
    if (kind === 2 || kind === 3) b.rows.forEach((r, k) => { if (!used.has(k)) rows.push([...a.cols.map(() => null), ...r]); }); return new MTable([...a.cols, ...b.cols], rows, [...a.types, ...b.types]); });
  def('Table.ExpandTableColumn', (t, c, cols, names) => { t = needTable(t); const i = t.ci(needText(c)); const want = asCols(cols); const nn = cell(names) ? asCols(names) : want; const rows = []; let types = want.map(() => null);
    for (const r of t.rows) { const sub = cell(r[i]); if (sub === null || (isTable(sub) && !sub.rows.length)) { rows.push([...r.slice(0, i), ...want.map(() => null), ...r.slice(i + 1)]); if (isTable(sub)) types = want.map(w => sub.cols.includes(w) ? sub.types[sub.cols.indexOf(w)] : null); continue; } const s = needTable(sub); types = want.map(w => s.cols.includes(w) ? s.types[s.cols.indexOf(w)] : null);
      for (const sr of s.rows) rows.push([...r.slice(0, i), ...want.map(w => { const j = s.cols.indexOf(w); if (j < 0) throw err(`The column '${w}' of the table wasn't found.`); return sr[j]; }), ...r.slice(i + 1)]); }
    const cols2 = [...t.cols.slice(0, i), ...nn, ...t.cols.slice(i + 1)]; const dupe = nn.find(n => t.cols.includes(n) && n !== t.cols[i]); if (dupe) throw err(`The field '${dupe}' already exists in the record.`);
    return new MTable(cols2, rows, [...t.types.slice(0, i), ...types, ...t.types.slice(i + 1)]); });
  def('Table.ExpandRecordColumn', (t, c, fields, names) => { t = needTable(t); const i = t.ci(needText(c)); const want = asCols(fields); const nn = cell(names) ? asCols(names) : want; return new MTable([...t.cols.slice(0, i), ...nn, ...t.cols.slice(i + 1)], t.rows.map(r => { const rec = cell(r[i]); return [...r.slice(0, i), ...want.map(w => rec === null ? null : rec.get(w, true)), ...r.slice(i + 1)]; }), [...t.types.slice(0, i), ...nn.map(() => null), ...t.types.slice(i + 1)]); });
  def('Table.ExpandListColumn', (t, c) => { t = needTable(t); const i = t.ci(needText(c)); const rows = []; for (const r of t.rows) { const l = cell(r[i]); if (l === null || !isList(l) || !l.length) { rows.push([...r.slice(0, i), isList(l) ? null : l, ...r.slice(i + 1)]); continue; } for (const v of l) rows.push([...r.slice(0, i), v, ...r.slice(i + 1)]); } return new MTable(t.cols, rows, t.types.map((x, j) => j === i ? null : x)); });
  def('Table.AggregateTableColumn', (t, c, aggs) => { t = needTable(t); const i = t.ci(needText(c)); const al = needList(aggs).map(needList); return new MTable([...t.cols.slice(0, i), ...al.map(a => needText(a[2])), ...t.cols.slice(i + 1)], t.rows.map(r => { const sub = cell(r[i]); return [...r.slice(0, i), ...al.map(a => wrap(() => cell(needFn(a[1]).call([isTable(sub) ? sub.col(needText(a[0])) : []])))), ...r.slice(i + 1)]; })); });
  // text
  const T1 = (n, f) => def(n, (s, ...a) => { s = cell(s); if (s === null) return null; return f(needText(s), ...a.map(x => x === undefined ? undefined : x)); });
  T1('Text.Upper', s => s.toUpperCase()); T1('Text.Lower', s => s.toLowerCase()); T1('Text.Proper', s => s.toLowerCase().replace(/(^|[^A-Za-zÀ-￿'])([a-zà-￿])/g, (m, a, b) => a + b.toUpperCase()));
  T1('Text.Trim', (s, ch) => { ch = cell(ch); if (ch) { const set = isList(ch) ? ch.map(cell) : [ch]; let a = 0, b = s.length; while (a < b && set.includes(s[a])) a++; while (b > a && set.includes(s[b - 1])) b--; return s.slice(a, b); } return s.replace(/^\s+|\s+$/g, ''); });
  T1('Text.TrimStart', (s, ch) => { ch = cell(ch); if (ch) { const set = isList(ch) ? ch.map(cell) : [ch]; let a = 0; while (a < s.length && set.includes(s[a])) a++; return s.slice(a); } return s.replace(/^\s+/, ''); });
  T1('Text.TrimEnd', (s, ch) => { ch = cell(ch); if (ch) { const set = isList(ch) ? ch.map(cell) : [ch]; let b = s.length; while (b > 0 && set.includes(s[b - 1])) b--; return s.slice(0, b); } return s.replace(/\s+$/, ''); });
  T1('Text.Clean', s => s.replace(/[\u0000-\u001F\u007F]/g, '')); T1('Text.Length', s => s.length); T1('Text.Reverse', s => [...s].reverse().join(''));
  T1('Text.Start', (s, n) => s.slice(0, needNum(n))); T1('Text.End', (s, n) => { n = needNum(n); return n ? s.slice(-n) : ''; });
  T1('Text.Middle', (s, st, n) => { st = needNum(st); n = cell(n); return s.slice(st, n === null || n === undefined ? undefined : st + n); });
  T1('Text.Range', (s, st, n) => { st = needNum(st); n = cell(n); if (st > s.length || (n !== null && n !== undefined && st + n > s.length)) throw err('The \'offset\' argument is out of range.'); return s.slice(st, n === null || n === undefined ? undefined : st + n); });
  T1('Text.At', (s, i) => { i = needNum(i); if (i >= s.length) throw err('The \'index\' argument is out of range.'); return s[i]; });
  const cmpText = (s, sub, c) => isCI(c) ? [s.toLowerCase(), needText(sub).toLowerCase()] : [s, needText(sub)];
  T1('Text.Contains', (s, sub, c) => { const [a, b] = cmpText(s, sub, c); return a.includes(b); }); T1('Text.StartsWith', (s, sub, c) => { const [a, b] = cmpText(s, sub, c); return a.startsWith(b); }); T1('Text.EndsWith', (s, sub, c) => { const [a, b] = cmpText(s, sub, c); return a.endsWith(b); });
  T1('Text.Replace', (s, o, n) => { o = needText(o); return o ? s.split(o).join(needText(n)) : s; });
  T1('Text.Remove', (s, ch) => { ch = cell(ch); const set = isList(ch) ? ch.map(cell) : [needText(ch)]; return [...s].filter(c => !set.includes(c)).join(''); });
  T1('Text.Select', (s, ch) => { ch = cell(ch); const set = isList(ch) ? ch.map(cell) : [needText(ch)]; return [...s].filter(c => set.includes(c)).join(''); });
  T1('Text.Split', (s, d) => s.split(needText(d))); T1('Text.SplitAny', (s, d) => s.split(new RegExp('[' + needText(d).replace(/[\]\\^-]/g, '\\$&') + ']'))); T1('Text.ToList', s => [...s]);
  T1('Text.PadStart', (s, n, c) => s.padStart(needNum(n), cell(c) ?? ' ')); T1('Text.PadEnd', (s, n, c) => s.padEnd(needNum(n), cell(c) ?? ' ')); T1('Text.Repeat', (s, n) => s.repeat(needNum(n)));
  T1('Text.Insert', (s, o, x) => s.slice(0, needNum(o)) + needText(x) + s.slice(needNum(o)));
  T1('Text.PositionOf', (s, sub, occ) => { sub = needText(sub); occ = cell(occ) ?? 0; if (occ === 1) return s.lastIndexOf(sub); if (occ === 2) { const out = []; let k = s.indexOf(sub); while (k >= 0) { out.push(k); k = s.indexOf(sub, k + 1); } return out; } return s.indexOf(sub); });
  const delimAt = (s, d, idx, after) => { d = needText(d); idx = cell(idx); let k = 0, fromEnd = false; if (isNum(idx)) k = idx; else if (isList(idx)) { k = needNum(idx[0]); fromEnd = cell(idx[1]) === 1; } const pos = []; let p = s.indexOf(d); while (p >= 0) { pos.push(p); p = s.indexOf(d, p + d.length); } if (fromEnd) pos.reverse(); return pos[k]; };
  T1('Text.BeforeDelimiter', (s, d, i) => { const p = delimAt(s, d, i); return p === undefined ? s : s.slice(0, p); });
  T1('Text.AfterDelimiter', (s, d, i) => { const p = delimAt(s, d, i); return p === undefined ? '' : s.slice(p + needText(d).length); });
  T1('Text.BetweenDelimiters', (s, a, b) => { const p = s.indexOf(needText(a)); if (p < 0) return ''; const rest = s.slice(p + needText(a).length); const q = rest.indexOf(needText(b)); return q < 0 ? rest : rest.slice(0, q); });
  def('Text.Combine', (list, sep) => needList(list).map(cell).filter(v => v !== null).map(needText).join(cell(sep) ?? ''));
  def('Text.From', (v, c) => toText(v, C(c))); def('Text.Format', (f, args) => needText(f).replace(/#\{(\d+)\}|#\[([^\]]+)\]/g, (m, i, n) => { const a = cell(args); return toText(i !== undefined ? a[+i] : a.get(n)) ?? ''; }));
  def('Character.FromNumber', n => String.fromCharCode(needNum(n))); def('Character.ToNumber', s => needText(s).charCodeAt(0));
  // numbers
  const N1 = (n, f) => def(n, (x, ...a) => { x = cell(x); if (x === null) return null; return f(needNum(x), ...a.map(cell)); });
  N1('Number.Abs', Math.abs); N1('Number.Sign', Math.sign); N1('Number.Sqrt', Math.sqrt); N1('Number.Exp', Math.exp); N1('Number.Ln', Math.log); N1('Number.Log10', Math.log10); N1('Number.Power', (x, y) => x ** y);
  N1('Number.Mod', (x, y) => x - y * Math.trunc(x / y)); N1('Number.IntegerDivide', (x, y) => Math.trunc(x / y)); N1('Number.IsEven', x => x % 2 === 0); N1('Number.IsOdd', x => Math.abs(x % 2) === 1); def('Number.IsNaN', x => Number.isNaN(cell(x)));
  const rnd = (x, d, mode) => { d = d ?? 0; const f = 10 ** d; switch (mode ?? 4) { case 0: return Math.ceil(x * f - 1e-9) / f; case 1: return Math.floor(x * f + 1e-9) / f; case 2: return roundAway(x, d); case 3: return Math.trunc(x * f) / f; default: return roundEven(x, d); } };
  N1('Number.Round', (x, d, m) => rnd(x, d, m)); N1('Number.RoundUp', (x, d) => rnd(x, d, 0)); N1('Number.RoundDown', (x, d) => rnd(x, d, 1)); N1('Number.RoundAwayFromZero', (x, d) => rnd(x, d, 2)); N1('Number.RoundTowardZero', (x, d) => rnd(x, d, 3));
  def('Number.From', (v, c) => convert(v, 'number', C(c))); def('Number.FromText', (v, c) => { v = cell(v); if (v === null) return null; const n = parseNumText(needText(v), C(c)); if (n === undefined || n === null) throw dataErr(v, 'Number'); return n; });
  def('Number.ToText', (v, f, c) => { v = cell(v); if (v === null) return null; v = needNum(v); f = cell(f); if (!f) return numText(v, C(c)); const m = /^([NFPD])(\d*)$/i.exec(f); if (m) { const d = m[2] === '' ? 2 : +m[2]; const k = m[1].toUpperCase(); if (k === 'P') return (v * 100).toFixed(d) + ' %'; if (k === 'N') return v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d }); if (k === 'D') return String(Math.round(v)).padStart(d, '0'); return v.toFixed(d); } return numText(v, C(c)); });
  for (const [n, to] of [['Int64.From', 'Int64'], ['Int32.From', 'Int64'], ['Decimal.From', 'number'], ['Double.From', 'number'], ['Currency.From', 'Currency'], ['Percentage.From', 'number'], ['Logical.From', 'logical'], ['Date.From', 'date'], ['DateTime.From', 'datetime'], ['Time.From', 'time'], ['Duration.From', 'duration']]) def(n, (v, c) => convert(v, to, C(c)));
  def('Logical.FromText', v => convert(needText(v), 'logical')); def('Logical.ToText', v => cell(v) === null ? null : String(cell(v)));
  // dates
  const D1 = (n, f) => def(n, (d, ...a) => { d = cell(d); if (d === null) return null; return f(asDateish(d), ...a.map(cell)); });
  const sameKind = (d, n) => isDT(d) ? new MDateTime(n + (d.n - Math.floor(d.n)) * 0) : new MDate(n);
  def('#date', (y, m, d) => { y = needNum(y); m = needNum(m); d = needNum(d); if (m < 1 || m > 12 || d < 1 || d > 31 || ymd(ser(y, m, d)).d !== d) throw err('The Date operation failed because the resulting value falls outside the range of allowed values.'); return new MDate(ser(y, m, d)); });
  def('#datetime', (y, m, d, h, mi, s) => new MDateTime(ser(needNum(y), needNum(m), needNum(d)) + (needNum(h) * 3600 + needNum(mi) * 60 + needNum(s)) / 86400));
  def('#time', (h, mi, s) => new MTime((needNum(h) * 3600 + needNum(mi) * 60 + needNum(s)) / 86400));
  def('#duration', (d, h, mi, s) => new MDuration(needNum(d) + (needNum(h) * 3600 + needNum(mi) * 60 + needNum(s)) / 86400));
  D1('Date.Year', d => ymd(d.n).y); D1('Date.Month', d => ymd(d.n).m); D1('Date.Day', d => ymd(d.n).d); D1('Date.QuarterOfYear', d => Math.ceil(ymd(d.n).m / 3));
  D1('Date.MonthName', d => MONTHS[ymd(d.n).m - 1]); D1('Date.DayOfWeekName', d => DAYS[ymd(d.n).wd]); D1('Date.DayOfWeek', (d, f) => (ymd(d.n).wd - (f ?? 0) + 7) % 7);
  D1('Date.DayOfYear', d => Math.floor(d.n) - ser(ymd(d.n).y, 1, 1) + 1); D1('Date.DaysInMonth', d => { const { y, m } = ymd(d.n); return ymd(ser(y, m + 1, 1) - 1).d; }); D1('Date.IsLeapYear', d => { const y = ymd(d.n).y; return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; });
  D1('Date.WeekOfYear', (d, f) => { const y = ymd(d.n).y; const jan1 = ser(y, 1, 1); const off = (ymd(jan1).wd - (f ?? 0) + 7) % 7; return Math.floor((Math.floor(d.n) - jan1 + off) / 7) + 1; });
  D1('Date.WeekOfMonth', (d, f) => { const { y, m } = ymd(d.n); const first = ser(y, m, 1); const off = (ymd(first).wd - (f ?? 0) + 7) % 7; return Math.floor((Math.floor(d.n) - first + off) / 7) + 1; });
  D1('Date.StartOfMonth', d => { const { y, m } = ymd(d.n); return sameKind(d, ser(y, m, 1)); }); D1('Date.EndOfMonth', d => { const { y, m } = ymd(d.n); return isDT(d) ? new MDateTime(ser(y, m + 1, 1) - 1 / 86400) : new MDate(ser(y, m + 1, 1) - 1); });
  D1('Date.StartOfYear', d => sameKind(d, ser(ymd(d.n).y, 1, 1))); D1('Date.EndOfYear', d => isDT(d) ? new MDateTime(ser(ymd(d.n).y + 1, 1, 1) - 1 / 86400) : new MDate(ser(ymd(d.n).y, 12, 31)));
  D1('Date.StartOfQuarter', d => { const { y, m } = ymd(d.n); return sameKind(d, ser(y, Math.floor((m - 1) / 3) * 3 + 1, 1)); }); D1('Date.EndOfQuarter', d => { const { y, m } = ymd(d.n); return new MDate(ser(y, Math.floor((m - 1) / 3) * 3 + 4, 1) - 1); });
  D1('Date.StartOfWeek', (d, f) => sameKind(d, Math.floor(d.n) - (ymd(d.n).wd - (f ?? 0) + 7) % 7)); D1('Date.EndOfWeek', (d, f) => sameKind(d, Math.floor(d.n) - (ymd(d.n).wd - (f ?? 0) + 7) % 7 + 6));
  const addMonths = (d, k) => { const { y, m, d: dd } = ymd(d.n); const tgt = ser(y, m + k, 1); const t = ymd(tgt); const last = ymd(ser(t.y, t.m + 1, 1) - 1).d; const n = ser(t.y, t.m, Math.min(dd, last)) + (d.n - Math.floor(d.n)); return isDT(d) ? new MDateTime(n) : new MDate(n); };
  D1('Date.AddDays', (d, k) => isDT(d) ? new MDateTime(d.n + needNum(k)) : new MDate(d.n + needNum(k))); D1('Date.AddWeeks', (d, k) => isDT(d) ? new MDateTime(d.n + 7 * needNum(k)) : new MDate(d.n + 7 * needNum(k)));
  D1('Date.AddMonths', (d, k) => addMonths(d, needNum(k))); D1('Date.AddQuarters', (d, k) => addMonths(d, 3 * needNum(k))); D1('Date.AddYears', (d, k) => addMonths(d, 12 * needNum(k)));
  D1('Date.IsInPreviousNMonths', (d, k) => { const now = ymd(ctx.now); const x = ymd(d.n); const diff = (now.y - x.y) * 12 + now.m - x.m; return diff >= 1 && diff <= needNum(k); });
  D1('Date.ToRecord', d => { const { y, m, d: dd } = ymd(d.n); return MRecord.of({ Year: y, Month: m, Day: dd }); });
  const fmtOpts = (o, c) => { o = cell(o); if (isRec(o)) return { f: o.get('Format', true), c: o.get('Culture', true) || C(c) }; return { f: o, c: C(c) }; };
  def('Date.ToText', (d, o, c) => { d = cell(d); if (d === null) return null; const x = fmtOpts(o, c); return x.f ? fmtDate(asDateish(d).n, x.f, x.c) : toText(isDT(d) ? new MDate(d.n) : d, x.c); });
  def('DateTime.ToText', (d, o, c) => { d = cell(d); if (d === null) return null; const x = fmtOpts(o, c); return x.f ? fmtDate(asDateish(d).n, x.f, x.c) : toText(d, x.c); });
  const cultOf = o => { o = cell(o); return isText(o) ? o : C(o); };
  def('Date.FromText', (s, o) => { s = cell(s); if (s === null) return null; return convert(needText(s), 'date', cultOf(o)); }); def('DateTime.FromText', (s, o) => { s = cell(s); if (s === null) return null; return convert(needText(s), 'datetime', cultOf(o)); });
  def('Time.FromText', s => convert(needText(s), 'time')); def('DateTime.LocalNow', () => new MDateTime(ctx.now + 0.5)); def('DateTime.FixedLocalNow', () => new MDateTime(ctx.now + 0.5)); def('Date.LocalNow', () => new MDate(ctx.now));
  def('DateTimeZone.LocalNow', () => new MDateTime(ctx.now + 0.5)); def('DateTime.Date', d => { d = cell(d); return d === null ? null : new MDate(asDateish(d).n); }); def('DateTime.Time', d => { d = cell(d); return d === null ? null : new MTime(asDateish(d).n % 1); });
  for (const [n, k] of [['Time.Hour', 3600], ['Time.Minute', 60], ['Time.Second', 1]]) def(n, t => { t = cell(t); if (t === null) return null; const s = Math.round((t.n % 1) * 86400); return k === 3600 ? Math.floor(s / 3600) : k === 60 ? Math.floor(s / 60) % 60 : s % 60; });
  const DU = (n, f) => def(n, d => { d = cell(d); if (d === null) return null; if (!isDur(d)) throw cantConvert(d, 'Duration'); return f(d.n); });
  DU('Duration.Days', n => Math.trunc(n + (n < 0 ? -1e-9 : 1e-9))); DU('Duration.Hours', n => Math.floor(Math.round((Math.abs(n) % 1) * 86400) / 3600) * Math.sign(n || 1)); DU('Duration.Minutes', n => Math.floor(Math.round((Math.abs(n) % 1) * 86400) / 60) % 60); DU('Duration.Seconds', n => Math.round((Math.abs(n) % 1) * 86400) % 60);
  DU('Duration.TotalDays', n => n); DU('Duration.TotalHours', n => n * 24); DU('Duration.TotalMinutes', n => n * 1440); DU('Duration.TotalSeconds', n => n * 86400);
  // lists
  const nums = l => needList(l).map(cell).filter(v => v !== null);
  def('List.Sum', l => { const v = nums(l); if (!v.length) return null; if (v.every(isDur)) return new MDuration(v.reduce((a, b) => a + b.n, 0)); return v.reduce((a, b) => a + needNum(b), 0); });
  def('List.Product', l => { const v = nums(l); return v.length ? v.reduce((a, b) => a * needNum(b), 1) : null; }); def('List.Average', l => { const v = nums(l); return v.length ? v.reduce((a, b) => a + needNum(b), 0) / v.length : null; });
  def('List.Min', (l, d) => { const v = nums(l); return v.length ? v.reduce((a, b) => mCompare(b, a) < 0 ? b : a) : (cell(d) ?? null); }); def('List.Max', (l, d) => { const v = nums(l); return v.length ? v.reduce((a, b) => mCompare(b, a) > 0 ? b : a) : (cell(d) ?? null); });
  def('List.Median', l => { const v = nums(l).map(needNum).sort((a, b) => a - b); if (!v.length) return null; const h = v.length >> 1; return v.length % 2 ? v[h] : (v[h - 1] + v[h]) / 2; });
  def('List.StandardDeviation', l => { const v = nums(l).map(needNum); const m = v.reduce((a, b) => a + b, 0) / v.length; return Math.sqrt(v.reduce((a, b) => a + (b - m) ** 2, 0) / (v.length - 1)); });
  def('List.Count', l => needList(l).length); def('List.NonNullCount', l => nums(l).length); def('List.IsEmpty', l => needList(l).length === 0);
  def('List.Distinct', (l, c) => { const eq = eqf(c); const out = []; for (const v of needList(l).map(cell)) if (!out.some(x => eq(x, v))) out.push(v); return out; });
  def('List.Contains', (l, v, c) => { const eq = eqf(c); return needList(l).some(x => eq(x, v)); }); def('List.ContainsAny', (l, vs, c) => { const eq = eqf(c); return needList(vs).some(v => needList(l).some(x => eq(x, v))); }); def('List.ContainsAll', (l, vs, c) => { const eq = eqf(c); return needList(vs).every(v => needList(l).some(x => eq(x, v))); });
  def('List.Select', (l, f) => needList(l).filter(x => call(f, x) === true)); def('List.Transform', (l, f) => needList(l).map(x => call(f, x)));
  def('List.Accumulate', (l, seed, f) => needList(l).reduce((a, x) => call(f, a, x), cell(seed)));
  def('List.Sort', (l, o) => { o = cell(o); const d = o === 1 ? -1 : 1; const v = needList(l).map(cell).slice(); if (isFn(o)) return v.sort((a, b) => o.params && o.params.length === 2 ? o.call([a, b]) : mCompare(o.call([a]), o.call([b]))); return v.sort((a, b) => mCompare(a, b) * d); });
  def('List.First', (l, d) => { l = needList(l); return l.length ? cell(l[0]) : (cell(d) ?? null); }); def('List.Last', (l, d) => { l = needList(l); return l.length ? cell(l[l.length - 1]) : (cell(d) ?? null); });
  def('List.FirstN', (l, n) => { l = needList(l); n = cell(n); if (isFn(n)) { let k = 0; while (k < l.length && n.call([l[k]]) === true) k++; return l.slice(0, k); } return l.slice(0, needNum(n)); });
  def('List.LastN', (l, n) => needList(l).slice(-needNum(n))); def('List.Skip', (l, n) => { l = needList(l); n = cell(n); if (isFn(n)) { let k = 0; while (k < l.length && n.call([l[k]]) === true) k++; return l.slice(k); } return l.slice(n ?? 1); });
  def('List.Range', (l, o, n) => { o = needNum(o); n = cell(n); return needList(l).slice(o, n === null || n === undefined ? undefined : o + n); }); def('List.Reverse', l => needList(l).slice().reverse());
  def('List.Combine', ls => needList(ls).flatMap(needList)); def('List.RemoveNulls', l => needList(l).filter(x => cell(x) !== null));
  def('List.RemoveItems', (l, r) => needList(l).filter(x => !needList(r).some(y => mEquals(x, y)))); def('List.RemoveMatchingItems', (l, r, c) => { const eq = eqf(c); return needList(l).filter(x => !needList(r).some(y => eq(x, y))); });
  def('List.Difference', (a, b) => needList(a).filter(x => !needList(b).some(y => mEquals(x, y)))); def('List.Intersect', ls => { const [f, ...rest] = needList(ls).map(needList); return f.filter(x => rest.every(r => r.some(y => mEquals(x, y)))); });
  def('List.Union', ls => { const out = []; for (const l of needList(ls).map(needList)) for (const x of l) if (!out.some(y => mEquals(x, y))) out.push(x); return out; });
  def('List.PositionOf', (l, v, occ, c) => { const eq = eqf(c); l = needList(l); const hits = l.map((x, i) => eq(x, v) ? i : -1).filter(i => i >= 0); occ = cell(occ) ?? 0; return occ === 2 ? hits : occ === 1 ? (hits.length ? hits[hits.length - 1] : -1) : (hits.length ? hits[0] : -1); });
  def('List.Numbers', (s, n, inc) => Array.from({ length: needNum(n) }, (_, i) => needNum(s) + i * (cell(inc) ?? 1)));
  def('List.Dates', (s, n, st) => { s = asDateish(s); st = cell(st); return Array.from({ length: needNum(n) }, (_, i) => isDT(s) ? new MDateTime(s.n + i * st.n) : new MDate(s.n + i * st.n)); });
  def('List.Generate', (init, cond, next, sel) => { const out = []; let v = call(init); let k = 0; while (call(cond, v) === true) { out.push(sel && cell(sel) ? call(sel, v) : v); v = call(next, v); if (++k > 100000) throw err('List.Generate ran too long.'); } return out; });
  def('List.Zip', ls => { ls = needList(ls).map(needList); const n = Math.max(0, ...ls.map(l => l.length)); return Array.from({ length: n }, (_, i) => ls.map(l => l[i] ?? null)); });
  def('List.AnyTrue', l => needList(l).some(x => cell(x) === true)); def('List.AllTrue', l => needList(l).every(x => cell(x) === true));
  def('List.MatchesAny', (l, f) => needList(l).some(x => call(f, x) === true)); def('List.MatchesAll', (l, f) => needList(l).every(x => call(f, x) === true));
  def('List.Repeat', (l, n) => { const out = []; for (let k = 0; k < needNum(n); k++) out.push(...needList(l)); return out; }); def('List.Buffer', l => needList(l));
  def('List.Mode', l => { const v = nums(l); let best = null, bc = 0; for (const x of v) { const c = v.filter(y => mEquals(x, y)).length; if (c > bc) { bc = c; best = x; } } return best; });
  def('List.Single', l => { l = needList(l); if (l.length !== 1) throw err('There were too many elements in the enumeration to complete the operation.'); return cell(l[0]); });
  // records and values
  def('Record.Field', (r, f) => cell(r).get(needText(f))); def('Record.FieldOrDefault', (r, f, d) => { r = cell(r); const i = r.idx(needText(f)); return i < 0 ? (cell(d) ?? null) : cell(r.vals[i]); });
  def('Record.FieldNames', r => cell(r).names.slice()); def('Record.FieldValues', r => cell(r).vals.map(v => v)); def('Record.FieldCount', r => cell(r).names.length); def('Record.HasFields', (r, f) => asCols(f).every(x => cell(r).idx(x) >= 0));
  def('Record.ToTable', r => { r = cell(r); return new MTable(['Name', 'Value'], r.names.map((n, i) => [n, r.vals[i]])); }); def('Record.FromList', (vals, names) => new MRecord(asCols(names), needList(vals).slice()));
  def('Record.AddField', (r, n, v) => { r = cell(r); return new MRecord([...r.names, needText(n)], [...r.vals, cell(v)]); }); def('Record.SelectFields', (r, f) => { r = cell(r); const fs = asCols(f); return new MRecord(fs, fs.map(x => r.get(x))); });
  def('Record.RemoveFields', (r, f) => { r = cell(r); const fs = asCols(f); const keep = r.names.filter(n => !fs.includes(n)); return new MRecord(keep, keep.map(n => r.vals[r.idx(n)])); });
  def('Value.Is', (v, t) => isType(cell(v), cell(t))); def('Value.Type', v => new MType(({ Number: 'number', Text: 'text', Logical: 'logical', Date: 'date', DateTime: 'datetime', Null: 'null', List: 'list', Record: 'record', Table: 'table' })[typeName(cell(v))] || 'any'));
  def('Value.ReplaceType', (v, t) => { v = cell(v); t = cell(t); if (isTable(v) && t && t.cols) return new MTable(t.cols.map(x => x[0]), v.rows, t.cols.map(x => typeKey(x[1]))); return v; });
  def('Value.Equals', (a, b) => mEquals(a, b)); def('Value.Compare', (a, b) => mCompare(a, b)); def('Value.FromText', (s, c) => { s = cell(s); if (s === null || !isText(s)) return s; const n = parseNumText(s, C(c)); if (n !== undefined && n !== null) return n; const d = parseDateText(s, C(c)); if (d !== null) return d % 1 ? new MDateTime(d) : new MDate(d); return s; });
  def('Error.Record', (r, m, d) => MRecord.of({ Reason: cell(r), Message: cell(m) ?? null, Detail: cell(d) ?? null }));
  def('Expression.Evaluate', (s, env) => evaluate(parseM(needText(s)), new Env(new Map(), rootEnv(ctx)))); def('Diagnostics.Trace', (a, b, v) => cell(v));
  return L;
}
const cell0 = v => v instanceof ErrCell ? v : cell(v);
// a plain {cols, rows, types} (from a lab or a workbook) as an M table; header:true for a sheet whose first row is data
export function toMTable(t, sheet) { if (t instanceof MTable) return t; if (sheet) { const n = Math.max(0, ...t.rows.map(r => r.length), (t.cols || []).length); const names = Array.from({ length: n }, (_, i) => `Column${i + 1}`); const rows = (t.cols ? [t.cols] : []).concat(t.rows); return new MTable(names, rows.map(r => names.map((_, i) => fromPlain(r[i]))), names.map(() => null)); }
  return new MTable(t.cols.slice(), t.rows.map(r => t.cols.map((_, i) => fromPlain(r[i]))), (t.types || t.cols.map(() => null)).slice()); }
// lab data uses {d: 'yyyy-mm-dd'} for dates and {dt: 'yyyy-mm-ddThh:mm'} for date-times
export function fromPlain(v) { if (v === undefined) return null; if (v && typeof v === 'object' && !Array.isArray(v)) { if (v.d) { const [y, m, d] = v.d.split('-').map(Number); return new MDate(ser(y, m, d)); } if (v.dt) { const [a, b] = v.dt.split('T'); const [y, m, d] = a.split('-').map(Number); return new MDateTime(ser(y, m, d) + (parseTimeText(b) || 0)); } } return v; }

// ---------------------------------------------------------------- engine: a set of queries over some sources
const LIBS = new WeakMap();
function rootEnv(ctx) { if (!ctx._root) { const lib = buildLib(ctx); const e = new Env(lib, null); e.resolve = name => ctx.query ? ctx.query(name) : undefined; ctx._root = e; } return ctx._root; }
export class MEngine {
  // src: { tables: {Name: {cols, rows, types?}}, files: {path: text | {workbook}}, folders: {path: [{name, text, modified?, hidden?}]}, culture, now (a day serial) }
  constructor(src = {}) { this.src = src; this.queries = {}; this.cache = new Map(); const self = this;
    this.ctx = { culture: src.culture || 'en-IN', now: src.now || ser(2026, 6, 30), tables: src.tables || {}, file: p => self.file(p), folder: p => self.folder(p), query: name => self.has(name) ? self.evaluate(name) : undefined }; }
  has(name) { return Object.prototype.hasOwnProperty.call(this.queries, name); }
  setQuery(name, text) { this.queries[name] = text; this.cache.clear(); } deleteQuery(name) { delete this.queries[name]; this.cache.clear(); }
  renameQuery(a, b) { const out = {}; for (const [k, v] of Object.entries(this.queries)) out[k === a ? b : k] = v; this.queries = out; this.cache.clear(); }
  setSource(src) { Object.assign(this.src, src); this.ctx.tables = this.src.tables || {}; this.ctx._root = null; this.cache.clear(); }
  file(p) { const files = this.src.files || {}; const base = s => String(s).split(/[\\/]/).pop().toLowerCase(); const k = Object.keys(files).find(f => f.toLowerCase() === String(p).toLowerCase()) || Object.keys(files).find(f => base(f) === base(p)); if (!k) return null; const f = files[k]; return new MBinary(base(k), typeof f === 'string' ? f : (f.text || ''), typeof f === 'string' ? null : f); }
  folder(p) { const fs = this.src.folders || {}; const norm = s => String(s).replace(/[\\/]+$/, '').toLowerCase(); const last = s => norm(s).split(/[\\/]/).pop(); let k = Object.keys(fs).find(f => norm(f) === norm(p)) || Object.keys(fs).find(f => last(f) === last(p)); if (!k && this.src.looseFolders && Object.keys(fs).length === 1) k = Object.keys(fs)[0]; if (!k) return null;  // looseFolders: an uploaded workbook's folder may have any name
    return fs[k].map(f => ({ name: f.name, bin: new MBinary(f.name, typeof f.text === 'string' ? f.text : '', f.workbook ? { workbook: f.workbook } : null), modified: f.modified || this.ctx.now, hidden: !!f.hidden, folder: String(p).replace(/[\\/]?$/, '\\') })); }
  parse(name) { return parseM(this.queries[name]); }
  evaluate(name) { if (this.cache.has(name)) { const c = this.cache.get(name); if (c.busy) throw err('A cyclic reference was encountered during evaluation.'); if (c.e) throw c.e; return c.v; } const c = { busy: true }; this.cache.set(name, c);
    try { c.v = force(evaluate(this.parse(name), new Env(new Map(), rootEnv(this.ctx)))); } catch (e) { c.e = e instanceof MError ? e : err(String(e && e.message || e)); } c.busy = false; if (c.e) throw c.e; return c.v; }
  eval(text, vars = {}) { const m = new Map(Object.entries(vars)); return force(evaluate(parseM(text), new Env(m, rootEnv(this.ctx)))); }
  // the Applied Steps of a query: [{name, text}] from its let … in, with the step the query returns
  steps(name) { return splitSteps(this.queries[name]); }
  evalStep(name, i) { const s = splitSteps(this.queries[name]); if (!s.steps.length) return this.evaluate(name); const text = joinSteps(s.steps.slice(0, i + 1), s.steps[i].name); return force(evaluate(parseM(text), new Env(new Map(), rootEnv(this.ctx)))); }
}
// let … in as editable steps: keep each step's own text, exactly as typed
export function splitSteps(src) { src = src || ''; let ast; try { ast = parseM(src); } catch (e) { return { steps: [], error: e, raw: src }; }
  if (ast.k !== 'let') return { steps: [{ name: 'Source', text: src.trim() }], result: 'Source', single: true };
  return { steps: ast.binds.map(b => ({ name: b.name, text: src.slice(b.s, b.end).trim() })), result: ast.body.k === 'id' ? ast.body.name : null }; }
export const mId = n => /^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)*$/.test(n) && !KW.has(n) ? n : `#"${String(n).replace(/"/g, '""')}"`;
export const mStr = s => '"' + String(s).replace(/"/g, '""').replace(/\n/g, '#(lf)').replace(/\r/g, '#(cr)').replace(/\t/g, '#(tab)') + '"';
export const mField = n => /^[A-Za-z_][A-Za-z0-9_.]*( [A-Za-z0-9_.]+)*$/.test(n) && !KW.has(n) ? `[${n}]` : `[#"${String(n).replace(/"/g, '""')}"]`;
export function joinSteps(steps, result) { return 'let\n' + steps.map((s, i) => `    ${mId(s.name)} = ${s.text}${i < steps.length - 1 ? ',' : ''}`).join('\n') + `\nin\n    ${mId(result || (steps.length ? steps[steps.length - 1].name : 'Source'))}`; }
// how the preview shows a value
export function showCell(v, culture) { if (v instanceof ErrCell) return { t: 'Error', cls: 'err' }; v = v && v.th ? force(v) : v; if (v === null || v === undefined) return { t: 'null', cls: 'null' };
  if (isTable(v)) return { t: 'Table', cls: 'nest' }; if (isRec(v)) return { t: 'Record', cls: 'nest' }; if (isList(v)) return { t: 'List', cls: 'nest' }; if (isFn(v)) return { t: 'Function', cls: 'nest' }; if (v instanceof MBinary) return { t: 'Binary', cls: 'nest' }; if (v instanceof MType) return { t: 'Type', cls: 'nest' };
  if (isLog(v)) return { t: v ? 'TRUE' : 'FALSE', cls: 'log' }; if (isNum(v)) return { t: numText(v, culture), cls: 'num' }; return { t: toText(v, culture), cls: isText(v) ? 'txt' : 'num' }; }
// a value as plain data for the grid: numbers, text, TRUE/FALSE, day serials (with a format), or null
export function toPlain(v) { if (v instanceof ErrCell) return { err: v.e.text }; if (v === null || v === undefined) return null; if (isNum(v) || isText(v) || isLog(v)) return v; if (isDate(v)) return { serial: v.n, fmt: 'dd-mm-yyyy' }; if (isDT(v)) return { serial: v.n, fmt: 'dd-mm-yyyy hh:mm' }; if (isTime(v)) return { serial: v.n, fmt: 'hh:mm' }; if (isDur(v)) return { serial: v.n, fmt: '[h]:mm:ss' }; return `[${typeName(v)}]`; }
export { force, cell as cellValue, Env };
