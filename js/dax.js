// DAX for the Data Model lab: measures and calculated columns, evaluated in a filter context that flows from lookup
// tables to fact tables along active relationships, with row context for calculated columns and iterators, context
// transition when a measure is used in a row context, CALCULATE's filter arguments, ALL/ALLSELECTED/REMOVEFILTERS,
// RELATED, RANKX, DIVIDE and time intelligence on the marked date table.
import { ser, ymd } from './m.js';
export const BL = Symbol('BLANK');
export class DaxError extends Error {}
const err = m => { throw new DaxError(m); };
const isBlank = v => v === BL || v === null || v === undefined;

// ---------------------------------------------------------------- parser
function lex(s) { const out = []; let i = 0; const n = s.length;
  while (i < n) { const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (s.startsWith('//', i) || s.startsWith('--', i)) { while (i < n && s[i] !== '\n') i++; continue; }
    if (s.startsWith('/*', i)) { const j = s.indexOf('*/', i + 2); i = j < 0 ? n : j + 2; continue; }
    if (c === '"') { let j = i + 1, v = ''; for (;;) { if (j >= n) err('A string is missing its closing quote.'); if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; } v += s[j++]; } out.push({ t: 'str', v, p: i }); i = j + 1; continue; }
    if (c === "'") { const j = s.indexOf("'", i + 1); if (j < 0) err('A table name is missing its closing quote.'); out.push({ t: 'tbl', v: s.slice(i + 1, j), p: i }); i = j + 1; continue; }
    if (c === '[') { const j = s.indexOf(']', i + 1); if (j < 0) err('A column or measure name is missing its closing ].'); out.push({ t: 'br', v: s.slice(i + 1, j).trim(), p: i }); i = j + 1; continue; }
    if (/[0-9]/.test(c) || (c === '.' && /[0-9]/.test(s[i + 1] || ''))) { const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i)); out.push({ t: 'num', v: +m[0], p: i }); i += m[0].length; continue; }
    if (/[A-Za-z_]/.test(c)) { const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i)); out.push({ t: 'id', v: m[0], p: i }); i += m[0].length; continue; }
    const two = s.slice(i, i + 2); if (['&&', '||', '<>', '<=', '>=', '=='].includes(two)) { out.push({ t: 'op', v: two, p: i }); i += 2; continue; }
    if ('+-*/^&=<>(),{}'.includes(c)) { out.push({ t: 'op', v: c, p: i }); i++; continue; }
    err(`The syntax for '${c}' is incorrect.`); }
  out.push({ t: 'eof', p: n }); return out; }
export function parseDax(src) { const T = lex(String(src).replace(/^\s*=/, '')); let k = 0; const peek = () => T[k], take = () => T[k++];
  const is = (t, v) => T[k].t === t && (v === undefined || T[k].v === v); const want = (t, v) => { if (!is(t, v)) err(`The syntax for '${T[k].t === 'eof' ? 'the end of the formula' : T[k].v}' is incorrect.${v ? ` Expected ${v}.` : ''}`); return take(); };
  const bin = (ops, next) => () => { let a = next(); while (T[k].t === 'op' && ops.includes(T[k].v)) { const op = take().v; a = { k: 'bin', op: op === '==' ? '=' : op, a, b: next() }; } return a; };
  const prim = () => { const t = peek();
    if (is('op', '(')) { take(); const e = expr(); want('op', ')'); return e; }
    if (is('op', '-')) { take(); return { k: 'neg', a: unary() }; } if (is('op', '+')) { take(); return unary(); }
    if (t.t === 'num') { take(); return { k: 'lit', v: t.v }; } if (t.t === 'str') { take(); return { k: 'lit', v: t.v }; }
    if (t.t === 'br') { take(); return { k: 'measure', name: t.v }; }
    if (t.t === 'tbl') { take(); if (is('br')) return { k: 'col', t: t.v, c: take().v }; return { k: 'table', t: t.v }; }
    if (t.t === 'id') { take(); if (is('op', '(')) { take(); const args = []; if (!is('op', ')')) do { args.push(expr()); } while (is('op', ',') && take()); want('op', ')'); return { k: 'call', f: t.v.toUpperCase(), args }; }
      if (is('br')) return { k: 'col', t: t.v, c: take().v }; const u = t.v.toUpperCase(); if (u === 'TRUE') return { k: 'lit', v: true }; if (u === 'FALSE') return { k: 'lit', v: false }; return { k: 'table', t: t.v }; }
    err(`The syntax for '${t.t === 'eof' ? 'the end of the formula' : t.v}' is incorrect.`); };
  const unary = () => prim(); const pow = bin(['^'], unary), mul = bin(['*', '/'], pow), add = bin(['+', '-'], mul), cat = bin(['&'], add), cmp = bin(['=', '<>', '<', '>', '<=', '>='], cat), and = bin(['&&'], cmp), or = bin(['||'], and); function expr() { return or(); }
  const e = expr(); if (!is('eof')) err(`The syntax for '${peek().v}' is incorrect.`); return e; }

// ---------------------------------------------------------------- filter contexts
// F: { cols: Map('T[C]' → Set of values), tabs: Map(T → Set of row indexes), outer: the report's own filters (for ALLSELECTED) }
export const emptyF = () => ({ cols: new Map(), tabs: new Map(), outer: null });
const cloneF = F => ({ cols: new Map(F.cols), tabs: new Map(F.tabs), outer: F.outer });
const keyOf = (t, c) => `${t}[${c}]`;
const splitKey = k => { const m = /^(.*)\[([^\]]*)\]$/.exec(k); return { t: m[1], c: m[2] }; };

export class Dax {
  // M: a Model (tables, rels, dateTable). measures: [{name, table, expr, fmt}]
  constructor(M, measures = []) { this.M = M; this.measures = new Map(); for (const m of measures) this.measures.set(m.name.toLowerCase(), { ...m }); this.ast = new Map(); this.memo = new Map(); this.stack = []; }
  parsed(src) { if (!this.ast.has(src)) this.ast.set(src, parseDax(src)); return this.ast.get(src); }
  T(t) { const x = this.M.tables[t] || Object.values(this.M.tables).find(q => q.name.toLowerCase() === String(t).toLowerCase()); if (!x) err(`Failed to resolve name '${t}'. It is not a valid table, variable, or function name.`); return x; }
  colIndex(t, c) { const T = this.T(t); let j = T.cols.indexOf(c); if (j < 0) j = T.cols.findIndex(x => x.toLowerCase() === String(c).toLowerCase()); if (j < 0) err(`Column '${c}' in table '${T.name}' cannot be found or may not be used in this expression.`); return j; }
  // ---- which rows of a table the filter context lets through
  filtered(t, F, seen = new Set()) { if (F.tabs.has(t)) return true; for (const k of F.cols.keys()) if (splitKey(k).t === t) return true; if (seen.has(t)) return false; seen.add(t); return this.M.rels.some(r => r.active && splitKey(r.from).t === t && this.filtered(splitKey(r.to).t, F, seen)); }
  blankOK(t, F) { if (F.tabs.has(t)) return false; for (const [k, s] of F.cols) if (splitKey(k).t === t && !s.has(BL)) return false; return this.M.rels.every(r => !r.active || splitKey(r.from).t !== t || !this.filtered(splitKey(r.to).t, F) || this.blankOK(splitKey(r.to).t, F)); }
  visible(t, F) { const T = this.T(t); t = T.name; F._vis = F._vis || new Map(); if (F._vis.has(t)) return F._vis.get(t); let rows = T.rows.map((_, i) => i);
    if (F.tabs.has(t)) { const s = F.tabs.get(t); rows = rows.filter(i => s.has(i)); }
    for (const [k, s] of F.cols) { const p = splitKey(k); if (p.t !== t) continue; const j = this.colIndex(t, p.c); rows = rows.filter(i => { const v = T.rows[i][j]; return s.has(isBlank(v) ? BL : v); }); }
    for (const r of this.M.rels) { if (!r.active) continue; const f = splitKey(r.from), o = splitKey(r.to); if (f.t !== t || !this.filtered(o.t, F)) continue;
      const D = this.T(o.t); const jo = this.colIndex(o.t, o.c); const keys = new Set(this.visible(o.t, F).map(i => D.rows[i][jo])); const all = this.M.keyIndex(r.to); const bok = this.blankOK(o.t, F); const jf = this.colIndex(t, f.c);
      rows = rows.filter(i => { const v = T.rows[i][jf]; return keys.has(v) || (bok && !all.has(v)); }); }
    F._vis.set(t, rows); return rows; }
  // ---- evaluation. R: row context [{t, i} | {col, v}]
  measure(name, F, R = []) { const m = this.measures.get(String(name).toLowerCase()); if (!m) err(`The value for '${name}' cannot be determined. Either the column doesn't exist, or there is no current row for this column.`);
    if (this.stack.includes(m.name)) err(`A circular dependency was detected: ${[...this.stack, m.name].join(', ')}.`);
    const G = R.length ? this.transition(F, R) : F; this.stack.push(m.name); try { return this.ev(this.parsed(m.expr), G, []); } finally { this.stack.pop(); } }
  transition(F, R) { const G = cloneF(F); for (const r of R) { if (r.col) G.cols.set(r.col, new Set([r.v])); else { for (const k of [...G.cols.keys()]) if (splitKey(k).t === r.t) G.cols.delete(k); G.tabs.set(r.t, new Set([r.i])); } } return G; }  // a whole row filters every column of its table: it replaces the filters on them
  value(src, F, R = []) { return this.ev(typeof src === 'string' ? this.parsed(src) : src, F || emptyF(), R); }
  ev(n, F, R) { switch (n.k) {
    case 'lit': return n.v;
    case 'neg': { const a = this.ev(n.a, F, R); return isBlank(a) ? BL : -num(a); }
    case 'measure': { if (this.measures.has(n.name.toLowerCase())) return this.measure(n.name, F, R); const rc = R.find(r => r.t && this.T(r.t).cols.some(c => c.toLowerCase() === n.name.toLowerCase())); if (rc) return this.cell(rc.t, rc.i, n.name); return this.measure(n.name, F, R); }
    case 'col': { const T = this.T(n.t); const r = [...R].reverse().find(x => (x.t && x.t === T.name) || (x.col && x.col.toLowerCase() === keyOf(T.name, n.c).toLowerCase())); if (!r) { this.colIndex(n.t, n.c); err(`A single value for column '${n.c}' in table '${T.name}' cannot be determined. This can happen when a measure formula refers to a column that contains many values without specifying an aggregation such as min, max, count, or sum to get a single result.`); } return r.col ? r.v : this.cell(T.name, r.i, n.c); }
    case 'table': return { kind: 'rows', t: this.T(n.t).name, rows: this.visible(n.t, F) };
    case 'bin': return this.bin(n, F, R);
    case 'call': { const f = FN[n.f]; if (!f) err(`Failed to resolve name '${n.f}'. It is not a valid table, variable, or function name.`); return f(this, n.args, F, R); } }
    err('Unknown expression.'); }
  cell(t, i, c) { const T = this.T(t); const v = T.rows[i][this.colIndex(t, c)]; return isBlank(v) ? BL : v; }
  bin(n, F, R) { const a = this.ev(n.a, F, R), b = this.ev(n.b, F, R); const op = n.op;
    if (op === '&&') return truthy(a) && truthy(b); if (op === '||') return truthy(a) || truthy(b);
    if (op === '&') return str(a) + str(b);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) { const x = isBlank(a) ? (typeof b === 'string' ? '' : 0) : a, y = isBlank(b) ? (typeof a === 'string' ? '' : 0) : b; const c = typeof x === 'string' || typeof y === 'string' ? String(x).toLowerCase().localeCompare(String(y).toLowerCase()) : num(x) - num(y);
      return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[op]; }
    if (op === '+' || op === '-') { if (isBlank(a) && isBlank(b)) return BL; const x = isBlank(a) ? 0 : num(a), y = isBlank(b) ? 0 : num(b); return op === '+' ? x + y : x - y; }
    if (isBlank(a) || isBlank(b)) return op === '/' && !isBlank(a) ? (num(a) === 0 ? NaN : (num(a) > 0 ? Infinity : -Infinity)) : BL;
    if (op === '*') return num(a) * num(b); if (op === '/') return num(b) === 0 ? (num(a) === 0 ? NaN : num(a) > 0 ? Infinity : -Infinity) : num(a) / num(b); if (op === '^') return Math.pow(num(a), num(b)); }
  // ---- a table expression → its rows (base table rows, or the values of one column)
  table(n, F, R) { const v = this.ev(n, F, R); if (!v || !v.kind) err('The expression refers to multiple columns. Multiple columns cannot be converted to a scalar value.'); return v; }
  iterate(tv) { return tv.kind === 'rows' ? tv.rows.map(i => ({ t: tv.t, i })) : tv.vals.map(v => ({ col: tv.col, v })); }
  colValues(t, c, rows) { const T = this.T(t); const j = this.colIndex(t, c); const s = new Set(); for (const i of rows) { const v = T.rows[i][j]; s.add(isBlank(v) ? BL : v); } return [...s]; }
  // CALCULATE's filter arguments, evaluated in the outer context, then applied to a copy of it
  applyFilters(F, args, R) { const G = cloneF(F); const adds = [];
    for (const a of args) {
      if (a.k === 'call' && ['ALL', 'REMOVEFILTERS', 'ALLSELECTED', 'ALLEXCEPT'].includes(a.f)) { this.remove(G, a, F); continue; }
      if (a.k === 'bin' && ['=', '<>', '<', '>', '<=', '>=', '&&', '||'].includes(a.op) || (a.k === 'call' && ['NOT', 'AND', 'OR'].includes(a.f))) { const cols = refs(a); if (!cols.length) err('A boolean filter in CALCULATE must refer to a column.');
        const ts = [...new Set(cols.map(c => this.T(c.t).name))]; if (ts.length > 1) err('A filter expression in CALCULATE can only refer to columns from one table.'); const t = ts[0];
        if (new Set(cols.map(c => c.c.toLowerCase())).size === 1) { const c = this.T(t).cols[this.colIndex(t, cols[0].c)]; const key = keyOf(t, c); const vals = this.colValues(t, c, this.T(t).rows.map((_, i) => i)).filter(v => truthy(this.ev(a, F, [...R, { col: key, v }]))); adds.push(['col', key, new Set(vals)]); }
        else { const rows = this.T(t).rows.map((_, i) => i).filter(i => truthy(this.ev(a, F, [...R, { t, i }]))); for (const c of cols) G.cols.delete(keyOf(t, this.T(t).cols[this.colIndex(t, c.c)])); adds.push(['tab', t, new Set(rows)]); }
        continue; }
      const tv = this.table(a, F, R);
      if (tv.kind === 'vals') { const p = splitKey(tv.col); const d = this.M.dateTable; if (d && d.table === p.t && d.col === p.c) { for (const k of [...G.cols.keys()]) if (splitKey(k).t === p.t) G.cols.delete(k); G.tabs.delete(p.t); } adds.push(['col', tv.col, new Set(tv.vals)]); }
      else { for (const k of [...G.cols.keys()]) if (splitKey(k).t === tv.t) G.cols.delete(k); adds.push(['tab', tv.t, new Set(tv.rows)]); } }
    for (const [kind, k, s] of adds) { if (kind === 'col') G.cols.set(k, s); else G.tabs.set(k, s); } delete G._vis; return G; }
  remove(G, a, F) { const tgt = a.args.map(x => x.k === 'col' ? { t: this.T(x.t).name, c: this.T(x.t).cols[this.colIndex(x.t, x.c)] } : x.k === 'table' ? { t: this.T(x.t).name } : err(`${a.f} needs a table or a column.`));
    if (a.f === 'ALLEXCEPT') { const t = tgt[0].t; const keep = new Set(tgt.slice(1).map(x => keyOf(x.t, x.c))); for (const k of [...G.cols.keys()]) if (splitKey(k).t === t && !keep.has(k)) G.cols.delete(k); G.tabs.delete(t); return; }
    for (const x of tgt) { if (x.c) G.cols.delete(keyOf(x.t, x.c)); else { for (const k of [...G.cols.keys()]) if (splitKey(k).t === x.t) G.cols.delete(k); G.tabs.delete(x.t); }
      if (a.f === 'ALLSELECTED' && F.outer) for (const [k, s] of F.outer) { const p = splitKey(k); if (p.t === x.t && (!x.c || p.c === x.c)) G.cols.set(k, s); } } }
  // the dates of the marked date table that the context lets through
  dates(n, F, R) { const tv = n.k === 'col' ? { t: this.T(n.t).name, c: this.T(n.t).cols[this.colIndex(n.t, n.c)] } : err('Pass the date column, such as Calendar[Date].'); return { key: keyOf(tv.t, tv.c), vals: this.colValues(tv.t, tv.c, this.visible(tv.t, F)).filter(v => !isBlank(v)).sort((a, b) => a - b), all: this.colValues(tv.t, tv.c, this.T(tv.t).rows.map((_, i) => i)) }; }
}
// ---------------------------------------------------------------- helpers and functions
function num(v) { if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; if (isBlank(v)) return 0; const n = Number(v); if (isNaN(n)) err(`Cannot convert value '${v}' of type Text to type Number.`); return n; }
function str(v) { return isBlank(v) ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v); }
function truthy(v) { return !isBlank(v) && v !== false && v !== 0 && v !== ''; }
function refs(n, out = []) { if (!n || typeof n !== 'object') return out; if (n.k === 'col') out.push({ t: n.t, c: n.c }); for (const k of ['a', 'b']) if (n[k]) refs(n[k], out); for (const x of n.args || []) refs(x, out); return out; }
const need = (args, a, b, name) => { if (args.length < a || args.length > b) err(`Too ${args.length < a ? 'few' : 'many'} arguments were passed to the ${name} function.`); };
const colArg = (D, n, f) => { if (n.k !== 'col') err(`The ${f} function only accepts a column reference as an argument.`); const T = D.T(n.t); return { t: T.name, c: T.cols[D.colIndex(n.t, n.c)] }; };
const nums = xs => xs.filter(x => !isBlank(x) && typeof x !== 'string' && typeof x !== 'boolean').map(Number);
const agg = (kind) => (D, a, F) => { need(a, 1, 1, kind); const { t, c } = colArg(D, a[0], kind); const T = D.T(t), j = D.colIndex(t, c); const xs = D.visible(t, F).map(i => T.rows[i][j]); return reduce(kind, xs); };
function reduce(kind, xs) { const n = nums(xs); if (kind === 'COUNT') return xs.filter(x => !isBlank(x) && typeof x !== 'boolean').length || BL; if (kind === 'COUNTA') return xs.filter(x => !isBlank(x)).length || BL; if (kind === 'DISTINCTCOUNT') { const s = new Set(xs.map(x => isBlank(x) ? BL : x)); return s.size || BL; }
  if (!n.length) return BL; if (kind === 'SUM') return n.reduce((p, q) => p + q, 0); if (kind === 'AVERAGE') return n.reduce((p, q) => p + q, 0) / n.length; if (kind === 'MIN') return Math.min(...n); if (kind === 'MAX') return Math.max(...n); }
const iter = (kind) => (D, a, F, R) => { need(a, 2, 2, kind + 'X'); const tv = D.table(a[0], F, R); const xs = D.iterate(tv).map(r => D.ev(a[1], F, [...R, r])); return reduce(kind, xs); };
const shift = (n, months) => { const d = ymd(n); const t = d.m - 1 + months; const y = d.y + Math.floor(t / 12), m = ((t % 12) + 12) % 12 + 1; const last = ymd(ser(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1).d; return ser(y, m, Math.min(d.d, last)); };
// dates moved by whole months; a month's last day also brings the rest of the target month, so April moves to all of March (as DAX does)
const shiftSet = (vals, k, all) => { const out = new Set(); for (const v of vals) { const x = shift(v, k); out.add(x); if (ymd(v + 1).d === 1) { const y = ymd(x); const end = ser(y.m === 12 ? y.y + 1 : y.y, y.m === 12 ? 1 : y.m + 1, 1) - 1; for (let n = x + 1; n <= end; n++) out.add(n); } } return [...out].filter(v => all.has(v)).sort((p, q) => p - q); };
const fyStart = (n, yearEnd) => { const [em, ed] = String(yearEnd || '12/31').split(/[\/-]/).map(Number); const d = ymd(n); let y = d.y; const endThis = ser(y, em, ed); if (n > endThis) return endThis + 1; return ser(y - 1, em, ed) + 1; };
// TOTALYTD/QTD/MTD(expression, dates, [year end]): the expression with the date table's filters replaced by the period to date
const totalTo = name => (D, a, F, R) => { need(a, 2, 4, name.replace('DATES', 'TOTAL')); const ye = a.slice(2).find(x => x.k === 'lit' && typeof x.v === 'string'); const dt = FN[name](D, ye ? [a[1], ye] : [a[1]], F, R); const G = cloneF(R.length ? D.transition(F, R) : F); const p = splitKey(dt.col); for (const k of [...G.cols.keys()]) if (splitKey(k).t === p.t) G.cols.delete(k); G.tabs.delete(p.t); G.cols.set(dt.col, new Set(dt.vals)); return D.ev(a[0], G, []); };
export const FN = {
  SUM: agg('SUM'), AVERAGE: agg('AVERAGE'), MIN: agg('MIN'), MAX: agg('MAX'), COUNT: agg('COUNT'), COUNTA: agg('COUNTA'), DISTINCTCOUNT: agg('DISTINCTCOUNT'),
  COUNTROWS: (D, a, F, R) => { need(a, 1, 1, 'COUNTROWS'); const tv = D.table(a[0], F, R); const n = tv.kind === 'rows' ? tv.rows.length : tv.vals.length; return n || BL; },
  SUMX: iter('SUM'), AVERAGEX: iter('AVERAGE'), MINX: iter('MIN'), MAXX: iter('MAX'), COUNTX: iter('COUNT'),
  CALCULATE: (D, a, F, R) => { if (!a.length) err('Too few arguments were passed to the CALCULATE function.'); const G = D.applyFilters(R.length ? D.transition(F, R) : F, a.slice(1), []); return D.ev(a[0], G, []); },
  CALCULATETABLE: (D, a, F, R) => { const G = D.applyFilters(R.length ? D.transition(F, R) : F, a.slice(1), []); return D.table(a[0], G, []); },
  FILTER: (D, a, F, R) => { need(a, 2, 2, 'FILTER'); const tv = D.table(a[0], F, R); if (tv.kind === 'rows') return { kind: 'rows', t: tv.t, rows: tv.rows.filter(i => truthy(D.ev(a[1], F, [...R, { t: tv.t, i }]))) }; return { kind: 'vals', col: tv.col, vals: tv.vals.filter(v => truthy(D.ev(a[1], F, [...R, { col: tv.col, v }]))) }; },
  ALL: (D, a, F) => { need(a, 1, 1, 'ALL'); const x = a[0]; if (x.k === 'col') { const { t, c } = colArg(D, x, 'ALL'); return { kind: 'vals', col: keyOf(t, c), vals: D.colValues(t, c, D.T(t).rows.map((_, i) => i)) }; } const T = D.T(x.t); return { kind: 'rows', t: T.name, rows: T.rows.map((_, i) => i) }; },
  ALLSELECTED: (D, a, F) => { const x = a[0]; const G = emptyF(); if (F.outer) for (const [k, s] of F.outer) G.cols.set(k, s); if (x.k === 'col') { const { t, c } = colArg(D, x, 'ALLSELECTED'); return { kind: 'vals', col: keyOf(t, c), vals: D.colValues(t, c, D.visible(t, G)) }; } const T = D.T(x.t); return { kind: 'rows', t: T.name, rows: D.visible(T.name, G) }; },
  VALUES: (D, a, F) => { need(a, 1, 1, 'VALUES'); const x = a[0]; if (x.k === 'col') { const { t, c } = colArg(D, x, 'VALUES'); return { kind: 'vals', col: keyOf(t, c), vals: D.colValues(t, c, D.visible(t, F)) }; } const T = D.T(x.t); return { kind: 'rows', t: T.name, rows: D.visible(T.name, F) }; },
  DISTINCT: (D, a, F) => FN.VALUES(D, a, F),
  SELECTEDVALUE: (D, a, F, R) => { need(a, 1, 2, 'SELECTEDVALUE'); const { t, c } = colArg(D, a[0], 'SELECTEDVALUE'); const v = D.colValues(t, c, D.visible(t, F)); return v.length === 1 ? v[0] : a[1] ? D.ev(a[1], F, R) : BL; },
  HASONEVALUE: (D, a, F) => { const { t, c } = colArg(D, a[0], 'HASONEVALUE'); return D.colValues(t, c, D.visible(t, F)).length === 1; },
  RELATED: (D, a, F, R) => { need(a, 1, 1, 'RELATED'); const { t, c } = colArg(D, a[0], 'RELATED'); const rc = [...R].reverse().find(x => x.t); const no = () => err(`The column '${t}[${c}]' either doesn't exist or doesn't have a relationship to any table available in the current context.`); if (!rc) no();
    const P = D.M.paths(rc.t); if (!P.has(t) || t === rc.t) no(); let tab = rc.t, row = rc.i; for (const r of P.get(t)) { const f = splitKey(r.from); const v = D.cell(tab, row, f.c); const k = D.M.keyIndex(r.to).get(isBlank(v) ? null : v); if (k === undefined) return BL; tab = splitKey(r.to).t; row = k; } return D.cell(tab, row, c); },
  RELATEDTABLE: (D, a, F, R) => { const T = D.T(a[0].t); return { kind: 'rows', t: T.name, rows: D.visible(T.name, R.length ? D.transition(F, R) : F) }; },
  IF: (D, a, F, R) => { need(a, 2, 3, 'IF'); return truthy(D.ev(a[0], F, R)) ? D.ev(a[1], F, R) : a[2] ? D.ev(a[2], F, R) : BL; },
  SWITCH: (D, a, F, R) => { const x = D.ev(a[0], F, R); let k = 1; for (; k + 1 < a.length; k += 2) { const v = D.ev(a[k], F, R); if ((x === true && truthy(v)) || (x !== true && D.bin({ op: '=', a: { k: 'lit', v: x }, b: { k: 'lit', v } }, F, R))) return D.ev(a[k + 1], F, R); } return k < a.length ? D.ev(a[k], F, R) : BL; },
  AND: (D, a, F, R) => truthy(D.ev(a[0], F, R)) && truthy(D.ev(a[1], F, R)), OR: (D, a, F, R) => truthy(D.ev(a[0], F, R)) || truthy(D.ev(a[1], F, R)), NOT: (D, a, F, R) => !truthy(D.ev(a[0], F, R)),
  ISBLANK: (D, a, F, R) => isBlank(D.ev(a[0], F, R)), BLANK: () => BL, TRUE: () => true, FALSE: () => false,
  DIVIDE: (D, a, F, R) => { need(a, 2, 3, 'DIVIDE'); const x = D.ev(a[0], F, R), y = D.ev(a[1], F, R); if (isBlank(y) || num(y) === 0) return a[2] ? D.ev(a[2], F, R) : BL; if (isBlank(x)) return BL; return num(x) / num(y); },
  ROUND: (D, a, F, R) => { const x = D.ev(a[0], F, R); if (isBlank(x)) return BL; const p = Math.pow(10, num(D.ev(a[1], F, R))); return Math.sign(num(x)) * Math.round(Math.abs(num(x)) * p + 1e-9) / p; },
  ABS: (D, a, F, R) => { const x = D.ev(a[0], F, R); return isBlank(x) ? BL : Math.abs(num(x)); },
  YEAR: (D, a, F, R) => ymd(num(D.ev(a[0], F, R))).y, MONTH: (D, a, F, R) => ymd(num(D.ev(a[0], F, R))).m, DAY: (D, a, F, R) => ymd(num(D.ev(a[0], F, R))).d,
  DATE: (D, a, F, R) => ser(num(D.ev(a[0], F, R)), num(D.ev(a[1], F, R)), num(D.ev(a[2], F, R))),
  FORMAT: (D, a, F, R) => { const x = D.ev(a[0], F, R); const f = str(D.ev(a[1], F, R)); if (isBlank(x)) return ''; if (/%/.test(f)) return (num(x) * 100).toFixed((f.split('.')[1] || '').replace('%', '').length) + '%'; return String(x); },
  RANKX: (D, a, F, R) => { need(a, 2, 5, 'RANKX'); const tv = D.table(a[0], F, R); const vals = D.iterate(tv).map(r => D.ev(a[1], F, [...R, r])).filter(v => !isBlank(v)).map(num); const me = D.ev(a[1], F, R); if (isBlank(me)) return BL; const asc = a[3] && /ASC/i.test(a[3].k === 'table' ? a[3].t : str(D.ev(a[3], F, R)));
    return 1 + vals.filter(v => asc ? v < num(me) - 1e-9 : v > num(me) + 1e-9).length; },
  DATESYTD: (D, a, F, R) => { const d = D.dates(a[0], F, R); if (!d.vals.length) return { kind: 'vals', col: d.key, vals: [] }; const last = d.vals[d.vals.length - 1]; const start = fyStart(last, a[1] ? str(D.ev(a[1], F, R)) : '12/31'); return { kind: 'vals', col: d.key, vals: d.all.filter(v => v >= start && v <= last) }; },
  DATESMTD: (D, a, F, R) => { const d = D.dates(a[0], F, R); if (!d.vals.length) return { kind: 'vals', col: d.key, vals: [] }; const last = d.vals[d.vals.length - 1]; const y = ymd(last); const start = ser(y.y, y.m, 1); return { kind: 'vals', col: d.key, vals: d.all.filter(v => v >= start && v <= last) }; },
  DATESQTD: (D, a, F, R) => { const d = D.dates(a[0], F, R); if (!d.vals.length) return { kind: 'vals', col: d.key, vals: [] }; const last = d.vals[d.vals.length - 1]; const y = ymd(last); const start = ser(y.y, y.m - ((y.m - 1) % 3), 1); return { kind: 'vals', col: d.key, vals: d.all.filter(v => v >= start && v <= last) }; },
  TOTALYTD: totalTo('DATESYTD'), TOTALQTD: totalTo('DATESQTD'), TOTALMTD: totalTo('DATESMTD'),
  DATESINPERIOD: (D, a, F, R) => { need(a, 4, 4, 'DATESINPERIOD'); const d = D.dates(a[0], F, R); const at = num(D.ev(a[1], F, R)); const k = num(D.ev(a[2], F, R)); const unit = String(a[3].k === 'table' ? a[3].t : str(D.ev(a[3], F, R))).toUpperCase();
    const eom = n => ymd(n + 1).d === 1; const mv = n => { const x = shift(at, n); if (!eom(at)) return x; const y = ymd(x); return ser(y.m === 12 ? y.y + 1 : y.y, y.m === 12 ? 1 : y.m + 1, 1) - 1; };  // from a month end, months move month end to month end
    const end = unit === 'DAY' ? at + k : unit === 'MONTH' ? mv(k) : unit === 'QUARTER' ? mv(3 * k) : unit === 'YEAR' ? mv(12 * k) : err('DATESINPERIOD needs DAY, MONTH, QUARTER or YEAR.'); const lo = k < 0 ? end + 1 : at, hi = k < 0 ? at : end - 1; return { kind: 'vals', col: d.key, vals: d.all.filter(v => v >= lo && v <= hi) }; },
  DATESBETWEEN: (D, a, F, R) => { need(a, 3, 3, 'DATESBETWEEN'); const d = D.dates(a[0], F, R); const lo = D.ev(a[1], F, R), hi = D.ev(a[2], F, R); return { kind: 'vals', col: d.key, vals: d.all.filter(v => (isBlank(lo) || v >= num(lo)) && (isBlank(hi) || v <= num(hi))) }; },
  SAMEPERIODLASTYEAR: (D, a, F, R) => { const d = D.dates(a[0], F, R); return { kind: 'vals', col: d.key, vals: shiftSet(d.vals, -12, new Set(d.all)) }; },
  DATEADD: (D, a, F, R) => { need(a, 3, 3, 'DATEADD'); const d = D.dates(a[0], F, R); const k = num(D.ev(a[1], F, R)); const unit = String(a[2].k === 'table' ? a[2].t : str(D.ev(a[2], F, R))).toUpperCase(); const all = new Set(d.all);
    if (unit === 'DAY') return { kind: 'vals', col: d.key, vals: [...new Set(d.vals.map(v => v + k))].filter(v => all.has(v)) }; const m = unit === 'MONTH' ? k : unit === 'QUARTER' ? 3 * k : unit === 'YEAR' ? 12 * k : err('DATEADD needs DAY, MONTH, QUARTER or YEAR.'); return { kind: 'vals', col: d.key, vals: shiftSet(d.vals, m, all) }; },
  PREVIOUSMONTH: (D, a, F, R) => { const d = D.dates(a[0], F, R); if (!d.vals.length) return { kind: 'vals', col: d.key, vals: [] }; const f = ymd(d.vals[0]); const s = ser(f.m === 1 ? f.y - 1 : f.y, f.m === 1 ? 12 : f.m - 1, 1), e = ser(f.y, f.m, 1) - 1; return { kind: 'vals', col: d.key, vals: d.all.filter(v => v >= s && v <= e) }; },
};
// a value for display in a PivotTable cell, with a measure's format: {kind: 'general'|'number'|'currency'|'percent', dp}
export function fmtDax(v, fmt) { if (isBlank(v)) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'string') return v; if (!isFinite(v)) return isNaN(v) ? 'NaN' : v > 0 ? '∞' : '-∞';
  const f = fmt || {}; const dp = f.dp ?? (f.kind === 'percent' ? 2 : f.kind && f.kind !== 'general' ? 0 : null); const grp = x => { const [w, d] = Math.abs(x).toFixed(dp ?? 0).split('.'); const g = w.replace(/\B(?=(\d{3})+(?!\d))/g, ','); return (x < 0 ? '-' : '') + g + (d ? '.' + d : ''); };
  if (f.kind === 'percent') return grp(v * 100) + '%'; if (f.kind === 'currency') return '₹' + grp(v); if (f.kind === 'number') return grp(v); return Number.isInteger(v) ? String(v) : String(Math.round(v * 10000) / 10000); }
export { isBlank };
