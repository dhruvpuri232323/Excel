// The Data Model (Power Pivot) in the browser: tables, relationships, a marked date table, Sort by Column, and
// PivotTables over several tables. A value from one table is filtered by a field from another only when a chain of
// active many-to-one relationships leads from the value's table to the field's table; otherwise every row shows the
// same total and the field list warns "Relationships between tables may be needed", as in Excel.
import { MDate, MDateTime, ser, ymd, toText } from './m.js';
import { Dax, BL, emptyF, fmtDax, DaxError } from './dax.js';

export const BLANK = '(blank)';
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const parseField = f => { const m = /^'?([^'[]+)'?\[([^\]]+)\]$/.exec(String(f).trim()); return m ? { t: m[1], c: m[2] } : null; };
export const fieldRef = (t, c) => `${t}[${c}]`;
export const FN = { sum: 'Sum', count: 'Count', average: 'Average', min: 'Min', max: 'Max', distinct: 'Distinct Count' };
export const valueName = v => v.kpi ? (v.part === 'status' ? `${v.kpi} Status` : v.part === 'goal' ? `${v.kpi} Goal` : v.kpi) : v.measure || v.name || `${FN[v.fn || 'sum']} of ${parseField(v.field).c}`;

// values: dates as day numbers with a 'date' type; {d}/{dt} from Python become dates
function cellIn(v) { if (v && typeof v === 'object') { if (v.d) { const [y, m, d] = v.d.split('-').map(Number); return ser(y, m, d); } if (v.dt) { const [a] = v.dt.split('T'); const [y, m, d] = a.split('-').map(Number); return ser(y, m, d); } if (v instanceof MDate || v instanceof MDateTime) return Math.floor(v.n); } return v === '' ? null : v; }
function typeOf(vals, raw) { const nn = raw.filter(v => v !== null && v !== undefined && v !== ''); if (!nn.length) return 'text'; if (nn.every(v => v && typeof v === 'object' && (v.d || v.dt || v instanceof MDate || v instanceof MDateTime))) return 'date'; if (nn.every(v => typeof v === 'number')) return Number.isInteger(nn[0]) && nn.every(Number.isInteger) ? 'whole' : 'decimal'; return 'text'; }
export function showVal(v, type) { if (v === null || v === undefined || v === BLANK) return v === BLANK ? BLANK : ''; if (type === 'date') { const d = ymd(v); return `${String(d.d).padStart(2, '0')}-${String(d.m).padStart(2, '0')}-${d.y}`; } if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100); return String(v); }

export class Model {
  // spec: { tables: {name: {cols, rows, types?}}, order? } · state: { rels: [{from, to, active}], dateTable: {table, col} | null, sortBy: {'T[C]': 'T[D]'}, added: {name: table} }
  constructor(spec, state = {}) {
    this.spec = spec || {}; this.tables = {}; const src = { ...(spec.tables || {}), ...(state.added || {}) };
    for (const [n, t] of Object.entries(src)) { const rows = (t.rows || []).map(r => r.map(cellIn)); const types = t.cols.map((c, j) => (t.types && t.types[j]) || typeOf(rows.map(r => r[j]), (t.rows || []).map(r => r[j]))); this.tables[n] = { name: n, cols: t.cols.slice(), types, rows }; }
    this.order = [...new Set([...(spec.order || Object.keys(spec.tables || {})), ...Object.keys(state.added || {})])].filter(n => this.tables[n]);
    this.rels = (state.rels || []).map(r => ({ from: r.from, to: r.to, active: r.active !== false }));
    this.dateTable = state.dateTable || null; this.sortBy = { ...(state.sortBy || {}) }; this._idx = new Map();
    this.measures = (state.measures || []).map(m => ({ ...m })); this.calcCols = (state.calcCols || []).map(c => ({ ...c })); this.calcErrors = {}; this.kpis = (state.kpis || []).map(k => ({ ...k })); this.hidden = [...(state.hidden || [])];
    for (const cc of this.calcCols) this.addCalc(cc); }
  state() { return { rels: this.rels.map(r => ({ ...r })), dateTable: this.dateTable, sortBy: { ...this.sortBy }, measures: this.measures.map(m => ({ ...m })), calcCols: this.calcCols.map(c => ({ ...c })), kpis: this.kpis.map(k => ({ ...k })), hidden: [...this.hidden] }; }
  dax() { return new Dax(this, this.measures); }
  // a calculated column: the formula runs once for each row (row context), and the result is stored like any column
  addCalc(cc) { const T = this.tables[cc.table]; if (!T) return; const D = this.dax(); let vals; try { vals = T.rows.map((_, i) => { const v = D.value(cc.expr, emptyF(), [{ t: T.name, i }]); return v === BL ? null : v; }); } catch (e) { this.calcErrors[cc.name] = e.message; vals = T.rows.map(() => null); }
    const j = T.cols.indexOf(cc.name); if (j >= 0) { T.rows.forEach((r, i) => { r[j] = vals[i]; }); } else { T.cols.push(cc.name); T.rows.forEach((r, i) => r.push(vals[i])); T.types.push(typeOf(vals, vals)); } this._idx.clear(); }
  measureOf(name) { return this.measures.find(m => m.name.toLowerCase() === String(name).toLowerCase()); }
  // a KPI (Power Pivot › KPIs): a base measure, a target (a measure or an absolute value) and two status thresholds, as fractions of the target
  kpiOf(name) { const k = this.kpis.find(x => x.base.toLowerCase() === String(name).toLowerCase()); return k && this.measureOf(k.base) ? k : null; }
  has(f) { const p = parseField(f); return !!(p && this.tables[p.t] && this.tables[p.t].cols.includes(p.c)); }
  col(t, c) { const T = this.tables[t]; const j = T.cols.indexOf(c); return T.rows.map(r => r[j]); }
  typeOfField(f) { const p = parseField(f); const T = this.tables[p.t]; return T ? T.types[T.cols.indexOf(p.c)] : null; }
  unique(f) { const p = parseField(f); const v = this.col(p.t, p.c).filter(x => x !== null && x !== undefined); return new Set(v).size === v.length; }
  // ---- relationships
  relKey(r) { return `${r.from}->${r.to}`; }
  findRel(a, b) { return this.rels.find(r => (r.from === a && r.to === b) || (r.from === b && r.to === a)); }
  // Excel's Create Relationship: Table/Column (the many side) and Related Table/Related Column (the one side); it turns a backwards pair round
  addRel(a, b) { if (!this.has(a) || !this.has(b)) return { err: 'Pick a column in each table.' }; const pa = parseField(a), pb = parseField(b);
    if (pa.t === pb.t) return { err: 'Pick columns from two different tables.' };
    if (this.findRel(a, b)) return { err: 'That relationship already exists.' };
    if (this.typeOfField(a) !== this.typeOfField(b) && !(/whole|decimal/.test(this.typeOfField(a)) && /whole|decimal/.test(this.typeOfField(b)))) return { err: 'The relationship cannot be created because the two columns have different data types.' };
    let from = a, to = b, note = null; if (!this.unique(to)) { if (this.unique(from)) { from = b; to = a; note = 'The related column had duplicates, so the relationship was created the other way round: the column with unique values is the one side.'; } else return { err: 'The relationship cannot be created because each column contains duplicate values. Select at least one column that contains only unique values.' }; }
    const tf = parseField(from).t, tt = parseField(to).t; const active = !this.reachable(tf, tt);  // one active path between two tables
    this.rels.push({ from, to, active }); this._idx.clear(); return { ok: true, note: active ? note : (note ? note + ' ' : '') + 'There’s already an active path between these tables, so this relationship is inactive.', rel: { from, to, active } }; }
  removeRel(i) { this.rels.splice(i, 1); this._idx.clear(); }
  setActive(i, on) { this.rels[i].active = on; this._idx.clear(); }
  // the tables reachable from t along active many-to-one relationships, with the path to each
  paths(t) { const out = new Map([[t, []]]); const q = [t]; while (q.length) { const x = q.shift(); for (const r of this.rels) { if (!r.active) continue; const f = parseField(r.from), to = parseField(r.to); if (f.t === x && !out.has(to.t)) { out.set(to.t, [...out.get(x), r]); q.push(to.t); } } } return out; }
  reachable(a, b) { return this.paths(a).has(b); }
  keyIndex(f) { if (this._idx.has(f)) return this._idx.get(f); const p = parseField(f); const m = new Map(); this.col(p.t, p.c).forEach((v, i) => { if (v !== null && !m.has(v)) m.set(v, i); }); this._idx.set(f, m); return m; }
  // the value of field f for row i of table t, following relationships; BLANK when there's no match
  valueFor(f, t, i, path) { const p = parseField(f); let tab = t, row = i; for (const r of path) { const fr = parseField(r.from); const v = this.tables[tab].rows[row][this.tables[tab].cols.indexOf(fr.c)]; const k = this.keyIndex(r.to).get(v); if (k === undefined) return BLANK; tab = parseField(r.to).t; row = k; } const v = this.tables[p.t].rows[row][this.tables[p.t].cols.indexOf(p.c)]; return v === null || v === undefined ? BLANK : v; }
  // ---- sorting labels: by the Sort by Column if there is one
  sorter(f) { const by = this.sortBy[f]; const p = parseField(f); let key = v => v; if (by && this.has(by)) { const pb = parseField(by); const m = new Map(); const a = this.col(p.t, p.c), b = this.col(pb.t, pb.c); a.forEach((v, i) => { if (!m.has(v)) m.set(v, b[i]); }); key = v => m.has(v) ? m.get(v) : v; }
    return (x, y) => { if (x === BLANK) return y === BLANK ? 0 : 1; if (y === BLANK) return -1; const a = key(x), b = key(y); if (typeof a === 'number' && typeof b === 'number') return a - b; return String(a).localeCompare(String(b), 'en', { numeric: true }); }; }
  // ---- PivotTable: layout {rows: [field], cols: [field], filters: [{field, values?}], values: [{field, fn} | {measure}]}
  // every cell is evaluated as DAX in its own filter context: the row and column labels, plus the report's filters
  pivot(L) { const rows = (L.rows || []).filter(f => this.has(f)), cols = (L.cols || []).filter(f => this.has(f)); const vals = (L.values || []).filter(v => v.kpi ? this.kpiOf(v.kpi) : v.measure ? this.measureOf(v.measure) : this.has(v.field)); const filters = (L.filters || []).filter(f => this.has(f.field));
    const res = { rows: [], cols: [[]], cell: new Map(), names: vals.map(valueName), fmts: vals.map(v => { if (v.kpi) { const K = this.kpiOf(v.kpi); if (v.part === 'status') return { kind: 'kpi', style: K.style || 'lights' }; if (v.part === 'goal' && K.target.measure) return (this.measureOf(K.target.measure) || {}).fmt || null; return this.measureOf(K.base).fmt || null; } return v.measure ? (this.measureOf(v.measure).fmt || null) : null; }), warn: false, errors: [], fields: { rows, cols } };
    const D = this.dax(); const IMP = { sum: 'SUM', count: 'COUNTA', average: 'AVERAGE', min: 'MIN', max: 'MAX', distinct: 'DISTINCTCOUNT' };
    const home = v => v.kpi ? this.measureOf(this.kpiOf(v.kpi).base).table : v.measure ? this.measureOf(v.measure).table : parseField(v.field).t; const exprOf = v => v.measure ? null : `${IMP[v.fn || 'sum']}('${parseField(v.field).t}'[${parseField(v.field).c}])`;
    for (const v of vals) if (!v.measure && !v.kpi) for (const f of [...rows, ...cols, ...filters.map(x => x.field)]) { const ft = parseField(f).t; if (ft !== home(v) && !this.reachable(home(v), ft)) res.warn = true; }  // explicit measures may ignore a table on purpose: only dragged fields warn
    // the members of a field: its values, and (blank) when some fact rows have no match or the column has empty cells
    const members = f => { const p = parseField(f); const raw = this.col(p.t, p.c); const s = [...new Set(raw.filter(x => x !== null && x !== undefined))]; let blank = raw.some(x => x === null || x === undefined);
      if (!blank) for (const v of vals) { const h = home(v); const P = this.paths(h); if (h === p.t || !P.has(p.t)) continue; const path = P.get(p.t); if (this.tables[h].rows.some((_, i) => this.valueFor(fieldRef(p.t, p.c), h, i, path) === BLANK && this._unmatched(h, i, path))) blank = true; }
      if (blank) s.push(BLANK); return s.sort(this.sorter(f)); };
    const cart = fields => { let out = [[]]; for (const f of fields) { const m = members(f); const next = []; for (const k of out) for (const x of m) next.push([...k, x]); out = next; } return out; };
    const base = emptyF(); for (const f of filters) if (f.values) base.cols.set(f.field, new Set(f.values.map(x => x === BLANK ? BL : x))); base.outer = new Map(base.cols);
    const ctx = (rk, ck) => { const F = { cols: new Map(base.cols), tabs: new Map(), outer: base.outer }; (rk || []).forEach((x, j) => F.cols.set(rows[j], new Set([x === BLANK ? BL : x]))); (ck || []).forEach((x, j) => F.cols.set(cols[j], new Set([x === BLANK ? BL : x]))); return F; };
    const evalv = (v, F) => { try { const r = v.kpi ? this.kpiValue(D, this.kpiOf(v.kpi), v.part, F) : v.measure ? D.measure(v.measure, F) : D.value(exprOf(v), F); return r === BL ? null : r; } catch (e) { if (!res.errors.includes(e.message)) res.errors.push(e.message); return null; } };
    if (!vals.length) { res.rows = rows.length ? cart(rows) : [[]]; return res; }
    let rks = rows.length ? cart(rows) : [[]], cks = cols.length ? cart(cols) : [[]];
    const put = (rk, ck) => vals.forEach((v, k) => res.cell.set(JSON.stringify([rk, ck, k]), evalv(v, ctx(rk, ck))));
    for (const rk of [...rks, null]) for (const ck of [...cks, null]) put(rk, ck);
    const live = (rk, ck) => vals.some((_, k) => res.cell.get(JSON.stringify([rk, ck, k])) !== null);
    if (rows.length) rks = rks.filter(rk => [...cks, null].some(ck => live(rk, ck))); if (cols.length) cks = cks.filter(ck => [...rks, null].some(rk => live(rk, ck)));
    res.rows = rks; res.cols = cks; return res; }
  // a KPI part in a filter context: the base measure, the goal, or the status (-1 below the lower threshold, 0 between, 1 at or above the upper)
  kpiValue(D, K, part, F) { const val = () => D.measure(K.base, F); const goal = () => K.target.measure ? D.measure(K.target.measure, F) : K.target.value; if (part === 'goal') return goal(); if (part !== 'status') return val();
    const a = val(), g = goal(); if (a === BL || g === BL || !g) return BL; const q = a / g; return q < K.lo ? -1 : q < K.hi ? 0 : 1; }
  _unmatched(t, i, path) { let tab = t, row = i; for (const r of path) { const f = parseField(r.from); const v = this.tables[tab].rows[row][this.tables[tab].cols.indexOf(f.c)]; const k = this.keyIndex(r.to).get(v); if (k === undefined) return true; tab = parseField(r.to).t; row = k; } return false; }
  get(res, rk, ck, k) { return res.cell.get(JSON.stringify([rk, ck, k || 0])); }
  // a calendar like Power Pivot's Date Table › New: whole years around the dates in the model
  calendar(name = 'Calendar') { let lo = Infinity, hi = -Infinity; for (const T of Object.values(this.tables)) T.types.forEach((ty, j) => { if (ty === 'date') for (const r of T.rows) if (typeof r[j] === 'number') { lo = Math.min(lo, r[j]); hi = Math.max(hi, r[j]); } });
    if (!isFinite(lo)) return null; const y1 = ymd(lo).y, y2 = ymd(hi).y; const rows = []; for (let n = ser(y1, 1, 1); n <= ser(y2, 12, 31); n++) { const d = ymd(n); rows.push([n, d.y, d.m, MON[d.m - 1], `${MON[d.m - 1].slice(0, 3)}-${d.y}`, ((n + 6) % 7) + 1, ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][(n + 6) % 7]]); }
    return { name, cols: ['Date', 'Year', 'Month Number', 'Month', 'MMM-YYYY', 'Day Of Week Number', 'Day Of Week'], types: ['date', 'whole', 'whole', 'text', 'text', 'whole', 'text'], rows }; }
}
export { toText, fmtDax, DaxError };
