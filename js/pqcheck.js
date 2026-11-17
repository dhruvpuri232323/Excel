// Checks Power Query lab tasks: runs the learner's query and the model answer on the lab's data, then on changed
// data (alsoWith), and compares the tables they give. Also turns a workbook's Tables into Power Query sources.
import { MEngine, MTable, MFn, MDate, MDateTime, MTime, MDuration, ErrCell, MError, showCell, mEquals, typeName, cellValue, mId } from './m.js';
import { parseRange, isDateFmt, isErr, colName } from './engine.js';
import { PQ_TYPES, syncedFunction } from './pq.js';

// every Table in the workbook, as Excel.CurrentWorkbook() sees it: dates come in as date-times, like in Excel
export function wbTables(wb) { const out = {}; if (!wb) return out;
  wb.sheets.forEach((s, si) => { for (const t of s.tables || []) { const g = parseRange(t.ref); const cols = []; for (let c = g.c1; c <= g.c2; c++) cols.push(String(wb.value(si, g.r1, c) ?? `Column${c - g.c1 + 1}`));
    const rows = []; for (let r = g.r1 + 1; r <= g.r2 - (t.totals ? 1 : 0); r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { let v = wb.value(si, r, c); if (isErr(v) || v === undefined || v === '') v = null; else if (typeof v === 'number') { const st = wb.styleAt ? wb.styleAt(si, r, c) : (wb.style ? wb.style(si, r, c) : null); if (st && st.fmt && isDateFmt(st.fmt)) v = new MDateTime(v); } row.push(v); } rows.push(row); }
    out[t.name] = new MTable(cols, rows, cols.map(() => null)); } });
  return out; }
export function pqEngine(ctx, override) { const pq = ctx.pq || {}; const base = { tables: { ...wbTables(ctx.wb), ...(pq.tables || {}) }, files: pq.files || {}, folders: pq.folders || {}, culture: pq.culture, now: pq.now, looseFolders: !!pq.looseFolders };
  if (override) { if (override.tables) base.tables = { ...base.tables, ...override.tables }; if (override.files) base.files = { ...base.files, ...override.files }; if (override.folders) base.folders = { ...base.folders, ...override.folders }; }
  const e = new MEngine(base); for (const n of pq.order || Object.keys(pq.queries || {})) if (pq.queries[n] !== undefined) e.setQuery(n, pq.queries[n]);
  if (override && override.queries) for (const [n, q] of Object.entries(override.queries)) e.setQuery(n, q);  // alsoWith can change a parameter, say
  const f = syncedFunction('Transform Sample File', e.queries['Transform Sample File'], e.queries); if (f) e.setQuery(f[0], f[1]); return e; }
const tyName = k => k && PQ_TYPES[k] ? PQ_TYPES[k][1] : 'Any (ABC123)';
const shown = v => { const s = showCell(v, 'en-IN'); return s.cls === 'txt' ? `“${s.t}”` : s.t; };
const near = (a, b) => { a = a instanceof ErrCell ? a : cellValue(a); b = b instanceof ErrCell ? b : cellValue(b); if (a instanceof ErrCell || b instanceof ErrCell) return a instanceof ErrCell && b instanceof ErrCell;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-7 * Math.max(1, Math.abs(b)); if ((a instanceof MDate || a instanceof MDateTime) && (b instanceof MDate || b instanceof MDateTime)) return a.constructor === b.constructor && Math.abs(a.n - b.n) < 1e-6; return mEquals(a, b); };
const rowKey = (r, idx) => idx.map(i => { const v = r[i]; if (v instanceof ErrCell) return 'ERR'; const x = cellValue(v); if (typeof x === 'number') return 'n' + Number(x.toPrecision(10)); return (x && x.constructor ? x.constructor.name : '') + ':' + showCell(x, 'en-IN').t; }).join('\u0001');
// compare a learner's table with the answer's: returns null when they match, or a message
export function compareTables(got, want, o = {}) {
  if (!(got instanceof MTable)) return `Your query gives ${typeName(got) === 'Null' ? 'nothing' : 'a ' + typeName(got).toLowerCase()}, not a table.`;
  const miss = want.cols.filter(c => !got.cols.includes(c)); const extra = got.cols.filter(c => !want.cols.includes(c));
  if (miss.length && extra.length === miss.length && miss.length <= 2) return `Column${miss.length > 1 ? 's' : ''} ${miss.map(c => `“${c}”`).join(' and ')} ${miss.length > 1 ? 'are' : 'is'} missing, and your query has ${extra.map(c => `“${c}”`).join(' and ')} instead. Check the spelling and capitals of the column name${miss.length > 1 ? 's' : ''}.`;
  if (miss.length) return `Column${miss.length > 1 ? 's' : ''} ${miss.map(c => `“${c}”`).join(', ')} ${miss.length > 1 ? 'are' : 'is'} missing.`;
  if (extra.length && !o.extraOk) return `Your query keeps ${extra.map(c => `“${c}”`).join(', ')}, which the answer doesn’t: remove ${extra.length > 1 ? 'them' : 'it'}.`;
  if (o.colOrder !== false && !o.extraOk && want.cols.some((c, i) => got.cols[i] !== c)) return `The columns are all there but in a different order. The answer has: ${want.cols.join(', ')}.`;
  // errors in the learner's cells come first: they explain more than a type or a value would
  for (const [j, c] of want.cols.entries()) { const g = got.cols.indexOf(c); const r = got.rows.findIndex(x => x[g] instanceof ErrCell); if (r >= 0 && !want.rows.some(x => x[j] instanceof ErrCell)) return cellMsg(r, c, got.rows[r][g], null); }
  if (o.types) { for (const c of want.cols) { const kw = want.types[want.cols.indexOf(c)], kg = got.types[got.cols.indexOf(c)]; const norm = k => k === 'Int64' && o.types === 'loose' ? 'number' : k; if ((Array.isArray(o.types) ? o.types.includes(c) : true) && norm(kw) !== norm(kg)) return `“${c}” is ${tyName(kg)}; set it to ${tyName(kw)}.`; } }
  if (got.rows.length !== want.rows.length) return `Your query gives ${got.rows.length} row${got.rows.length === 1 ? '' : 's'}; the answer has ${want.rows.length}.${got.rows.length > want.rows.length ? ' Some rows that should go are still there.' : ' Some rows that should stay have gone.'}`;
  const gi = want.cols.map(c => got.cols.indexOf(c)), wi = want.cols.map((c, i) => i);
  if (o.order) { for (let r = 0; r < want.rows.length; r++) for (let j = 0; j < want.cols.length; j++) { const a = got.rows[r][gi[j]], b = want.rows[r][j]; if (!near(a, b)) return cellMsg(r, want.cols[j], a, b); } return null; }
  const pool = new Map(); for (const r of got.rows) { const k = rowKey(r, gi); pool.set(k, (pool.get(k) || 0) + 1); }
  for (let r = 0; r < want.rows.length; r++) { const k = rowKey(want.rows[r], wi); if (pool.get(k)) { pool.set(k, pool.get(k) - 1); continue; }
    // the learner's row that is nearest: all but one cell the same, preferring a row that matches no answer row and a cell that differs only in spaces or capitals
    const wantKeys = new Set(want.rows.map(x => rowKey(x, wi))); const soft = v => { v = cellValue(v); return typeof v === 'string' ? v.replace(/\s+/g, ' ').trim().toLowerCase() : null; };
    const cands = got.rows.map((g, i) => i).filter(i => wi.filter(j => near(got.rows[i][gi[j]], want.rows[r][j])).length >= want.cols.length - 1);
    const score = i => { const j = wi.find(j => !near(got.rows[i][gi[j]], want.rows[r][j])); return (wantKeys.has(rowKey(got.rows[i], gi)) ? 0 : 2) + (j !== undefined && soft(got.rows[i][gi[j]]) !== null && soft(got.rows[i][gi[j]]) === soft(want.rows[r][j]) ? 1 : 0); };
    const pos = cands.length ? cands.reduce((a, b) => score(b) > score(a) ? b : a) : -1; if (pos >= 0) { const j = wi.find(j => !near(got.rows[pos][gi[j]], want.rows[r][j])); return cellMsg(pos, want.cols[j], got.rows[pos][gi[j]], want.rows[r][j]); }
    return `The answer has a row ${want.cols.slice(0, 4).map((c, j) => `${c} = ${shown(want.rows[r][j])}`).join(', ')} that your query doesn’t.`; }
  return null; }
function cellMsg(r, c, a, b) { if (a instanceof ErrCell) return `Row ${r + 1} of “${c}” shows Error: ${a.e.reason}: ${a.e.message}`; const x = cellValue(a), y = cellValue(b);
  if (x !== null && y !== null && typeof x !== typeof y) return `Row ${r + 1}, “${c}”: yours is ${typeName(x).toLowerCase()} ${shown(x)}; the answer is ${typeName(y).toLowerCase()} ${shown(y)}. Check the column’s data type.`;
  if (x instanceof MDateTime && y instanceof MDate) return `Row ${r + 1}, “${c}”: yours is a date and time; the answer is a date. Change the column’s type to Date.`;
  const hint = typeof x === 'string' && typeof y === 'string' ? (x.trim() === y.trim() || x.replace(/\s+/g, ' ') === y.replace(/\s+/g, ' ') ? ' Yours has extra spaces.' : x.toLowerCase() === y.toLowerCase() ? ' Only the capitals differ.' : '') : '';
  return `Row ${r + 1}, “${c}”: yours shows ${shown(x)}, the answer shows ${shown(y)}.${hint}`; }
export function checkPQ(task, ctx) {
  const fail = msg => ({ ok: false, msg }); const pq = ctx.pq; if (!pq) return fail('Open the Power Query Editor first.');
  if (task.kind === 'pqload') { const l = (pq.loads || {})[task.query]; if (!pq.queries[task.query]) return fail(`There’s no query called “${task.query}”.`); if (!l) return fail(`Close & Load the query “${task.query}” first.`);
    if (task.to && l.to !== task.to) return fail(task.to === 'connection' ? `“${task.query}” is loaded to a sheet. Use Close & Load To… › Only Create Connection.` : `“${task.query}” is only a connection. Use Close & Load To… › Table.`); if (task.model !== undefined && !!l.model !== !!task.model) return fail(task.model ? 'Tick “Add this data to the Data Model” in Import Data.' : 'Untick “Add this data to the Data Model”.'); return { ok: true, msg: task.okMsg || 'Loaded the right way.' }; }
  const qn = task.query; if (!pq.queries[qn]) { const near = Object.keys(pq.queries).find(n => n.toLowerCase() === String(qn).toLowerCase()); return fail(near ? `Rename the query “${near}” to “${qn}”, with the same capitals (Query Settings › Name).` : `There’s no query called “${qn}” yet.${task.create ? ' ' + task.create : ''}`); }
  const text = pq.queries[qn]; for (const s of task.mustUse || []) if (!text.includes(s)) return fail(task.mustUseMsg || `Use ${s} in this query.`); for (const s of task.mustNot || []) if (text.includes(s)) return fail(task.mustNotMsg || `Don’t use ${s} here.`);
  // the answer runs beside the learner's other queries, except those the task brings its own answers for (refQueries: the helper queries of Combine Files, say)
  const run = override => { const e = pqEngine(ctx, override); let got, want; try { got = e.evaluate(qn); } catch (er) { return { err: er }; } let ea = e;
    if (task.refQueries) { const q = { ...pq.queries, ...task.refQueries }; ea = pqEngine({ ...ctx, pq: { ...pq, queries: q, order: [...new Set([...(pq.order || Object.keys(pq.queries)), ...Object.keys(task.refQueries)])] } }, override); }
    ea.setQuery('__answer', task.ref); try { want = ea.evaluate('__answer'); if (want instanceof MFn) return { got, want, e, ea }; } catch (er) { return { fatal: 'The model answer failed: ' + (er.message || er) }; } return { got, want }; };
  const r0 = run(null); if (r0.fatal) return fail(r0.fatal); if (r0.err) return fail(`Your query stops with an error: ${r0.err.reason || 'Expression.Error'}: ${r0.err.message} Click each step in Applied Steps to find where it starts.`);
  // a custom function: call the learner's and the answer's on the same arguments (task.fnArgs, as M expressions)
  if (r0.want instanceof MFn) { if (!(r0.got instanceof MFn)) return fail(`“${qn}” gives ${shown(r0.got)}, not a function. A function query starts with its parameters, such as (d as date) => …`);
    for (const a of task.fnArgs || []) { let x, y; try { y = r0.ea.eval(`__answer(${a})`); } catch (er) { return fail('The model answer failed: ' + er.message); } try { x = r0.e.eval(`${mId(qn)}(${a})`); } catch (er) { return fail(`${qn}(${a}) stops with an error: ${er.message}`); }
      if (!near(x, y)) return fail(`${qn}(${a}) gives ${shown(x)}; it should give ${shown(y)}.`); }
    return { ok: true, msg: task.okMsg || 'The function gives the right answer for every test value.' }; }
  if (task.value !== undefined || !(r0.want instanceof MTable)) { const ok = near(r0.got, r0.want); if (!ok) return fail(`Your query gives ${shown(r0.got)}; the answer is ${shown(r0.want)}.`); }
  else { const m = compareTables(r0.got, r0.want, task.compare || {}); if (m) return fail(m); }
  for (const [k, alt] of (task.alsoWith || []).entries()) { const r = run(alt); if (r.fatal) return fail(r.fatal); const why = alt.say ? ` (${alt.say})` : '';
    if (r.err) return fail(`Your query works on this data, but on other data${why} it stops with an error: ${r.err.message} ${task.alsoMsg || 'Make each step work for any data, not just these rows.'}`);
    const m = r.want instanceof MTable ? compareTables(r.got, r.want, task.compare || {}) : (near(r.got, r.want) ? null : `it gives ${shown(r.got)}, not ${shown(r.want)}`); if (m) return fail(`Your query works on this data, but not on other data${why}: ${m.replace(/^Your query /, 'it ')} ${task.alsoMsg || ''}`.trim()); }
  if (task.load) { const l = (pq.loads || {})[qn]; if (!l) return fail(`The query is right. Now Close & Load it.`); }
  const what = r0.want instanceof MTable ? 'table' : 'value'; return { ok: true, msg: task.okMsg || ((task.alsoWith || []).length ? `That’s the ${what} the answer gives, and it still works when the data changes.` : `That’s the ${what} the answer gives.`) };
}
// the lab's starting state, for applySolution and tests
export function pqStart(lab) { const p = lab.pq || {}; return { queries: { ...(p.queries || {}) }, order: (p.order || Object.keys(p.queries || {})).slice(), active: p.active || (p.order || Object.keys(p.queries || {}))[0], loads: { ...(p.loads || {}) }, files: p.files, folders: p.folders, tables: p.tables, culture: p.culture, now: p.now }; }
export function pqApply(task, pq) { if (task.kind === 'pqload') { pq.loads[task.query] = { ...(pq.loads[task.query] || {}), to: task.to || 'table', model: !!task.model }; if (!pq.queries[task.query] && task.ref) { pq.queries[task.query] = task.ref; pq.order.push(task.query); } return; }
  for (const [n, q] of Object.entries(task.refQueries || {})) { if (pq.queries[n] === undefined) pq.order.push(n); pq.queries[n] = q; if (/^(Sample File|Parameter1|Transform Sample File|Transform File)$/.test(n)) pq.loads[n] = pq.loads[n] || { to: 'connection' }; }
  if (!pq.queries[task.query]) pq.order.push(task.query); pq.queries[task.query] = task.solution || task.ref; for (const [a, b] of Object.entries(task.renames || {})) { if (pq.queries[a] !== undefined && a !== task.query) { delete pq.queries[a]; pq.order = pq.order.filter(x => x !== a); } void b; } if (task.load) pq.loads[task.query] = { to: 'table' }; }

// Close & Load: each query loaded "to a table" gets its own sheet, named after the query, with a Table on it
export function loadQueries(wb, ctx, only) { const e = pqEngine(ctx); const st = {}; const S = ctx.pq;
  for (const n of S.order || Object.keys(S.queries)) { const l = (S.loads || {})[n]; if (!l || (only && !only.includes(n))) continue; if (l.to !== 'table') { st[n] = {}; continue; }
    let t; try { t = e.evaluate(n); } catch (er) { st[n] = { err: er.message }; continue; } if (!(t instanceof MTable)) { st[n] = { err: `the query gives a ${typeName(t).toLowerCase()}, not a table` }; continue; }
    const si = writeQueryTable(wb, n, t, S.culture, l.sheet); if (S.loads[n].sheet !== wb.sheets[si].name) S.loads[n] = { ...l, sheet: wb.sheets[si].name }; st[n] = { rows: t.rows.length, errors: t.rows.reduce((a, r) => a + r.filter(v => v instanceof ErrCell).length, 0) }; }
  wb.invalidate(); return st; }
// a loaded query gets a sheet of its own, named after it, as in Excel: “Orders (2)” when a sheet called Orders is already there; a refresh writes over that same sheet
export function writeQueryTable(wb, name, t, culture, sheetName) { let si = sheetName ? wb.sheets.findIndex(s => s.name === sheetName) : -1; let s;
  if (si < 0) { const b = String(name).slice(0, 31); let sn = b, k = 2; while (wb.sheets.some(x => x.name.toLowerCase() === sn.toLowerCase())) sn = `${b.slice(0, 26)} (${k++})`; s = wb.addSheet(sn); si = wb.sheets.length - 1; } else { s = wb.sheets[si]; s.cells = new Map(); s.tables = []; }
  const hdr = { bold: true, fill: '#4472C4', color: '#FFFFFF' }; t.cols.forEach((c, j) => wb.setRaw(s, 0, j, { v: c, ...hdr }, true));
  t.rows.forEach((r, i) => r.forEach((v, j) => { const x = v instanceof ErrCell ? null : cellValue(v); const band = i % 2 ? {} : { fill: '#D9E1F2' }; let cell;
    if (x === null || x === undefined) cell = { v: null, ...band }; else if (x instanceof MDate) cell = { v: x.n, fmt: 'dd-mm-yyyy', ...band }; else if (x instanceof MDateTime) cell = { v: x.n, fmt: 'dd-mm-yyyy hh:mm', ...band }; else if (x instanceof MTime) cell = { v: x.n, fmt: 'hh:mm:ss', ...band }; else if (x instanceof MDuration) cell = { v: x.n, fmt: '[h]:mm:ss', ...band };
    else if (typeof x === 'object') cell = { v: `[${typeName(x)}]`, ...band }; else cell = { v: x, ...band }; wb.setRaw(s, i + 1, j, cell, true); }));
  t.cols.forEach((c, j) => { s.widths[j] = Math.min(220, Math.max(70, 8 * Math.max(String(c).length, ...t.rows.slice(0, 50).map(r => showCell(r[j], culture).t.length)) + 16)); });
  let tn = String(name).replace(/[^A-Za-z0-9_.]/g, '_').replace(/^([^A-Za-z_])/, '_$1'); const taken = n => wb.sheets.some((x, k) => k !== si && (x.tables || []).some(q => q.name.toLowerCase() === n.toLowerCase())); if (taken(tn)) { let k = 2; while (taken(`${tn}_${k}`)) k++; tn = `${tn}_${k}`; }
  s.tables = [{ name: tn, ref: `A1:${colName(Math.max(0, t.cols.length - 1))}${Math.max(2, t.rows.length + 1)}` }]; s.rows = Math.max(s.rows, t.rows.length + 21); s.cols = Math.max(s.cols, t.cols.length + 6); return si; }
// the model answer as a workbook: every query's result on its own sheet, and the M code on a sheet of its own
export function solvedWorkbook(wb, lab, tasks) { const pq = pqStart(lab); for (const t of tasks) if (t.kind === 'pq' || t.kind === 'pqload') pqApply(t, pq);
  const want = new Set(tasks.filter(t => t.kind === 'pq' || t.kind === 'pqload').map(t => t.query)); for (const n of pq.order) if (want.has(n) && !(pq.loads[n] && pq.loads[n].to === 'connection')) pq.loads[n] = pq.loads[n] || { to: 'table' };
  loadQueries(wb, { wb, pq }); const s = wb.addSheet('Queries (M)'); let r = 0; const put = (v, st) => { wb.setRaw(s, r++, 0, st ? { v, ...st } : v, true); };
  put('The model answer’s Power Query queries. To use one: Data › Get Data › Blank Query, then Advanced Editor, and paste the code.', { italic: true });
  for (const n of pq.order) { r++; put(n, { bold: true, fill: '#DDEBF7' }); for (const line of String(pq.queries[n]).split(/\r?\n/)) put(line || ' '); } s.widths[0] = 900; wb.invalidate(); return pq; }

// a standalone task (module tests, the exam) starts from the query as it was just before the task (task.before)
export function pqForTask(labPq, t) { if (!labPq) return null; if (!t.before) return labPq; const q = { ...(labPq.queries || {}) }; q[t.query] = t.before; const order = (labPq.order || Object.keys(labPq.queries || {})).slice(); if (!order.includes(t.query)) order.push(t.query);
  for (const [n, x] of Object.entries(t.beforeQueries || {})) { q[n] = x; if (!order.includes(n)) order.push(n); }  // other queries the task needs, such as a parameter made in an earlier task
  return { ...labPq, queries: q, order, active: t.query, start: 'editor', loads: {} }; }
