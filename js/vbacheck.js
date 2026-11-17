// Checking VBA work: the learner's macro and the model answer each run on a fresh copy of the lab's workbook
// (and on the task's variants: more rows, other answers to an InputBox, another active cell), and the results are compared:
// cells, formatting, sheets, files saved, Debug.Print output, message boxes. Structure checks read the code itself.
import { Program, Session, runMacro, CompileError, VbaError, EMPTY, MISSING, VArray, VDate, isObj, toStr, numStr } from './vba.js';
import { ExcelHost, Rng } from './vbaxl.js';
import { Workbook, parseRange, parseAddr, addr, rangeStr, isErr, formatValue, ERR, mkArr } from './engine.js';

export const VBA_KINDS = ['vba', 'vbafn', 'vbacode', 'vbaran', 'vbacell'];
const clone = x => JSON.parse(JSON.stringify(x));
const lc = s => String(s).toLowerCase();
const DEFAULT_SHEET = { sheets: [{ name: 'Sheet1' }] };

// ---------------------------------------------------------------- the lab's state
export function vbaStart(lab) { const v = lab.vba || {}; return { modules: (v.modules || [{ name: 'Module1', kind: 'std', code: '' }]).map(m => ({ kind: 'std', ...m })), sheet: null, view: v.start || 'excel', mod: (v.open || (v.modules && v.modules[0] && v.modules[0].name) || 'Module1'), bps: [], files: {}, imm: [], keys: { ...(v.keys || {}) }, buttons: (v.userButtons || []).slice() }; }
// every module the VBE shows: one per sheet, ThisWorkbook, then the learner's modules (and PERSONAL.XLSB's)
export function allModules(state, wb) { const out = []; const have = n => (state.modules || []).find(m => lc(m.name) === lc(n));
  wb.sheets.forEach((s, i) => { const nm = s.code || 'Sheet' + (i + 1); const m = have(nm); out.push({ name: nm, kind: 'sheet', sheet: s.name, code: m ? m.code : '' }); });
  const tw = have('ThisWorkbook'); out.push({ name: 'ThisWorkbook', kind: 'workbook', code: tw ? tw.code : '' });
  for (const m of state.modules || []) if (m.kind === 'std' || m.kind === 'form') out.push(m); return out; }
export function makeHost(lab, wb, files, folders) { const v = lab.vba || {}; const all = {}; for (const [p, f] of Object.entries(v.files || {})) all[p] = clone(f); for (const [p, f] of Object.entries(files || {})) { if (f === null) delete all[p]; else all[p] = clone(f); }
  return new ExcelHost({ wb, name: v.book || 'Book1.xlsm', path: v.path || 'C:\\Users\\you\\Documents', files: all, folders: folders || v.folders || [], now: v.now }); }

// ---------------------------------------------------------------- procedures: split a module into its parts, merge the model answer in
const PROC_RE = /^\s*(?:public\s+|private\s+|friend\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+([A-Za-z_]\w*)/i;
const END_RE = /^\s*end\s+(sub|function|property)\b/i;
export function splitProcs(code) { const lines = String(code || '').replace(/\r/g, '').split('\n'); const head = [], procs = []; let cur = null, lead = [];
  for (const l of lines) { if (!cur) { const m = PROC_RE.exec(l); if (m) { cur = { name: m[1], lines: [...lead, l] }; lead = []; continue; } if (/^\s*('|rem\b|$)/i.test(l)) { lead.push(l); continue; } head.push(...lead, l); lead = []; continue; }
    cur.lines.push(l); if (END_RE.test(l)) { procs.push(cur); cur = null; } }
  if (cur) procs.push(cur); return { head, procs, tail: lead }; }
const trimBlank = a => { const x = a.slice(); while (x.length && !x[0].trim()) x.shift(); while (x.length && !x[x.length - 1].trim()) x.pop(); return x; };
export function mergeCode(mine, model) { const L = splitProcs(mine), M = splitProcs(model); const head = trimBlank(L.head);
  for (const l of trimBlank(M.head)) if (!head.some(h => lc(h.trim()) === lc(l.trim()))) { if (/^\s*option\s+explicit/i.test(l)) head.unshift(l); else head.push(l); }
  const procs = L.procs.map(p => ({ ...p })); for (const p of M.procs) { const i = procs.findIndex(q => lc(q.name) === lc(p.name)); if (i >= 0) procs[i] = p; else procs.push(p); }
  const parts = []; if (head.length) parts.push(head.join('\n')); for (const p of procs) parts.push(trimBlank(p.lines).join('\n')); return parts.join('\n\n') + '\n'; }
export function vbaApply(task, st, wb, lab) { const s = task.msol || {}; if (task.kind === 'vbaran') { if (wb && lab) runOn(wb, lab, st, task); return; }
  for (const m of s.modules || []) { const kind = m.kind || (/^sheet\d+$|^thisworkbook$/i.test(m.name) ? (lc(m.name) === 'thisworkbook' ? 'workbook' : 'sheet') : 'std'); const have = st.modules.find(x => lc(x.name) === lc(m.name)); if (have) have.code = s.replace ? m.code : mergeCode(have.code, m.code); else st.modules.push({ ...m, kind }); }
  if (task.kind === 'vbacell' && wb) { const { si, g } = locate(wb, task.range); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) wb.setInput(si, r, c, cellFormula(task, r, c)); if (st) st.sheet = wb.toSpec(); }
  if (s.trust) st.trust = s.trust; if (s.keys) st.keys = { ...(st.keys || {}), ...s.keys }; if (s.buttons) { st.buttons = st.buttons || []; for (const b of s.buttons) if (!st.buttons.some(x => lc(x.macro) === lc(b.macro))) st.buttons.push({ ...b }); }
  if (s.removeModules) st.modules = st.modules.filter(m => !s.removeModules.some(n => lc(n) === lc(m.name))); }
export function vbaForTask(labVba, t) { if (!labVba || !t.before) return labVba; const mods = (labVba.modules || []).map(m => ({ ...m })); for (const m of t.before) { const h = mods.find(x => lc(x.name) === lc(m.name)); if (h) h.code = m.code; else mods.push({ kind: 'std', ...m }); } return { ...labVba, modules: mods }; }
export function modelState(lab, task) { const st = vbaStart(lab); if (task.before) for (const m of task.before) { const h = st.modules.find(x => lc(x.name) === lc(m.name)); if (h) h.code = m.code; else st.modules.push({ kind: 'std', ...m }); } vbaApply(task, st); return st; }

function growRows(wb, name, times) { const si = name ? Math.max(0, wb.sheets.findIndex(s => lc(s.name) === lc(name))) : 0; const u = wb.usedRange(wb.sheets[si]); if (!u || u.r2 < 1) return; const n = u.r2;
  for (let k = 1; k < times; k++) for (let r = 1; r <= n; r++) for (let c = 0; c <= u.c2; c++) { const t = wb.input(si, r, c); if (t !== '' && t !== null && t !== undefined) wb.setInput(si, r + k * n, c, t); } }
// 'vbaran': the lab's own workbook shows what the macro does (the learner ran it); the solution runs it there
function runOn(wb, lab, st, task) { const host = makeHost(lab, wb); if (task.select) { const { si, g } = locate(wb, task.select); host.sel = { book: host.this, si, g, ar: g.r1, ac: g.c1 }; host.this.activeSi = si; } const ms = modelState(lab, task); const prog = new Program(allModules(ms, wb).map(m => ({ ...m })), host); if (prog.firstError()) return null; return runMacro(prog, task.macro, { inputs: (task.inputs || []).slice(), answers: (task.answers || []).slice() }); }
function checkRan(task, ctx) { const wb = ctx.wb; const lab = (ctx.vba && ctx.vba.lab) || {}; if (!wb) return { ok: false, msg: 'Open the lab first.' }; const model = runCopy(lab, modelState(lab, task), task); if (model.compile || model.r.error) return { ok: false, msg: '(Sheetah’s model answer didn’t run here.)' };
  const d = diffBooks(wb, model.wb, { sheets: false, ...(task.compare || {}) }); if (d) return { ok: false, msg: `${task.notYet || `Not yet: run ${task.macro} (Developer › Macros, pick it, Run).`} ${d}` }; return { ok: true, msg: task.success || task.explain || 'Done: the macro ran.' }; }
// ---------------------------------------------------------------- running one copy
export function selectIn(host, a) { const { si, g } = locate(host.this.wb, a); host.sel = { book: host.this, si, g, ar: g.r1, ac: g.c1 }; host.this.activeSi = si; }
// the formula a task wants in a cell: {r} is that cell's row number, {c} its column letter
function cellFormula(task, r, c) { return task.formula.replace(/\{r\}/g, String(r + 1)).replace(/\{c\}/g, colLetter(c)); }
function colLetter(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
// vbacell: the learner typed formulas that use their own Function; each must give what the model Function gives in that formula
function checkCells(task, ctx) { const fail = msg => ({ ok: false, msg }); const lab = (ctx.vba && ctx.vba.lab) || {}; const st = (ctx.vba && ctx.vba.state) || vbaStart(lab);
  const spec = ctx.wb ? ctx.wb.toSpec() : (st.sheet || lab.sheet || DEFAULT_SHEET); const mine = new Workbook(clone(spec)); const host = makeHost(lab, mine); const prog = new Program(allModules(st, mine).map(m => ({ ...m })), host); const ce = prog.firstError(); if (ce) return fail(`The code doesn’t compile yet: ${ce.message}${ce.mod ? ' (' + ce.mod + ')' : ''}. Fix that first: a function can’t work in a cell while the project has compile errors.`);
  mine.udf = udfHook(prog, host); mine.invalidate();
  const want = new Workbook(clone(spec)); const mh = makeHost(lab, want); const mp = new Program(allModules(modelState(lab, task), want).map(m => ({ ...m })), mh); want.udf = udfHook(mp, mh);
  const { si, g } = locate(want, task.range); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) want.setInput(si, r, c, cellFormula(task, r, c)); want.invalidate();
  const fnRe = new RegExp('\\b' + task.fn + '\\s*\\(', 'i'); const sh = want.sheets[si].name;
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const a = addr(r, c); const f = String(mine.input(si, r, c) ?? ''); const ex = cellFormula(task, r, c);
    if (!f.startsWith('=')) return fail(`${sh}!${a} should hold a formula that uses your function, like ${ex}.${f ? ` It holds ${show(mine.value(si, r, c))}.` : ' It’s empty.'}`);
    if (!fnRe.test(f)) return fail(`${sh}!${a} should use your function ${task.fn}, like ${ex}. It holds ${f}.`);
    const va = mine.value(si, r, c), vb = want.value(si, r, c);
    if (va && va.code === '#NAME?') return fail(`${sh}!${a} shows #NAME?: Excel can’t find ${task.fn}. It must be a Public Function in a standard module (Module1), not in a sheet’s code window or Private, and spelt the same in the cell.`);
    if (!near(va, vb)) return fail(`${sh}!${a} should show ${show(vb && vb.code ? vb.code : vb)} (${ex}), but it shows ${show(va && va.code ? va.code : va)}.${va && va.code === '#VALUE!' ? ` #VALUE! means ${task.fn} stopped with an error for this cell, or tried to change cells: run it from the Immediate window (? ${task.fn}(…)) to see why.` : ''}`); }
  return { ok: true, msg: task.success || task.explain || 'Correct.' }; }
function locate(wb, a) { const m = /^'?(.+?)'?!(.+)$/.exec(a); const si = m ? wb.sheets.findIndex(s => lc(s.name) === lc(m[1])) : 0; return { si: Math.max(0, si), g: parseRange((m ? m[2] : a).replace(/\$/g, '')) }; }
export function runCopy(lab, state, task, variant = {}) {
  const spec = clone(variant.sheet || task.start || lab.sheet || DEFAULT_SHEET); const wb = new Workbook(spec);
  if (variant.repeat) growRows(wb, variant.repeatSheet, variant.repeat); // a month of data from a day's: the same rows, many times
  for (const [a, v] of Object.entries({ ...(task.edits || {}), ...(variant.edits || {}) })) { const { si, g } = locate(wb, a); wb.setInput(si, g.r1, g.c1, v === null ? '' : String(v)); }
  const host = makeHost(lab, wb, { ...(task.files || {}), ...(variant.files || {}) }, variant.folders || task.folders);
  const sel = variant.select || task.select; if (sel) { const { si, g } = locate(wb, sel); host.sel = { book: host.this, si, g, ar: g.r1, ac: g.c1 }; host.this.activeSi = si; }
  const prog = new Program(allModules(state, wb).map(m => ({ ...m })), host); const ce = prog.firstError(); if (ce) return { compile: ce, wb, host, prog }; wb.udf = udfHook(prog, host);
  const inputs = (variant.inputs || task.inputs || []).slice(), answers = (variant.answers || task.answers || []).slice(); let r;
  const script = (variant.form || task.form || []).slice(); const formLog = [];
  const formUi = req => { if (req.ui !== 'form') return undefined; formLog.push({ name: req.form.name, first: !!req.first, snap: req.form.snap(), step: (variant.form || task.form || []).length - script.length }); return script.length ? script.shift() : { close: true }; };
  const trig = variant.trigger || task.trigger;
  if (trig && trig.change) { const { si, g } = locate(wb, trig.change); for (let rr = g.r1; rr <= g.r2; rr++) for (let cc = g.c1; cc <= g.c2; cc++) wb.setInput(si, rr, cc, String(trig.value ?? '')); const M = [...prog.mods.values()].find(m => m.kind === 'sheet' && lc(m.sheet) === lc(wb.sheets[si].name)); const P = M && M.ast.procs.find(p => p.lname === 'worksheet_change'); if (!P) return { r: { done: true }, wb, host, prog, out: prog.out, ui: [], noHandler: M ? M.name : true };
    r = runProc(prog, { proc: P, mod: M }, [hostRange(host, si, g)], inputs, answers, formUi); }
  else if (trig && trig.open) { const M = prog.mods.get('thisworkbook'); const P = M && M.ast.procs.find(p => p.lname === 'workbook_open'); if (!P) return { r: { done: true }, wb, host, prog, out: prog.out, ui: [], noHandler: 'ThisWorkbook' }; r = runProc(prog, { proc: P, mod: M }, [], inputs, answers, formUi); }
  else if (task.showForm) { const F = prog.forms.get(task.showForm.toLowerCase()); if (!F) return { compile: Object.assign(new CompileError(`There’s no UserForm called ${task.showForm}.`, 0), { noForm: task.showForm }), wb, host, prog }; const log = []; const s = new Session(prog, { ui: req => { log.push(req); if (req.ui === 'inputbox') return inputs.length ? inputs.shift() : ''; if (req.ui === 'msgbox') return answers.length ? answers.shift() : undefined; return formUi(req); }, maxSteps: task.maxSteps || 2e6 }); s.start('x'); s.it = prog.showForm(F); r = s.runSync('run'); while (r.paused) r = s.runSync('run'); r = { ...r, ui: log }; }
  else { const runs = [...(variant.twice ? [task.macro, ...(task.chain || [])] : []), task.macro, ...(task.chain || [])]; // a chain: macros run one after another on the same workbook; twice: the whole lot again
    for (const [k, mac] of runs.entries()) { r = runMacro(prog, mac, { inputs: inputs.slice(), answers: answers.slice(), maxSteps: task.maxSteps || 2e6, args: variant.args || task.args, ui: formUi }); if (k && r.ui && r0ui) r.ui = [...r0ui, ...r.ui]; var r0ui = r.ui; if (r.error) { r.error.macro = mac; break; } } }
  return { r, wb, host, prog, out: prog.out || [], ui: r.ui || [], formLog, script: variant.form || task.form || [] };
}
function hostRange(host, si, g) { return new Rng(host, host.this, si, g); }
function runProc(prog, P, args, inputs, answers, formUi) { const log = []; inputs = inputs.slice(); answers = answers.slice(); const s = new Session(prog, { ui: req => { log.push(req); if (req.ui === 'inputbox') return inputs.length ? inputs.shift() : ''; if (req.ui === 'msgbox') return answers.length ? answers.shift() : undefined; return formUi ? formUi(req) : undefined; }, maxSteps: 2e6 }); s.start(P, args); let r = s.runSync('run'); while (r.paused) r = s.runSync('run'); return { ...r, ui: log }; }
// =GSTAmount(H2) in a cell: Excel calls the VBA Function. A typed parameter gets the cell's value; Variant, Object and Range get the Range itself.
// While it runs the function may not change cells (Excel's rule for functions called from a sheet): if it tries, or fails, the cell shows #VALUE!.
const SCALAR_T = new Set(['double', 'single', 'long', 'longlong', 'integer', 'byte', 'currency', 'string', 'boolean', 'date']);
export function udfHook(prog, host) { const fns = new Map(); for (const M of prog.mods.values()) if (M.kind === 'std' || !M.kind) for (const P of M.ast.procs) if (P.kind === 'function' && P.scope !== 'private') fns.set(P.lname, { proc: P, mod: M });
  const toVb = (a, pr) => { if (a === undefined) return MISSING; if (a && a.code && !a.range) return a; if (a && a.range && !a.arr && a.g) { const single = a.g.r1 === a.g.r2 && a.g.c1 === a.g.c2; if (pr && SCALAR_T.has(pr.type)) { const v = single ? a.single ?? a.vals[0][0] : a.vals[0][0]; return v === null || v === undefined ? (pr.type === 'string' ? '' : 0) : v; } return new Rng(host, host.this, a.si, a.g); }
    if (a && a.range) { const rows = a.vals; if (rows.length === 1 && rows[0].length === 1) return rows[0][0]; const out = new VArray([[1, rows.length], [1, rows[0].length]]); let k = 0; for (const r of rows) for (const x of r) out.d[k++] = x; return out; } return a === null ? EMPTY : a; };
  const back = v => { if (v && v.code) return Object.values(ERR).find(e => e.code === v.code) || ERR.VALUE; if (v instanceof VDate) return v.n; if (v === EMPTY || v === undefined || v === null) return 0; if (v instanceof VArray) { const two = v.b.length === 2; const nr = two ? v.b[0][1] - v.b[0][0] + 1 : 1, nc = two ? v.b[1][1] - v.b[1][0] + 1 : v.b[0][1] - v.b[0][0] + 1; const rows = []; for (let r = 0; r < nr; r++) { const row = []; for (let c = 0; c < nc; c++) { const x = two ? v.get1([v.b[0][0] + r, v.b[1][0] + c]) : v.get1([v.b[0][0] + c]); row.push(back(x)); } rows.push(row); } return mkArr(rows); } if (isObj(v)) return v.dflt ? back(v.dflt([])) : ERR.VALUE; return v; };
  return { has: lk => fns.has(lk), names: () => [...fns.values()].map(F => F.proc.name), call(lk, args, si) { const F = fns.get(lk); if (host.inUdf > 20) return ERR.VALUE; const ps = F.proc.params || [];
    if (args.length > ps.filter(p => !p.paramArray).length && !ps.some(p => p.paramArray)) return ERR.VALUE; if (args.length < ps.filter(p => !p.optional && !p.paramArray).length) return ERR.VALUE;
    host.inUdf = (host.inUdf || 0) + 1; try { const r = callFunction(prog, F, args.map((a, i) => toVb(a, ps[i]))); return r.error ? ERR.VALUE : back(r.value); } catch (e) { return ERR.VALUE; } finally { host.inUdf--; } } }; }
export function callFunction(prog, name, args) { const s = new Session(prog, { maxSteps: 1e6, ui: () => undefined }); s.start(name, args); let r = s.runSync('run'); while (r.paused) r = s.runSync('run'); return r; }

// ---------------------------------------------------------------- messages
// a number in a date or time cell reads as the date, in the model’s format
const showAs = (v, fmt) => (typeof v === 'number' && fmt && /[dyh]/i.test(fmt.replace(/"[^"]*"/g, ''))) ? `“${formatValue(v, fmt)}”` : show(v);
const show = v => { if (v === null || v === undefined || v === '' || v === EMPTY) return 'empty'; if (typeof v === 'string') return `“${v}”`; if (typeof v === 'number') return numStr(Math.round(v * 1e9) / 1e9); if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (isErr(v)) return v.code; if (v && v.n !== undefined) return '#' + v.n + '#'; return String(v); };
const where = e => e.line ? ` on line ${e.line}${e.mod ? ' of ' + e.mod : ''}` : '';
export function errText(e, name) { if (e instanceof CompileError) return `The code doesn’t compile: “${e.message}”${where(e)}${e.word ? ` (at ${e.word})` : ''}. Debug › Compile VBAProject in the lab takes you to it.`;
  if (e && e.runaway) return `${name ? name + ' ' : 'The macro '}ran for too long and was stopped${where(e)}: is there a loop that never ends? Check that the loop changes what its condition tests.`;
  if (e && e.num === 28 && /^worksheet_change$|^workbook_sheetchange$/i.test(e.proc || '')) return `${e.proc} stops with run-time error '28': Out of stack space${where(e)}. It writes to the sheet, and that change fires ${e.proc} again, and again… React only to the column you watch (If Intersect(Target, Me.Range("F:F")) Is Nothing Then Exit Sub), and when you write back into that column, put Application.EnableEvents = False before the write and Application.EnableEvents = True after it.`;
  return `Running ${name || 'the macro'} stops with run-time error '${e.num < 0 ? `${e.num} (${(e.num >>> 0).toString(16).toUpperCase()})` : e.num}': ${e.desc || e.message}${where(e)}. Run it in the lab and press Debug: the yellow line is where it fails.`; }
const near = (a, b) => { if (a === b) return true; const blank = x => x === null || x === undefined || x === ''; if (blank(a) && blank(b)) return true; if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b)); if (isErr(a) && isErr(b)) return a.code === b.code; return false; };
const STYLE_WORDS = { hasFill: 'fill', bold: 'bold', italic: 'italic', fill: 'fill colour', color: 'font colour', fmt: 'number format', align: 'alignment', wrap: 'wrap text', border: 'borders', size: 'font size', underline: 'underline' };
const COLOUR_NAMES = { '#DDEBF7': 'light blue', '#FFF2CC': 'light yellow', '#E2EFDA': 'light green', '#FCE4D6': 'light orange', '#EDE7F6': 'lavender', '#F2F2F2': 'light grey', '#23263A': 'dark navy', '#FFFF00': 'yellow', '#FF0000': 'red', '#00B050': 'green', '#0000FF': 'blue', '#FFFFFF': 'white', '#000000': 'black', '#C6EFCE': 'light green', '#FFC7CE': 'light red', '#FFEB9C': 'light yellow' };
const colourWord = h => COLOUR_NAMES[h] ? `${COLOUR_NAMES[h]} (${h})` : h;
// the cells to compare: everything with a value, and, when formatting is compared, everything with formatting too
function extent(wb, s, styles) { let g = wb.usedRange(s); if (styles && styles.length) for (const [k, cell] of s.cells) { if (!cell.s || !Object.keys(cell.s).length) continue; const r = Math.floor(k / 16384), c = k % 16384; g = g ? { r1: Math.min(g.r1, r), c1: Math.min(g.c1, c), r2: Math.max(g.r2, r), c2: Math.max(g.c2, c) } : { r1: r, c1: c, r2: r, c2: c }; } return g; }
function normStyle(k, v) { if (v === undefined || v === null || v === false || v === '') return k === 'fmt' ? 'General' : null; if (k === 'fill' || k === 'color') return String(v).toUpperCase(); return v; }
// the first difference between what the learner's macro left and what the model answer left
export function diffBooks(got, want, cmp = {}) {
  const gs = got.sheets.map(s => s.name), ws = want.sheets.map(s => s.name);
  if (cmp.sheets !== false) { for (const n of ws) if (!gs.some(x => lc(x) === lc(n))) return `After the macro runs there should be a sheet called “${n}”, and there isn’t one.`; if (cmp.sheets === 'exact' || cmp.sheetOrder) { for (const n of gs) if (!ws.some(x => lc(x) === lc(n))) return `After the macro runs there’s a sheet called “${n}” that the model answer doesn’t have.`; if (cmp.sheetOrder && gs.join('|').toLowerCase() !== ws.join('|').toLowerCase()) return `The sheets should end up in this order: ${ws.join(', ')}. Yours are ${gs.join(', ')}.`; } }
  const list = Array.isArray(cmp.only) ? cmp.only : ws; const styles = cmp.styles || [];
  for (const n of list) { const wi = want.sheets.findIndex(s => lc(s.name) === lc(n)), gi = got.sheets.findIndex(s => lc(s.name) === lc(n)); if (wi < 0 || gi < 0) continue; const W = want.sheets[wi], G = got.sheets[gi];
    const rs = cmp.range ? (typeof cmp.range === 'object' ? (cmp.range[n] || cmp.range['*']) : cmp.range) : null; let g = rs ? parseRange(rs) : null; if (!g) { const a = extent(want, W, styles), b = extent(got, G, styles); if (!a && !b) continue; const u = a && b ? { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) } : (a || b); g = { r1: 0, c1: 0, r2: u.r2, c2: u.c2 }; }
    const pre = ws.length > 1 || n !== ws[0] ? `${n}!` : '';
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const a = got.value(gi, r, c), b = want.value(wi, r, c);
      if (!near(a, b) && !(cmp.text && got.display(gi, r, c) === want.display(wi, r, c))) { const wf = want.style(wi, r, c).fmt; if ((b === null || b === '') && a !== null && a !== '') return `${pre}${addr(r, c)} should be empty after the macro runs, but it holds ${showAs(a, wf)}.`; return `${pre}${addr(r, c)} should be ${showAs(b, wf)} after the macro runs, but it is ${showAs(a, wf)}.`; }
      if (cmp.formulas) { const fa = got.formula(gi, r, c), fb = want.formula(wi, r, c); if (fb && !fa) return `${pre}${addr(r, c)} should hold a formula (like ${fb}), so it updates later. Your macro typed the value ${show(a)} instead: set .Formula, not .Value.`; if (fa && !fb && cmp.formulas === 'exact') return `${pre}${addr(r, c)} should hold the value, not the formula ${fa}: copy it with .Value = .Value.`; }
      for (const k of styles) { const sa = k === 'hasFill' ? !!got.style(gi, r, c).fill || null : normStyle(k, got.style(gi, r, c)[k]), sb = k === 'hasFill' ? !!want.style(wi, r, c).fill || null : normStyle(k, want.style(wi, r, c)[k]); if (String(sa) !== String(sb)) return `${pre}${addr(r, c)}: the ${STYLE_WORDS[k] || k} should be ${sb === null ? 'off' : k === 'fill' || k === 'color' ? colourWord(sb) : k === 'hasFill' ? 'on' : show(sb)} after the macro runs, but it’s ${sa === null ? 'off' : k === 'fill' || k === 'color' ? colourWord(sa) : k === 'hasFill' ? 'on' : show(sa)}.`; } }
    if (cmp.hidden) { const h = s => [...new Set(s.hidden || [])].sort((x, y) => x - y).join(','); if (h(G) !== h(W)) return `On ${n}, the macro should leave ${(W.hidden || []).length ? 'rows ' + (W.hidden || []).slice(0, 6).map(x => x + 1).join(', ') + ((W.hidden || []).length > 6 ? '…' : '') + ' hidden' : 'no rows hidden'}; yours leaves ${(G.hidden || []).length ? (G.hidden || []).length + ' hidden' : 'none hidden'}.`; }
    if (cmp.widths) for (let c = g.c1; c <= g.c2; c++) { const a = (G.widths || {})[c], b = (W.widths || {})[c]; if (b && Math.abs((a || 64) - b) > 6) return `On ${n}, column ${String.fromCharCode(65 + c)} should be widened to fit (AutoFit).`; } }
  return null; }
function fileWords(f) { return { xlsx: 'workbook', xlsm: 'macro workbook', csv: 'CSV file', pdf: 'PDF', txt: 'text file', folder: 'folder' }[f.kind] || 'file'; }
// emails the macro made (Outlook's Outbox, or windows it displayed): how many, to whom, subject, attachments, and Send vs Display
const MAIL_WORDS = { to: 'To', cc: 'Cc', subject: 'Subject', attachments: 'attachments' };
export function diffMail(A, B, cmp = {}) { const n = x => `${x} email${x === 1 ? '' : 's'}`;
  if (A.length !== B.length) return B.length ? `The macro should make ${n(B.length)}; yours makes ${A.length || 'none'}.${A.length > B.length ? ' Check which rows should get one.' : ''}` : `The macro shouldn’t make any email here; yours makes ${A.length}.`;
  const norm = (f, v) => f === 'attachments' ? (v || []).map(x => x.toLowerCase()).join(', ') : String(v || '').replace(/\s*;\s*/g, '; ').trim();
  for (let k = 0; k < B.length; k++) { const a = A[k], b = B[k]; const lab = B.length > 1 ? `Email ${k + 1} (to ${b.to || b.cc})` : 'The email';
    for (const f of ['to', 'cc', 'subject', 'attachments']) if (norm(f, a[f]) !== norm(f, b[f])) return f === 'attachments' ? `${lab} should have ${b.attachments.length ? 'the attachment ' + b.attachments.join(', ') : 'no attachment'}; yours has ${a.attachments.length ? a.attachments.join(', ') : 'none'}.` : `${lab} should have ${MAIL_WORDS[f]} “${b[f]}”; yours has “${a[f] || ''}”.`;
    if (cmp.mail === 'body' && flat(a.body) !== flat(b.body)) return `${lab}’s body should read “${flat(b.body)}”. Yours reads “${flat(a.body)}”.`;
    if (cmp.mail === 'greet') { const first = t => (String(t || '').split(/\r?\n/).find(x => x.trim()) || '').trim(); if (first(a.body) !== first(b.body)) return `${lab} should start “${first(b.body)}”; yours starts “${first(a.body)}”.`; }
    if ((a.how === 'send') !== (b.how === 'send')) return b.how === 'send' ? `${lab} should be sent with .Send; yours only ${a.how === 'display' ? 'displays it (.Display)' : 'saves a draft'}.` : `${lab} should be shown with .Display so it can be checked first; yours sends it straight away with .Send.`; }
  return null; }
const PAGE_WORDS = { orientation: v => `Orientation = ${v === 2 ? 'xlLandscape' : 'xlPortrait'}`, fittopageswide: v => `FitToPagesWide = ${v}`, fittopagestall: v => `FitToPagesTall = ${v === false ? 'False' : v}`, zoom: v => `Zoom = ${v === false ? 'False' : v}`, printtitlerows: v => `PrintTitleRows = "${v}"`, centerfooter: v => `CenterFooter = "${v}"`, rightfooter: v => `RightFooter = "${v}"`, leftfooter: v => `LeftFooter = "${v}"`, centerheader: v => `CenterHeader = "${v}"`, printarea: v => `PrintArea = "${v}"` };
export function diffFiles(gh, wh, cmp = {}) { const made = h => [...h.fs.values()].filter(f => f.made); const G = made(gh), W = made(wh); const key = f => lc(f.path);
  for (const f of W) { const m = G.find(x => key(x) === key(f)); if (!m) { if (f.kind === 'folder') return `The macro should make the folder ${f.path} (MkDir) when it isn’t there, and it doesn’t.`; const near2 = G.find(x => x.kind === f.kind); return `The macro should save a ${fileWords(f)} as ${f.path}${near2 ? `. Yours saved ${near2.path}: check the folder and file name` : ', and nothing was saved there'}.`; }
    if (f.kind === 'csv' || f.kind === 'txt') { const t = s => String(s || '').replace(/\r\n/g, '\n').replace(/\n+$/, ''); if (t(m.text) !== t(f.text)) { const a = t(m.text).split('\n'), b = t(f.text).split('\n'); const i = b.findIndex((x, k) => x !== a[k]); return `${f.path.split('\\').pop()} doesn’t hold what it should: line ${i + 1} should be “${b[i] ?? ''}”, but it is “${a[i] ?? ''}”.`; } }
    else if (f.kind === 'pdf') { const pa = m.pages || [], pb = f.pages || []; if (pa.length !== pb.length) return `${f.path.split('\\').pop()} should have ${pb.length} page${pb.length === 1 ? '' : 's'}; yours has ${pa.length}.`; for (let i = 0; i < pb.length; i++) { const sb = pb[i].setup || {}, sa = pa[i].setup || {}; for (const k of Object.keys(sb)) if (JSON.stringify(sa[k] ?? null) !== JSON.stringify(sb[k]) && String(sa[k] ?? '').replace(/\$/g, '').toLowerCase() !== String(sb[k]).replace(/\$/g, '').toLowerCase()) return `${f.path.split('\\').pop()} should be set up with ${(PAGE_WORDS[k] || (v => `${k} = ${v}`))(sb[k])} (on the sheet’s PageSetup, before exporting).`; }
      for (let i = 0; i < pb.length; i++) if (JSON.stringify(pa[i].rows) !== JSON.stringify(pb[i].rows) || pa[i].sheet !== pb[i].sheet) return `Page ${i + 1} of ${f.path.split('\\').pop()} should show ${pb[i].sheet}, as it looks after the macro has run.`; }
    else if (f.spec) { const d = diffBooks(new Workbook(clone(m.spec)), new Workbook(clone(f.spec)), { ...(cmp.fileCompare || {}), sheets: 'exact' }); if (d) return `In the saved ${f.path.split('\\').pop()}: ${d}`; } }
  if (cmp.files === 'exact') for (const f of G) if (!W.some(x => key(x) === key(f))) return `Your macro also saved ${f.path}, which it shouldn’t.`;
  if (cmp.kill) { const left = h => [...h.fs.values()].filter(f => !f.made).map(f => lc(f.path)).sort().join('|'); if (left(gh) !== left(wh)) return 'The macro should delete exactly the files the model answer deletes, and leave the others.'; }
  return null; }
const flat = s => String(s ?? '').replace(/\r\n|\r|\n/g, ' ⏎ ').replace(/\s+/g, ' ').trim();
// Debug.Print puts a space before positive numbers: lines are compared trimmed
export function diffOut(a, b) { const A = a.map(x => x.trim()), B = b.map(x => x.trim()); for (let i = 0; i < Math.max(A.length, B.length); i++) if (A[i] !== B[i]) { if (B[i] === undefined) return `The Immediate window should show ${B.length} line${B.length === 1 ? '' : 's'}; yours shows more: “${A[i]}”.`; if (A[i] === undefined) return `The Immediate window should show “${B[i]}” as line ${i + 1}, but your macro stops printing after ${A.length} line${A.length === 1 ? '' : 's'}.`; return `Line ${i + 1} in the Immediate window should be “${B[i]}”, but it is “${A[i]}”.`; } return null; }
function diffUi(a, b, cmp) { const boxes = x => x.filter(u => u.ui === 'msgbox'); const A = boxes(a), B = boxes(b); if (A.length !== B.length) return B.length ? (A.length ? `The macro should show ${B.length} message box${B.length === 1 ? '' : 'es'}; yours shows ${A.length}.` : `The macro should finish with a message: “${flat(B[0].prompt)}”. Yours shows none: add MsgBox.`) : `The macro shouldn’t show a message box here; yours shows “${flat(A[0].prompt)}”.`;
  for (let i = 0; i < B.length; i++) { const pa = flat(A[i].prompt), pb = flat(B[i].prompt); if (cmp.ui === 'loose' ? !(pa.toLowerCase().includes(pb.toLowerCase()) || pb.toLowerCase().includes(pa.toLowerCase())) : pa !== pb) return `The message should read “${pb}”. Yours reads “${pa}”.`; if (cmp.uiButtons && (A[i].buttons || 0) !== (B[i].buttons || 0)) return `The message “${pb}” should use the buttons and icon the task asks for (the second argument of MsgBox).`; if (cmp.uiTitle && (A[i].title || '') !== (B[i].title || '')) return `The message box’s title should be “${B[i].title || 'Microsoft Excel'}”: it’s the third argument of MsgBox.`; }
  return null; }

const actWords = a => { if (!a) return 'when the form opens'; if (a.click) return `after clicking ${a.click}`; if (a.close) return 'after closing the form'; if (a.dblclick) return `after double-clicking ${a.dblclick}`; if (a.set) return 'after ' + Object.entries(a.set).map(([k, v]) => typeof v === 'boolean' ? `${v ? 'ticking' : 'clearing'} ${k}` : `entering “${v}” in ${k}`).join(' and '); return 'next'; };
const listWords = l => l.length ? l.slice(0, 6).join(', ') + (l.length > 6 ? '…' : '') : '(empty)';
// only: true compares every control; 'open' only which forms open and close; a list of names compares just those controls
export function diffForms(A, B, script, only = true) {
  if (!A.length && !B.length) return null;
  if (!A.length) return `The macro should open the form ${B[0].name} (${B[0].name}.Show).`; if (!B.length) return `The macro shouldn’t open ${A[0].name} here.`;
  for (let i = 0; i < Math.min(A.length, B.length); i++) { const a = A[i], b = B[i]; const when = actWords(i ? script[i - 1] : null); const When = when[0].toUpperCase() + when.slice(1);
    if (a.name.toLowerCase() !== b.name.toLowerCase()) return `${When}, the form showing should be ${b.name}; it’s ${a.name}.`;
    if (only === 'open') continue;
    if ((only === true || (Array.isArray(only) && only.includes('Caption'))) && a.snap.caption !== b.snap.caption) return `${When}, the form’s title bar should read “${b.snap.caption}”; it reads “${a.snap.caption}” (Me.Caption).`;
    for (const [n, cb] of Object.entries(b.snap.controls)) { const ca = a.snap.controls[n]; if (!ca || (Array.isArray(only) && !only.some(x => x.toLowerCase() === n.toLowerCase()))) continue;
      if (cb.list && JSON.stringify(cb.list) !== JSON.stringify(ca.list)) return `${When}, ${n}’s list should be ${listWords(cb.list)}; it’s ${listWords(ca.list)}.`;
      if (cb.type === 'Label' || cb.type === 'CommandButton' || cb.type === 'Frame') { if (cb.caption !== ca.caption) return `${When}, ${n} should say “${cb.caption}”; it says “${ca.caption}”.`; }
      else if (JSON.stringify(cb.value) !== JSON.stringify(ca.value)) return `${When}, ${n} should ${cb.value === '' || cb.value === null ? 'be empty' : cb.value === true ? 'be ticked' : cb.value === false ? 'not be ticked' : `hold “${cb.value}”`}; it ${ca.value === '' || ca.value === null ? 'is empty' : ca.value === true ? 'is ticked' : ca.value === false ? 'isn’t ticked' : `holds “${ca.value}”`}.`;
      if (cb.enabled !== ca.enabled) return `${When}, ${n} should be ${cb.enabled ? 'enabled' : 'disabled (Enabled = False)'}; it’s ${ca.enabled ? 'enabled' : 'disabled'}.`;
      if (cb.visible !== ca.visible) return `${When}, ${n} should be ${cb.visible ? 'visible' : 'hidden (Visible = False)'}.`; } }
  if (A.length === B.length) return null;
  if (A.length > B.length) return `${actWords(script[B.length - 1]).replace(/^./, c => c.toUpperCase())}, the form should close (Unload Me), but it stays open.`;
  return `${actWords(script[A.length - 1]).replace(/^./, c => c.toUpperCase())}, the form closes, but it should stay open.`; }

// ---------------------------------------------------------------- code-reading checks
function codeOf(state, wb, modName) { const mods = allModules(state, wb || new Workbook(DEFAULT_SHEET)); return modName ? mods.filter(m => lc(m.name) === lc(modName) || (modName === 'std' && m.kind === 'std') || (modName === 'personal' && m.project === 'PERSONAL.XLSB')) : mods; }
function procText(state, wb, name) { for (const m of allModules(state, wb)) for (const p of splitProcs(m.code).procs) if (lc(p.name) === lc(name)) return { text: p.lines.join('\n'), mod: m }; return null; }
const stripComments = code => code.split('\n').map(l => { let q = false; for (let i = 0; i < l.length; i++) { if (l[i] === '"') q = !q; else if (l[i] === "'" && !q) return l.slice(0, i); } return /^\s*rem\b/i.test(l) ? '' : l; }).join('\n');
function structure(task, state, wb) { for (const n of task.need || []) { const src = n.proc ? (procText(state, wb, n.proc) || { text: '' }).text : codeOf(state, wb, n.mod).map(m => m.code).join('\n'); const body = n.comments ? src : stripComments(src); if (!new RegExp(n.re, n.flags || 'im').test(body)) return n.msg; }
  for (const n of task.not || []) { const src = n.proc ? (procText(state, wb, n.proc) || { text: '' }).text : codeOf(state, wb, n.mod).map(m => m.code).join('\n'); const body = n.comments ? src : stripComments(src); if (new RegExp(n.re, n.flags || 'im').test(body)) return n.msg; }
  for (const m of task.modules || []) if (!(state.modules || []).some(x => lc(x.name) === lc(m))) return `There should be a module called ${m}. Insert › Module adds one; rename it in the Properties window, (Name).`;
  return null; }
function findMacro(task, state, wb, prog) { const name = task.macro; const P = prog.findProc(name); const any = procText(state, wb, name);
  if (!P) { if (any && any.mod.kind !== 'std') return `${name} is in ${any.mod.name}’s code window. Macros you run from Alt+F8 belong in a module: Insert › Module, then move the Sub there.`; const all = allModules(state, wb).flatMap(m => splitProcs(m.code).procs.map(p => p.name)); const close = all.find(x => lc(x).replace(/_/g, '') === lc(name).replace(/_/g, '') || lev(lc(x), lc(name)) <= 2);
    return close ? `There’s a macro called ${close}, but the task needs it called exactly ${name}. Rename the Sub (the name after Sub).` : `There’s no macro called ${name} yet. ${task.recorded ? 'Developer › Record Macro, type the name ' + name + ', do the steps, then Stop Recording.' : 'Type Sub ' + name + '() in a module and press Enter.'}`; }
  if (P.proc.kind !== 'sub' && task.kind === 'vba') return `${name} is a Function. A macro you run is a Sub: change Function to Sub (and End Function to End Sub).`;
  if (P.proc.params.length && !task.args && task.kind === 'vba') return `${name} takes arguments, so it can’t run on its own from the Macros list. Take out what’s between the brackets: Sub ${name}().`;
  if (task.where === 'personal' && P.mod.project !== 'PERSONAL.XLSB') return `${name} works, but it’s stored in this workbook. Record it with Store macro in: Personal Macro Workbook, so it’s there in every workbook.`;
  if (task.where === 'std' && P.mod.project === 'PERSONAL.XLSB') return `${name} should be stored in this workbook (a module in VBAProject (Book1.xlsm)), not in PERSONAL.XLSB.`;
  if (task.inModule && lc(P.mod.name) !== lc(task.inModule)) return `${name} should be in ${task.inModule}; it’s in ${P.mod.name}.`;
  return null; }
function lev(a, b) { const d = Array.from({ length: a.length + 1 }, (_, i) => [i]); for (let j = 1; j <= b.length; j++) d[0][j] = j; for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); return d[a.length][b.length]; }

// ---------------------------------------------------------------- the check
export function checkVBA(task, ctx) { const fail = msg => ({ ok: false, msg }); const pass = () => ({ ok: true, msg: task.success || task.explain || 'Correct.' }); if (task.kind === 'vbaran') return checkRan(task, ctx); if (task.kind === 'vbacell') return checkCells(task, ctx);
  const V = ctx.vba; if (!V) return fail('Open the VBA lab first.'); const lab = V.lab || {}; const state = V.state; const wb0 = new Workbook(clone(task.start || lab.sheet || DEFAULT_SHEET));
  // compile first: nothing runs while any module has a compile error
  const p0 = new Program(allModules(state, wb0).map(m => ({ ...m })), makeHost(lab, wb0)); const ce = p0.firstError();
  if (task.kind === 'vbacode') { const s = structure(task, state, wb0); if (s) return fail(s); if (task.compiles !== false && ce) return fail(errText(ce)); if (task.macro) { const fm = findMacro(task, state, wb0, p0); if (fm) return fail(fm); } return extras(task, state, lab) || pass(); }
  if (ce) return fail(errText(ce));
  if (task.kind === 'vba' && !task.trigger && !task.showForm) { const fm = findMacro(task, state, wb0, p0); if (fm) return fail(fm); }
  if (task.kind === 'vbafn') return checkFn(task, state, lab, p0);
  const ms = modelState(lab, task); const cmp = { values: true, ...(task.compare || {}) };
  const variants = [{}, ...(task.alsoWith || [])];
  for (const [i, v] of variants.entries()) { const mine = runCopy(lab, state, task, v), model = runCopy(lab, ms, task, v);
    if (model.compile || model.r.error) return fail('(Sheetah’s model answer didn’t run here: ' + errText(model.compile || model.r.error, (model.r && model.r.error && model.r.error.macro) || task.macro) + ')');
    const lead = i ? `Right for the lab’s data, but not ${v.label || 'on another copy of the data'}: ` : '';
    if (mine.compile) return fail(errText(mine.compile));
    if (mine.noHandler) return fail(typeof mine.noHandler === 'string' ? `${mine.noHandler}’s code window should have a ${task.trigger.open ? 'Workbook_Open' : 'Worksheet_Change'} event: pick ${task.trigger.open ? 'Workbook' : 'Worksheet'} in the left drop-down and ${task.trigger.open ? 'Open' : 'Change'} in the right one.` : 'There’s no event procedure to run.');
    if (mine.r.error) return fail(lead + errText(mine.r.error, mine.r.error.macro || task.macro));
    let d = null;
    if (cmp.forms !== false) d = diffForms(mine.formLog, model.formLog, model.script, cmp.forms ?? true);
    if (!d && cmp.ui !== false) d = diffUi(mine.ui, model.ui, cmp); // a missing message is the clearest thing to say first
    if (!d && cmp.inputs !== false) { const ia = mine.ui.filter(u => u.ui === 'inputbox').length, ib = model.ui.filter(u => u.ui === 'inputbox').length; if (ia !== ib) d = ib > ia ? `The macro should ask ${ib === 1 ? 'a question with InputBox' : ib + ' questions with InputBox'}${ia ? `; yours asks ${ia}` : ''}.` : `The macro asks ${ia} question${ia === 1 ? '' : 's'}; it should ask ${ib || 'none'}.`; }
    if (!d && cmp.values !== false) d = diffBooks(mine.wb, model.wb, cmp);
    if (!d && cmp.out) d = diffOut(mine.out, model.out);
    if (!d && cmp.files) d = diffFiles(mine.host, model.host, cmp);
    if (!d && cmp.active) { const an = mine.host.sel.book.wb.sheets[mine.host.sel.si].name, bn = model.host.sel.book.wb.sheets[model.host.sel.si].name; if (lc(an) !== lc(bn)) d = `When the macro finishes, ${bn} should be the active sheet (it’s ${an}).`; else if (cmp.active === 'cell') { const ga = rangeStr(mine.host.sel.g), gb = rangeStr(model.host.sel.g); if (ga !== gb) d = `When the macro finishes, ${gb} should be selected; it’s ${ga}.`; } }
    if (!d && cmp.books) { const a = mine.host.books.length, b = model.host.books.length; if (a !== b) d = b < a ? `The macro leaves ${a - b} extra workbook${a - b === 1 ? '' : 's'} open: close what you open (wb.Close SaveChanges:=False).` : 'The macro should leave the new workbook open.'; }
    if (!d && cmp.mail) d = diffMail(mine.host.outbox, model.host.outbox, cmp);
    if (!d && cmp.app) { const ea = mine.host.events !== false, eb = model.host.events !== false; if (ea !== eb) d = eb ? 'When the macro finishes, Application.EnableEvents is still False, so no event macro (Worksheet_Change, Workbook_Open…) will run again until something turns it back on. Set it back to True before the macro ends, on every way out, including after an error.' : 'The macro should leave Application.EnableEvents off here.'; else if (mine.host.calc !== model.host.calc) d = model.host.calc === -4105 ? 'When the macro finishes, calculation is still manual, so formulas in the workbook stop updating. Set Application.Calculation = xlCalculationAutomatic before the macro ends, on every way out, including after an error.' : 'The macro should leave calculation set the way the task asks.'; else if (model.host.status === false && mine.host.status !== false) d = `When the macro finishes, the status bar still says “${mine.host.status}”. Give it back to Excel at the end: Application.StatusBar = False.`; }
    if (!d && cmp.speed) { const t = mine.prog.clock || 0; if (t > cmp.speed) d = `It works, but it takes ${t.toFixed(1)} seconds on this data. Make it faster: ${cmp.speedHint || 'turn off ScreenUpdating while it runs, and read and write the range in one go.'}`; }
    if (d) return fail(lead + d + (i ? (v.why ? ' ' + v.why : '') : (task.whyMain ? ' ' + task.whyMain : ''))); }
  const s = structure(task, state, wb0); if (s) return fail(s);
  return extras(task, state, lab) || pass(); }
function extras(task, st, lab = {}) { if (task.trust) { const tr = st.trust || (lab.vba || {}).trust || 'ok'; if (tr !== 'ok') return { ok: false, msg: tr === 'blocked' ? 'Macros in this workbook are still blocked (the red SECURITY RISK bar). Unblock the file in its Properties, then close and reopen it.' : tr === 'reopen' ? 'You’ve unblocked the file. Now close and reopen it, then enable the content.' : 'Macros are still disabled: click Enable Content on the yellow bar.' }; } if (task.key) { const k = Object.entries(st.keys || {}).find(([, m]) => lc(m) === lc(task.macro)); if (!k) return { ok: false, msg: `${task.macro} works. Now give it the shortcut Ctrl+Shift+${task.key}: Developer › Macros, pick it, Options…, or set it when you record.` }; if (k[0].toUpperCase() !== task.key.toUpperCase()) return { ok: false, msg: `${task.macro}’s shortcut is Ctrl+Shift+${k[0]}; the task asks for Ctrl+Shift+${task.key}.` }; }
  if (task.button) { const b = (st.buttons || []).find(x => lc(x.macro) === lc(task.macro)); if (!b) return { ok: false, msg: `${task.macro} works. Now put a button on the sheet for it: Developer › Insert › Button, then pick ${task.macro} in Assign Macro.` }; if (typeof task.button === 'string' && lc(b.label.trim()) !== lc(task.button)) return { ok: false, msg: `The button should read “${task.button}”; it reads “${b.label}”. Right-click it › Edit Text.` }; }
  return null; }
function checkFn(task, state, lab, prog) { const fail = msg => ({ ok: false, msg }); const name = task.fn; const P = prog.findProc(name); const any = procText(state, new Workbook(clone(lab.sheet || DEFAULT_SHEET)), name);
  if (!P) return fail(any ? `${name} is in ${any.mod.name}. A function you call from a cell or another macro belongs in a module (Insert › Module).` : `There’s no function called ${name} yet. Type Function ${name}(…) in a module.`);
  if (P.proc.kind !== 'function') return fail(`${name} is a Sub. To give back an answer it must be a Function: Function ${name}(…) … ${name} = answer … End Function.`);
  const wantN = (task.cases[0] || []).length; if (P.proc.params.filter(p => !p.optional && !p.paramArray).length > wantN) return fail(`${name} should take ${wantN} argument${wantN === 1 ? '' : 's'}${task.params ? ': ' + task.params : ''}.`);
  const mp = new Program(allModules(modelState(lab, task), new Workbook(clone(lab.sheet || DEFAULT_SHEET))).map(m => ({ ...m })), makeHost(lab, new Workbook(clone(lab.sheet || DEFAULT_SHEET))));
  for (const args of task.cases) { const call = `${name}(${args.map(a => typeof a === 'string' ? `"${a}"` : String(a)).join(', ')})`;
    const p2 = new Program(allModules(state, new Workbook(clone(lab.sheet || DEFAULT_SHEET))).map(m => ({ ...m })), makeHost(lab, new Workbook(clone(lab.sheet || DEFAULT_SHEET))));
    const a = callFunction(p2, name, args), b = callFunction(mp, name, args); if (b.error) return fail('(Sheetah’s model answer failed for ' + call + ': ' + errText(b.error) + ')');
    if (a.error) return fail(`${call} stops with run-time error '${a.error.num}': ${a.error.desc}${where(a.error)}.`);
    const va = a.value, vb = b.value; const sa = isObj(va) ? va.tn : va, sb = isObj(vb) ? vb.tn : vb;
    if (!(near(sa, sb) || (sa && sa.n !== undefined && sb && sb.n !== undefined && sa.n === sb.n) || (typeof sb === 'number' && typeof sa === 'number' && task.tol && Math.abs(sa - sb) <= task.tol))) { if (sa === EMPTY || sa === undefined) return fail(`${call} gives back nothing. Inside the function, set ${name} = the answer before End Function.`); return fail(`${call} should give back ${show(sb)}, but gives ${show(sa)}.`); } }
  const s = structure(task, state, new Workbook(clone(lab.sheet || DEFAULT_SHEET))); if (s) return fail(s);
  return { ok: true, msg: task.success || task.explain || 'Correct.' }; }

export function describeVBA(task) { const c = task.compare || {}; if (task.kind === 'vbacell') return [`every cell of ${task.range} holds a formula using ${task.fn}`, `each shows what ${task.formula.replace(/\{r\}/g, 'n')} gives with the model ${task.fn}`]; if (task.kind === 'vbaran') return [`the lab’s workbook shows what ${task.macro} does when it runs`]; if (task.kind === 'vbacode') return [(task.need || []).map(n => n.what || '').filter(Boolean).join('; ') || 'the code itself'];
  if (task.kind === 'vbafn') return [`a Function called ${task.fn} in a module`, `that gives the model answer’s result for ${task.cases.length} different input${task.cases.length === 1 ? '' : 's'}`];
  const parts = [`${task.trigger ? (task.trigger.open ? 'a Workbook_Open event' : 'a Worksheet_Change event') : `a macro called ${task.macro} in a module`}`, `that, run on a fresh copy of the workbook, leaves the same ${['values', ...(c.styles || []).map(k => STYLE_WORDS[k] || k)].join(', ')}${c.files ? ', saved files' : ''}${c.out ? ', Immediate-window output' : ''} as the model answer`];
  if ((task.alsoWith || []).length) parts.push(`and still does on ${task.alsoWith.map(v => v.label || 'another copy').join(' and ')}`); if (task.key) parts.push(`with the shortcut Ctrl+Shift+${task.key}`); if (task.button) parts.push('with a button on the sheet that runs it'); return parts; }
export function vbaSolvedText(lab, tasks) { const st = vbaStart(lab); for (const t of tasks) if (VBA_KINDS.includes(t.kind)) vbaApply(t, st); return st.modules.filter(m => m.code.trim()).map(m => `' ===== ${m.name} =====\n${m.code.trim()}`).join('\n\n'); }
