// Checking Office Scripts tasks: run the learner's script and the model script on fresh copies of the lab's workbook
// (and on the task's other cases), then compare the sheets, the Output lines and the returned value.
import { Workbook, parseRange } from './engine.js';
import { runScript } from './oscript.js';
import { diffBooks } from './vbacheck.js';

export const OS_KINDS = ['os', 'osfn', 'oscode'];
const clone = x => JSON.parse(JSON.stringify(x));
const lc = s => String(s || '').toLowerCase();
export function osStart(lab) { const v = lab.os || {}; return { scripts: (v.scripts || []).map(s => ({ ...s })), cur: (v.scripts || [])[0] ? v.scripts[0].name : null, sheet: null, active: 0 }; }
export function osApply(task, st) { for (const s of (task.msol && task.msol.scripts) || []) { const have = st.scripts.find(x => lc(x.name) === lc(s.name)); if (have) have.code = s.code; else st.scripts.push({ ...s }); st.cur = s.name; } }
function edit(wb, edits) { for (const [a, v] of Object.entries(edits || {})) { const m = /^(?:'?(.+?)'?!)?([A-Z]+\d+)$/.exec(a); const si = m[1] ? wb.sheets.findIndex(s => lc(s.name) === lc(m[1])) : 0; const g = parseRange(m[2]); if (v === null) wb.clear(si, g.r1, g.c1, true); else if (typeof v === 'number') wb.setRaw(wb.sheets[si], g.r1, g.c1, v); else wb.setInput(si, g.r1, g.c1, String(v)); } wb.invalidate(); }
function runOnCopy(code, lab, task, variant) { const wb = new Workbook(clone(variant.sheet || task.start || lab.sheet)); edit(wb, variant.edits); const sel = variant.select || task.select; let active = 0, s = null;
  if (sel) { const m = /^'?(.+?)'?!(.+)$/.exec(sel); active = m ? Math.max(0, wb.sheets.findIndex(x => lc(x.name) === lc(m[1]))) : 0; s = { s: wb.sheets[active], g: parseRange(m ? m[2] : sel) }; }
  const r = runScript(code, wb, { args: variant.args || task.args || [], active, sel: s }); return { ...r, wb }; }
const say = v => v === undefined ? 'nothing (undefined)' : typeof v === 'string' ? `“${v}”` : JSON.stringify(v);
function sameVal(a, b) { if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9 * Math.max(1, Math.abs(b)); return JSON.stringify(a) === JSON.stringify(b); }
export function describeOS(task) { const c = task.compare || {}; if (task.kind === 'oscode') return [(task.need || []).map(n => n.what || '').filter(Boolean).join('; ') || 'the script itself'];
  if (task.kind === 'osfn') return [`main gives back the model answer’s value for each case`]; return ['the workbook after the script runs (values' + (c.styles ? ', ' + c.styles.join(', ') : '') + ')', ...(c.out ? ['the Output lines'] : []), ...(task.alsoWith || []).map(v => 'again ' + (v.label || 'with other data'))]; }
export function checkOS(task, ctx) { const fail = msg => ({ ok: false, msg }); const pass = () => ({ ok: true, msg: task.success || task.explain || 'Correct.' });
  const lab = (ctx.os && ctx.os.lab) || {}; const st = (ctx.os && ctx.os.state) || osStart(lab); const name = task.script;
  const mine = st.scripts.find(s => lc(s.name) === lc(name)); if (!mine) return fail(`There’s no script called “${name}” yet. Automate › New Script, then double-click its name to rename it ${name}.`);
  const model = ((task.msol && task.msol.scripts) || []).find(s => lc(s.name) === lc(name)) || (task.msol && task.msol.scripts && task.msol.scripts[0]);
  if (task.kind === 'oscode' || task.need || task.not) { for (const n of task.need || []) { const re = new RegExp(n.re, n.flags || 'm'); if (!re.test(mine.code)) return fail(n.msg); } for (const n of task.not || []) { const re = new RegExp(n.re, n.flags || 'm'); if (re.test(mine.code)) return fail(n.msg); } if (task.kind === 'oscode') return pass(); }
  const variants = [{}, ...(task.alsoWith || [])]; const cmp = task.compare || {};
  for (let i = 0; i < variants.length; i++) { const v = variants[i]; const lead = i ? `Right for the lab’s data, but not ${v.label || 'with other data'}: ` : '';
    const a = runOnCopy(mine.code, lab, task, v), b = runOnCopy(model.code, lab, task, v); if (b.error) return fail(`(Sheetah’s model script failed here: ${b.error.message})`);
    if (a.error) return fail(`${lead}${name} stops${a.error.line ? ` on line ${a.error.line}` : ''}: ${a.error.message}${/undefined \(reading/.test(a.error.message) ? ' Something you asked for doesn’t exist: check sheet and range names (getWorksheet gives back undefined when there’s no such sheet).' : ''}`);
    let d = null;
    if (task.kind === 'osfn') { if (!sameVal(a.value, b.value)) d = `main(workbook${(v.args || task.args || []).map(x => ', ' + JSON.stringify(x)).join('')}) should give back ${say(b.value)}; yours gives back ${say(a.value)}.${a.value === undefined ? ' Add a return statement, and a return type after the brackets.' : ''}`; }
    const words = t => t && t.replace(/the macro has run/g, 'the script has run').replace(/the macro runs/g, 'the script runs').replace(/Your macro typed/g, 'Your script typed').replace(/set \.Formula, not \.Value/g, 'use setFormula, not setValue').replace(/the macro should/gi, 'the script should');
    if (!d && cmp.values !== false && task.kind !== 'osfn') d = words(diffBooks(a.wb, b.wb, cmp));
    if (!d && task.kind === 'osfn' && cmp.sheet) d = words(diffBooks(a.wb, b.wb, cmp));
    if (!d && cmp.out) { const A = a.out, B = b.out; for (let k = 0; k < Math.max(A.length, B.length) && !d; k++) if (A[k] !== B[k]) d = B[k] === undefined ? `The Output should have ${B.length} line${B.length === 1 ? '' : 's'}; yours has more: “${A[k]}”.` : A[k] === undefined ? `The Output should show “${B[k]}” as line ${k + 1}; yours stops after ${A.length} line${A.length === 1 ? '' : 's'}.` : `Output line ${k + 1} should be “${B[k]}”, but it is “${A[k]}”.`; }
    if (!d && cmp.active && a.active !== b.active) d = `When the script finishes, the ${b.wb.sheets[b.active].name} sheet should be the active one (call .activate() on it).`;
    if (d) return fail(lead + d + (i && v.why ? ' ' + v.why : (!i && task.whyMain ? ' ' + task.whyMain : ''))); }
  return pass(); }
