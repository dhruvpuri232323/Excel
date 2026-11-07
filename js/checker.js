// Checks for in-app lab tasks. Every hint and message is written in the lesson; nothing here uses AI.
import { Workbook, parseRange, parseAddr, addr, rangeStr, shiftFormula, isErr, eachCell, compareValues, colIndex } from './engine.js';

const near = (a, b) => typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b)) : (typeof a === 'string' && typeof b === 'string' ? a.trim().toLowerCase() === b.trim().toLowerCase() : a === b || (isErr(a) && isErr(b) && a.code === b.code));
// text-cleaning tasks need the exact text (spaces and capitals matter); elsewhere small differences are forgiven
const same = (exact, a, b) => exact && typeof a === 'string' && typeof b === 'string' ? a === b : near(a, b);
const show = v => v === null || v === undefined ? 'blank' : isErr(v) ? v.code : typeof v === 'number' ? String(+v.toFixed(6)) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : '“' + v + '”';
function sheetIdx(wb, name) { if (name === undefined || name === null) return 0; const i = wb.sheetIndexByName(name); return i < 0 ? 0 : i; }
function fnsIn(f) { return (String(f).toUpperCase().match(/[A-Z][A-Z0-9.]*(?=\()/g) || []); }
export function randomize(wb, si, inputs, seed) {
  let x = seed * 9301 + 49297; const rnd = () => { x = (x * 9301 + 49297) % 233280; return x / 233280; };
  for (const rg of inputs || []) { const g = parseRange(rg); eachCell(g, (r, c) => { const v = wb.value(si, r, c); if (typeof v === 'number' && !wb.formula(si, r, c)) { const mag = Math.max(10, Math.abs(v)); const nv = Math.round((rnd() * 1.6 + 0.2) * mag * 100) / 100; wb.cellObj(wb.sheets[si], r, c, true).v = Number.isInteger(v) ? Math.round(nv) : nv; } }); }
  wb.invalidate();
}
function mistakeFor(task, f, v, extra = {}) {
  for (const m of task.mistakes || []) {
    const F = String(f || '').toUpperCase().replace(/\s/g, '');
    if (m.if === 'contains' && f && F.includes(m.text.toUpperCase().replace(/\s/g, ''))) return m.msg;
    if (m.if === 'notContains' && f && !F.includes(m.text.toUpperCase().replace(/\s/g, ''))) return m.msg;
    if (m.if === 'regex' && f && new RegExp(m.re, 'i').test(f)) return m.msg;
    if (m.if === 'value' && near(v, m.value)) return m.msg;
    if (m.if === 'noFormula' && !f && v !== null) return m.msg;
    if (m.if === 'error' && isErr(v)) return m.msg;
    if (m.if === 'firstOnly' && extra.firstOnly) return m.msg;
    if (m.if === 'text' && typeof v === 'string') return m.msg;
  }
  return null;
}
export function checkTask(task, ctx) {
  const wb = ctx.wb; const si = sheetIdx(wb, task.sheet);
  const fail = msg => ({ ok: false, msg }); const pass = msg => ({ ok: true, msg: msg || task.success || 'Correct.' });
  switch (task.kind) {
    case 'cell': {
      const p = parseAddr(task.cell); const f = wb.formula(si, p.r, p.c); const v = wb.value(si, p.r, p.c); const e = task.expect || {};
      if (v === null && !f) return fail(`${task.cell} is still empty.`);
      const mk = mistakeFor(task, f, v); if (mk && !(e.ref && near(v, wb.evaluate(e.ref, si)) && !(task.mistakes || []).some(m => m.strict))) { /* checked below */ }
      if (e.formula !== false && (e.formula || e.ref) && !f) return fail(mistakeFor(task, f, v) || `${task.cell} shows the right kind of answer, but it was typed in. Type a formula that starts with = so it updates when the data changes.`);
      for (const fn of e.mustUse || []) if (!fnsIn(f).includes(fn.toUpperCase()) && !String(f).toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Use ${fn} in this formula.`);
      for (const fn of e.mustNotUse || []) if (String(f || '').toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Try this without ${fn}.`);
      if (e.regex && !new RegExp(e.regex, 'i').test(String(f || '').replace(/\s/g, ''))) return fail(mistakeFor(task, f, v) || e.regexMsg || 'Not quite. Check the formula against the task.');
      const want = e.ref ? wb.evaluate(e.ref, si) : e.value;
      if (want !== undefined && !same(e.exact, v, want)) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}, but the answer should be ${show(want)}.`);
      if (e.number && typeof v !== 'number') return fail(mistakeFor(task, f, v) || `${task.cell} holds text, not a number.`);
      if (e.between && !(typeof v === 'number' && v >= e.between[0] && v <= e.between[1])) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}; it should be a number from ${e.between[0]} to ${e.between[1]}.`);
      if (e.integer && !(typeof v === 'number' && Number.isInteger(v))) return fail(mistakeFor(task, f, v) || `${task.cell} should be a whole number.`);
      if (e.oneOf) { // the answer must be one of the values in a range (random picks)
        const g = parseRange(e.oneOf); const pool = []; eachCell(g, (r, c) => pool.push(wb.value(si, r, c)));
        if (!pool.some(x => same(false, v, x))) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}, which isn’t one of the values in ${e.oneOf}.`);
      }
      if (e.ref && f && ctx.inputs && ctx.inputs.length) { // try the formula on fresh data
        for (let k = 1; k <= 3; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); const a = w2.value(si, p.r, p.c), b = w2.evaluate(e.ref, si); if (!near(a, b)) return fail(mistakeFor(task, f, a) || `This works for today’s numbers, but gives the wrong answer when the data changes. Point at the cells instead of typing numbers into the formula.`); }
      }
      const soft = mistakeFor(task, f, v); if (soft && (task.mistakes || []).find(m => m.msg === soft && m.strict)) return fail(soft);
      return pass();
    }
    case 'range': {
      const g = parseRange(task.range); const first = { r: g.r1, c: g.c1 }; let firstOk = false, bad = null, badMsg = null, empty = 0;
      const cells = []; eachCell(g, (r, c) => cells.push({ r, c }));
      for (const { r, c } of cells) {
        const f = wb.formula(si, r, c), v = wb.value(si, r, c);
        if (v === null && !f) { empty++; if (!bad) { bad = addr(r, c); badMsg = `${addr(r, c)} is still empty. Fill the formula across the whole range ${task.range}.`; } continue; }
        if (task.formula !== false && !f) { if (!bad) { bad = addr(r, c); badMsg = mistakeFor(task, f, v) || `${addr(r, c)} was typed in. Use a formula (and the fill handle) so every cell updates with the data.`; } continue; }
        const ref = shiftFormula(task.ref, r - first.r, c - first.c); const want = wb.evaluate(ref, si);
        const ok = same(task.exact, v, want); if (r === first.r && c === first.c) firstOk = ok;
        if (!ok && !bad) { bad = addr(r, c); badMsg = null; }
      }
      if (bad) { const f = wb.formula(si, parseAddr(bad).r, parseAddr(bad).c); return fail(badMsg || mistakeFor(task, f, wb.value(si, parseAddr(bad).r, parseAddr(bad).c), { firstOnly: firstOk && empty === 0 }) || (firstOk ? `The first cell is right, but ${bad} is wrong. When you copied the formula, a reference moved that should have stayed fixed. Lock it with $ (press F4).` : `${bad} doesn’t give the right answer yet.`)); }
      for (const fn of task.mustUse || []) if (!cells.every(({ r, c }) => String(wb.formula(si, r, c) || '').toUpperCase().includes(fn.toUpperCase()))) return fail(`Use ${fn} in every cell of ${task.range}.`);
      if (task.regex) { const f0 = wb.formula(si, g.r1, g.c1) || ''; if (!new RegExp(task.regex, 'i').test(f0.replace(/\s/g, ''))) return fail(task.regexMsg || 'Check the formula in the first cell.'); }
      if (ctx.inputs && ctx.inputs.length) for (let k = 1; k <= 2; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); for (const { r, c } of cells) { const ref = shiftFormula(task.ref, r - first.r, c - first.c); if (!same(task.exact, w2.value(si, r, c), w2.evaluate(ref, si))) return fail(`${addr(r, c)} breaks when the numbers change. Point at cells rather than typing values into formulas.`); } }
      return pass();
    }
    case 'spill': { // one formula whose answer spills into the cells below and to the right
      const p = parseAddr(task.cell); const f = wb.formula(si, p.r, p.c); const v = wb.value(si, p.r, p.c);
      if (!f) return fail(v === null ? `${task.cell} is still empty.` : (mistakeFor(task, f, v) || `${task.cell} was typed in. Type one formula, starting with =, in ${task.cell}.`));
      if (isErr(v) && v.code === '#SPILL!') return fail(mistakeFor(task, f, v) || `${task.cell} shows #SPILL!: something is in the way. Clear the cells inside the dashed box below and to the right of ${task.cell}, so the answer has room to spill.`);
      for (const fn of task.mustUse || []) if (!String(f).toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Use ${fn} in this formula.`);
      for (const fn of task.mustNotUse || []) if (String(f).toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Try this without ${fn}.`);
      if (task.regex && !new RegExp(task.regex, 'i').test(f.replace(/\s/g, ''))) return fail(mistakeFor(task, f, v) || task.regexMsg || 'Not quite. Check the formula against the task.');
      const same = (w, k) => { const want = w.evaluateArray(task.ref, si), got = w.spillArray(si, p.r, p.c);
        if (got.length !== want.length || got[0].length !== want[0].length) return `${task.cell} spills ${got.length} row${got.length === 1 ? '' : 's'} × ${got[0].length} column${got[0].length === 1 ? '' : 's'}; the answer is ${want.length} × ${want[0].length}${k ? ' when the data changes' : ''}.`;
        for (let i = 0; i < want.length; i++) for (let j = 0; j < want[0].length; j++) if (!near(got[i][j], want[i][j])) return `${addr(p.r + i, p.c + j)} shows ${show(got[i][j])}, but it should be ${show(want[i][j])}${k ? ' when the data changes' : ''}.`;
        return null; };
      if (task.between) { // random answers: check the shape and that every value is in range
        const got = wb.spillArray(si, p.r, p.c); const [lo, hi] = task.between;
        if (got.length !== task.rows || got[0].length !== (task.cols || 1)) return fail(mistakeFor(task, f, v) || `${task.cell} spills ${got.length} × ${got[0].length}; the task asks for ${task.rows} × ${task.cols || 1}.`);
        for (const row of got) for (const x of row) { if (typeof x !== 'number' || x < lo || x > hi) return fail(mistakeFor(task, f, v) || `${show(x)} is outside ${lo} to ${hi}.`); if (task.integer && !Number.isInteger(x)) return fail(mistakeFor(task, f, v) || `${show(x)} has decimals; the task asks for whole numbers.`); }
        return pass();
      }
      let bad = same(wb, 0); if (bad) return fail(mistakeFor(task, f, v) || bad);
      if (ctx.inputs && ctx.inputs.length) for (let k = 1; k <= 2; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); bad = same(w2, k); if (bad) return fail(`This works for today’s numbers, but not when the data changes: ${bad} Point at the cells instead of typing values into the formula.`); }
      return pass();
    }
    case 'values': {
      for (const [a, want] of Object.entries(task.cells)) {
        const p = parseAddr(a); const v = wb.value(si, p.r, p.c); const st = wb.style(si, p.r, p.c);
        if (v === null) return fail(`${a} is still empty.`);
        if (typeof want === 'number' && typeof v === 'string') return fail(mistakeFor(task, null, v) || `${a} looks like a number, but Excel stored it as text (it sits on the left). Retype it without spaces or an apostrophe.`);
        if (typeof want === 'string' && typeof v === 'number' && task.textNeeded) return fail(mistakeFor(task, null, v) || `${a} was turned into a number. Start the entry with an apostrophe (') to keep it as text.`);
        if (want && typeof want === 'object' && want.date) { if (typeof v !== 'number') return fail(`${a} isn’t a real date yet. Type it as day-month-year, like 15-08-2026.`); continue; }
        if (!near(v, want) && !(typeof want === 'string' && typeof v === 'string' && task.loose && v.toLowerCase().includes(want.toLowerCase()))) return fail(mistakeFor(task, wb.formula(si, p.r, p.c), v) || `${a} shows ${show(v)}; it should be ${show(want)}.`);
        if (task.formula === false && wb.formula(si, p.r, p.c)) return fail(`Type a plain value in ${a}, not a formula.`);
      }
      return pass();
    }
    case 'format': {
      const g = parseRange(task.range); let msg = null;
      eachCell(g, (r, c) => { if (msg) return; const st = wb.style(si, r, c); const a = addr(r, c); const fmt = st.fmt || 'General';
        if (task.fmt && !task.fmt.includes(fmt)) msg = task.fmtMsg || `${a} uses the ${fmt} format. Pick the format the task asks for.`;
        else if (task.fmtRegex && !new RegExp(task.fmtRegex).test(fmt)) msg = task.fmtMsg || `${a} doesn’t have the right number format yet.`;
        else if (task.bold !== undefined && !!st.bold !== task.bold) msg = `${a} should ${task.bold ? '' : 'not '}be bold.`;
        else if (task.fill && !st.fill) msg = `${a} needs a fill colour.`;
        else if (task.align && st.align !== task.align) msg = `${a} should be aligned ${task.align}.`;
        else if (task.wrap && !st.wrap) msg = `Turn on Wrap text for ${a}.`; });
      if (msg) return fail(msg);
      if (task.minWidth) { const s = wb.sheets[si]; const w = s.widths[colIndex(task.minWidth.col)] || 80; if (w < task.minWidth.px) return fail(`Column ${task.minWidth.col} is still too narrow to show its numbers. Widen it, or double-click its right border to AutoFit.`); }
      return pass();
    }
    case 'width': { const s = wb.sheets[si]; const w = s.widths[colIndex(task.col)] || 80; if (w < task.min) return fail(task.msg || `Column ${task.col} is still too narrow. Drag its right border, or double-click it to AutoFit.`); return pass(); }
    case 'select': {
      const sel = ctx.grid.selection(); const want = parseRange(task.range); const got = parseRange(sel.range);
      if (task.sheet && ctx.grid.si !== sheetIdx(wb, task.sheet)) return fail(`Go to the ${task.sheet} sheet first.`);
      if (rangeStr(want) !== rangeStr(got)) return fail(task.msg || `You have ${sel.range} selected. The task asks for ${task.range}.`);
      if (task.keysOnly && ctx.usedMouse) return fail('Try again using only the keyboard: click the start cell, then use the keys.');
      return pass();
    }
    case 'sorted': {
      const g = parseRange(task.range); const col = colIndex(task.col); const vals = []; for (let r = g.r1; r <= g.r2; r++) vals.push(wb.value(si, r, col));
      for (let i = 1; i < vals.length; i++) { const c = compareValues(vals[i - 1], vals[i]); if (task.desc ? c < 0 : c > 0) return fail(`The rows are not yet sorted by column ${task.col}, ${task.desc ? 'largest first' : 'smallest first'}. Click a cell in column ${task.col}, then use the sort button.`); }
      if (task.keep) { // rows stay together: each row's key cell still matches its partner
        const kc = colIndex(task.keep.col), pc = colIndex(task.keep.with); const orig = new Map(task.keep.pairs); for (let r = g.r1; r <= g.r2; r++) { const k = wb.value(si, r, kc); if (orig.has(String(k)) && !near(wb.value(si, r, pc), orig.get(String(k)))) return fail('Some rows got mixed up: a value moved without the rest of its row. Undo, select one cell in the data, and sort the whole table.'); }
      }
      return pass();
    }
    case 'name': {
      const d = wb.names[task.name.toLowerCase()]; if (!d) return fail(`There is no name called ${task.name} yet. Select ${task.ref}, click the Name Box, type ${task.name} and press Enter.`);
      const got = d.ref.replace(/\$/g, '').replace(/^.*!/, ''); if (got.toUpperCase() !== task.ref.toUpperCase()) return fail(`${task.name} points to ${got}. It should point to ${task.ref}.`);
      return pass();
    }
    case 'sheet': {
      const i = wb.sheetIndexByName(task.name); if (i < 0) return fail(task.msg || `There’s no sheet called “${task.name}” yet. Double-click a sheet tab to rename it.`);
      if (task.index !== undefined && i !== task.index) return fail(`“${task.name}” is there, but it should be sheet number ${task.index + 1} from the left.`);
      return pass();
    }
    case 'noBlankRows': { const g = parseRange(task.range); const blankRow = r => { for (let c = g.c1; c <= g.c2; c++) if (!wb.isEmpty(si, r, c)) return false; return true; }; let first = -1, last = -1; for (let r = g.r1; r <= g.r2; r++) if (!blankRow(r)) { if (first < 0) first = r; last = r; } for (let r = first; r <= last && first >= 0; r++) if (blankRow(r)) return fail(`Row ${r + 1} is still blank. Right-click its row number and choose Delete row.`); if (task.rows && last - first + 1 !== task.rows) return fail(`The list should have ${task.rows} rows; it has ${last - first + 1}.`); return pass(); }
    case 'count': { const g = parseRange(task.range); let n = 0; eachCell(g, (r, c) => { if (!wb.isEmpty(si, r, c)) n++; }); if (n !== task.equals) return fail(task.msg || `${task.range} has ${n} filled cells; expected ${task.equals}.`); return pass(); }
    case 'mc': { if (ctx.answer === undefined || ctx.answer === null) return fail('Pick an answer.'); return ctx.answer === task.answer ? pass(task.explain) : fail((task.why && task.why[ctx.answer]) || task.explainWrong || 'Not quite. Have another look.'); }
    case 'hotspot': { if (!ctx.answer) return fail('Click on the screenshot.'); return ctx.answer === task.answer ? pass(task.explain) : fail(`That’s the ${ctx.answer}. ${task.hintWrong || 'Look again.'}`); }
    case 'order': { const a = ctx.answer || []; const ok = a.length === task.items.length && a.every((x, i) => x === task.items[i]); return ok ? pass(task.explain) : fail('Not in the right order yet. Think about what Excel needs first.'); }
    case 'keys': { const log = ctx.keyLog || []; const want = task.combo.toLowerCase(); const hit = log.some(k => [k.ctrl && 'ctrl', k.shift && 'shift', k.alt && 'alt', k.key.toLowerCase()].filter(Boolean).join('+') === want); return hit ? pass() : fail(`Press ${task.label || task.combo} while the grid is selected.`); }
  }
  return fail('Unknown task type');
}
// Apply a task's model answer to the grid (used by "Show solution" and by the automated tests).
export function applySolution(task, ctx) {
  const wb = ctx.wb, grid = ctx.grid; const si = sheetIdx(wb, task.sheet);
  switch (task.kind) {
    case 'cell': case 'spill': { const p = parseAddr(task.cell); wb.setInput(si, p.r, p.c, task.solution); break; }
    case 'range': { const g = parseRange(task.range); eachCell(g, (r, c) => wb.setInput(si, r, c, shiftFormula(task.solutionFormula || task.ref, r - g.r1, c - g.c1))); break; }
    case 'count': case 'values': for (const [a, v] of Object.entries(task.solutionValues || task.cells)) { const p = parseAddr(a); wb.setInput(si, p.r, p.c, typeof v === 'object' && v && v.date ? v.date : String(v)); } break;
    case 'format': { const g = parseRange(task.range); const st = {}; if (task.fmt) st.fmt = task.fmt[0] === 'General' ? null : task.fmt[0]; if (task.solutionFmt) st.fmt = task.solutionFmt; if (task.bold !== undefined) st.bold = task.bold || null; if (task.fill) st.fill = '#FFF2CC'; if (task.align) st.align = task.align; if (task.wrap) st.wrap = true; eachCell(g, (r, c) => wb.setStyle(si, r, c, st)); if (task.minWidth) wb.sheets[si].widths[colIndex(task.minWidth.col)] = task.minWidth.px + 10; break; }
    case 'width': wb.sheets[si].widths[colIndex(task.col)] = task.min + 10; wb.invalidate(); break;
    case 'select': if (task.sheet) grid.setSheet(sheetIdx(wb, task.sheet)); grid.select(task.range); break;
    case 'sorted': { const g = parseRange(task.region || task.range); wb.sortRegion(si, g, [{ col: colIndex(task.col), desc: !!task.desc }], !!task.region); break; }
    case 'name': wb.defineName(task.name, (task.sheet ? `'${task.sheet}'!` : '') + task.ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')); break;
    case 'sheet': { const i = task.from !== undefined ? task.from : (task.index ?? 0); if (wb.sheets[i]) wb.sheets[i].name = task.name; else wb.addSheet(task.name); wb.invalidate(); break; }
    case 'noBlankRows': { const g = parseRange(task.range); for (let r = g.r2; r >= g.r1; r--) { let blank = true; for (let c = g.c1; c <= g.c2; c++) if (!wb.isEmpty(si, r, c)) blank = false; if (blank) wb.deleteRows(si, r, 1); } break; }
  }
  grid && grid.render();
}

// "What Sheetah checked": the checks behind a task, in plain words (shown under a result)
export function describeCheck(task, inputs) {
  const e = task.expect || {}; const out = [];
  const f = x => String(x || '').replace(/^=/, '');
  switch (task.kind) {
    case 'cell':
      if (e.formula !== false && (e.formula || e.ref)) out.push(`${task.cell} holds a formula (it starts with =), not a typed-in number`);
      if (e.ref) out.push(`it gives the same answer as ${f(e.ref)}`); else if (e.value !== undefined) out.push(`it shows ${e.value}`);
      (e.mustUse || []).forEach(fn => out.push(`it uses ${fn}`)); (e.mustNotUse || []).forEach(fn => out.push(`it doesn’t use ${fn}`));
      if (e.between) out.push(`it is a number from ${e.between[0]} to ${e.between[1]}${e.integer ? ', with no decimals' : ''}`);
      if (e.oneOf) out.push(`it shows one of the values in ${e.oneOf}`);
      if (e.ref && inputs && inputs.length) out.push('it still gives the right answer when Sheetah changes the numbers in the data (3 different sets)');
      break;
    case 'spill':
      out.push(`${task.cell} holds one formula (it starts with =), and its answer spills into the cells next to it`); out.push(task.between ? `it spills ${task.rows} × ${task.cols || 1} values, each from ${task.between[0]} to ${task.between[1]}${task.integer ? ', whole numbers' : ''}` : `the whole spilled answer matches ${f(task.ref)}, cell by cell`);
      (task.mustUse || []).forEach(fn => out.push(`it uses ${fn}`)); if (inputs && inputs.length) out.push('it still gives the right answer when Sheetah changes the numbers (2 different sets)');
      break;
    case 'range':
      out.push(`every cell in ${task.range} holds a formula`); out.push(`each one gives the right answer for its own row or column (the first is ${f(task.ref)})`);
      (task.mustUse || []).forEach(fn => out.push(`every formula uses ${fn}`)); if (inputs && inputs.length) out.push('they still work when Sheetah changes the numbers (2 different sets)');
      break;
    case 'values': out.push(...Object.entries(task.cells).map(([a, v]) => `${a} holds ${typeof v === 'object' && v && v.date ? 'a real date' : typeof v === 'number' ? `the number ${v}` : `“${v}”`}`)); if (task.formula === false) out.push('they are typed values, not formulas'); break;
    case 'format': out.push(`every cell in ${task.range}${task.bold ? ' is bold' : ''}${task.fill ? ' has a fill colour' : ''}${task.align ? ` is aligned ${task.align}` : ''}${task.wrap ? ' wraps its text' : ''}${task.fmt || task.fmtRegex ? ' has the right number format' : ''}`); break;
    case 'width': out.push(`column ${task.col} is at least ${task.min} pixels wide`); break;
    case 'select': out.push(`exactly ${task.range} is selected${task.keysOnly ? ', using only the keyboard' : ''}`); break;
    case 'sorted': out.push(`the rows are in order by column ${task.col}, ${task.desc ? 'largest first' : 'smallest first'}`); if (task.keep) out.push('each row stayed together when it moved'); break;
    case 'name': out.push(`a name called ${task.name} exists`, `it points to ${task.ref}${task.sheet ? ' on ' + task.sheet : ''}`); break;
    case 'sheet': out.push(`a sheet called “${task.name}” exists`); if (task.index !== undefined) out.push(`it is sheet number ${task.index + 1} from the left`); break;
    case 'noBlankRows': out.push(`there are no empty rows between the first and last record in ${task.range}`); break;
    case 'count': out.push(`${task.range} has ${task.equals} filled cells`); break;
    case 'keys': out.push(`you pressed ${task.label || task.combo}`); break;
    case 'mc': out.push('the option you chose'); break; case 'hotspot': out.push('the part of the screenshot you clicked'); break; case 'order': out.push('the order of the steps'); break;
  }
  return out;
}
