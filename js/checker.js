// Checks for in-app lab tasks. Every hint and message is written in the lesson; nothing here uses AI.
import { Workbook, parseRange, parseAddr, addr, rangeStr, shiftFormula, isErr, eachCell, compareValues, colIndex, isDateFmt } from './engine.js';
import { chartFromConfig, axisStart } from './charts.js';
import { checkPQ, pqApply } from './pqcheck.js';
import { checkModel, modelApply, describeModel, MODEL_KINDS } from './modelcheck.js';
import { checkVBA, vbaApply, describeVBA, VBA_KINDS } from './vbacheck.js';
import { checkOS, osApply, describeOS, OS_KINDS } from './oscheck.js';
// the values of a block of cells, for charts: 'A1:B13' or 'Sheet!A1:B13'
export function blockOf(wb, si, ref, depth = 0) { let t = String(ref || '').trim().replace(/^=/, ''); if (!t || depth > 3) return null;
  const isName = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(t) && !/^[A-Za-z]{1,3}\d+$/.test(t);
  const nm = isName ? wb.names[t.toLowerCase()] : null; if (nm) return blockOf(wb, si, nm.ref, depth + 1); // a defined name, like Top5
  const tb = isName && wb.tableByName ? wb.tableByName(t) : null; if (tb) return blockOf(wb, tb.si === si ? si : tb.si, (tb.si === si ? '' : `'${wb.sheets[tb.si].name}'!`) + tb.t.ref, depth + 1); // a Table: headers and data, growing with it
  if (t.includes('[') && wb.structured) { try { return blockOf(wb, si, wb.structured(t, si), depth + 1); } catch (e) { return null; } } // Sales[[Month]:[Revenue]]
  const sp = /^(?:'?(.+?)'?!)?\$?([A-Za-z]+)\$?(\d+)#$/.exec(t); // a spilled range, like E2#
  if (sp) { const i = sp[1] ? wb.sheetIndexByName(sp[1]) : si; if (i < 0) return null; const a = parseRange(sp[2].toUpperCase() + sp[3]); const b = wb.spillBlock(i, a.r1, a.c1); const g = b ? b.g : a; const out = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const v = wb.value(i, r, c); row.push(isErr(v) ? null : v); } out.push(row); } return out; }
  const m = /^(?:'?(.+?)'?!)?(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)$/.exec(t); if (!m) return null; const i = m[1] ? wb.sheetIndexByName(m[1]) : si; if (i < 0) return null; const g = parseRange(m[2].replace(/\$/g, '').toUpperCase()); if (!g) return null; const out = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const v = wb.value(i, r, c); row.push(isErr(v) ? null : v); } out.push(row); } return out; }
export function chartOf(wb, si, cfg) { let fmt = 'num'; try { const g = parseRange(String(cfg.range).replace(/^.*!/, '').replace(/\$/g, '').toUpperCase()); const r = Math.min(g.r1 + 1, g.r2); const f = (wb.style(si, r, g.c2) || {}).fmt || ''; fmt = /%/.test(f) ? 'pct' : /#,##,##0/.test(f) ? 'inr' : 'num'; if (cfg.type === 'bar' && cfg.stacked) { for (let c = g.c1; c <= g.c2; c++) { const fc = (wb.style(si, r, c) || {}).fmt || ''; if (typeof wb.value(si, r, c) === 'number' && isDateFmt(fc)) { fmt = 'date'; break; } } } } catch (e) {} return chartFromConfig(cfg, ref => blockOf(wb, si, ref), fmt); }

const near = (a, b) => typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b)) : (typeof a === 'string' && typeof b === 'string' ? a.trim().toLowerCase() === b.trim().toLowerCase() : a === b || (isErr(a) && isErr(b) && a.code === b.code));
// text-cleaning tasks need the exact text (spaces and capitals matter); elsewhere small differences are forgiven
const same = (exact, a, b) => exact && typeof a === 'string' && typeof b === 'string' ? a === b : near(a, b);
const show = v => v === null || v === undefined ? 'blank' : isErr(v) ? v.code : typeof v === 'number' ? String(+v.toFixed(6)) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : '“' + v + '”';
function sheetIdx(wb, name) { if (name === undefined || name === null) return 0; const i = wb.sheetIndexByName(name); return i < 0 ? 0 : i; }
function fnsIn(f) { return (String(f).toUpperCase().match(/[A-Z][A-Z0-9.]*(?=\()/g) || []); }
export function randomize(wb, si0, inputs, seed) {
  let x = seed * 9301 + 49297; const rnd = () => { x = (x * 9301 + 49297) % 233280; return x / 233280; };
  for (const rg0 of inputs || []) { let rg = rg0, si = si0; const bang = rg0.lastIndexOf('!'); if (bang > 0) { const i = wb.sheetIndexByName(rg0.slice(0, bang).replace(/^'|'$/g, '')); if (i < 0) continue; si = i; rg = rg0.slice(bang + 1); } /* 'TB!B2:C30': inputs on another sheet */ const g = parseRange(rg); eachCell(g, (r, c) => { const v = wb.value(si, r, c); if (typeof v === 'number' && !wb.formula(si, r, c)) { const pct = Math.abs(v) <= 1 && v !== 0 && /%/.test((wb.style(si, r, c) || {}).fmt || ''); /* rates stay rates: 8.5% becomes 1.7%–15.3%, not 1,700% */ if (pct) { wb.cellObj(wb.sheets[si], r, c, true).v = Math.round((rnd() * 1.6 + 0.2) * v * 1e4) / 1e4; return; } const mag = Math.max(10, Math.abs(v)); const nv = Math.sign(v || 1) * Math.round((rnd() * 1.6 + 0.2) * mag * 100) / 100; /* keep the sign: a cost stays negative */ wb.cellObj(wb.sheets[si], r, c, true).v = Number.isInteger(v) ? Math.round(nv) : nv; } }); }
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
  if (task.kind === 'pq' || task.kind === 'pqload') return checkPQ(task, ctx);
  if (MODEL_KINDS.includes(task.kind)) return checkModel(task, ctx);
  if (VBA_KINDS.includes(task.kind)) return checkVBA(task, ctx);
  if (OS_KINDS.includes(task.kind)) return checkOS(task, ctx);
  if (task.alsoWith && ctx.wb) { // the answer must also follow when a selector cell changes (drop-downs, form controls)
    const res = checkTask({ ...task, alsoWith: null }, ctx); if (!res.ok) return res; const si = sheetIdx(ctx.wb, task.sheet);
    for (const set of task.alsoWith) { const w2 = new Workbook(ctx.wb.toSpec()); for (const [a, v] of Object.entries(set)) { const p = parseAddr(a); w2.setInput(si, p.r, p.c, String(v)); }
      const r2 = checkTask({ ...task, alsoWith: null }, { ...ctx, wb: w2, inputs: null, grid: null });
      if (!r2.ok) { const [a, v] = Object.entries(set)[0]; return { ok: false, msg: task.alsoMsg || `Right for what is picked now, but not when ${a} changes to ${show(v)}: ${r2.msg} Make the formula read ${a}, so the answer follows the selector.` }; } }
    return res; }
  const wb = ctx.wb; const si = sheetIdx(wb, task.sheet);
  const fail = msg => ({ ok: false, msg }); const pass = msg => ({ ok: true, msg: msg || task.success || 'Correct.' });
  switch (task.kind) {
    case 'cell': {
      const p = parseAddr(task.cell); const f = wb.formula(si, p.r, p.c); const v = wb.value(si, p.r, p.c); const e = task.expect || {};
      if (v === null && !f) return fail(`${task.cell} is still empty.`);
      const mk = mistakeFor(task, f, v); if (mk && !(e.ref && near(v, wb.evaluate(e.ref, si, p)) && !(task.mistakes || []).some(m => m.strict))) { /* checked below */ }
      if (e.formula !== false && (e.formula || e.ref) && !f) return fail(mistakeFor(task, f, v) || `${task.cell} shows the right kind of answer, but it was typed in. Type a formula that starts with = so it updates when the data changes.`);
      for (const fn of e.mustUse || []) if (!fnsIn(f).includes(fn.toUpperCase()) && !String(f).toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Use ${fn} in this formula.`);
      for (const fn of e.mustNotUse || []) if (String(f || '').toUpperCase().includes(fn.toUpperCase())) return fail(mistakeFor(task, f, v) || `Try this without ${fn}.`);
      if (e.regex && !new RegExp(e.regex, 'i').test(String(f || '').replace(/\s/g, ''))) return fail(mistakeFor(task, f, v) || e.regexMsg || 'Not quite. Check the formula against the task.');
      const want = e.ref ? wb.evaluate(e.ref, si, p) : e.value;
      if (want !== undefined && !same(e.exact, v, want) && typeof v === 'number' && typeof want === 'number' && want !== 0 && near(-v, want)) return fail(mistakeFor(task, f, v) || e.signMsg || `${task.cell} shows ${show(v)}: the right amount, but the sign is the wrong way round.${/\b(PMT|IPMT|PPMT|PV|FV|NPER|CUMIPMT|CUMPRINC|XNPV|NPV)\(/i.test(String(f || '')) ? ' Excel’s finance functions show money you pay out as negative and money you receive as positive; put a minus in front of the function (or of the amount you pass in) to flip it.' : ' Check which value you subtract from which.'}`);
      if (want !== undefined && !same(e.exact, v, want)) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}, but the answer should be ${show(want)}.`);
      for (const part of e.contains || []) { const want = String(part).startsWith('=') ? wb.evaluate(part, si) : part; const txt = typeof want === 'number' ? String(want) : String(want); if (!String(v ?? '').toLowerCase().includes(txt.toLowerCase())) return fail(mistakeFor(task, f, v) || e.containsMsg || `${task.cell} shows ${show(v)}; it should include ${show(txt)}.`); }
      if (e.number && typeof v !== 'number') return fail(mistakeFor(task, f, v) || `${task.cell} holds text, not a number.`);
      if (e.between && !(typeof v === 'number' && v >= e.between[0] && v <= e.between[1])) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}; it should be a number from ${e.between[0]} to ${e.between[1]}.`);
      if (e.integer && !(typeof v === 'number' && Number.isInteger(v))) return fail(mistakeFor(task, f, v) || `${task.cell} should be a whole number.`);
      if (e.oneOf) { // the answer must be one of the values in a range (random picks)
        const g = parseRange(e.oneOf); const pool = []; eachCell(g, (r, c) => pool.push(wb.value(si, r, c)));
        if (!pool.some(x => same(false, v, x))) return fail(mistakeFor(task, f, v) || `${task.cell} shows ${show(v)}, which isn’t one of the values in ${e.oneOf}.`);
      }
      if (e.ref && f && ctx.inputs && ctx.inputs.length) { // try the formula on fresh data
        for (let k = 1; k <= 3; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); const a = w2.value(si, p.r, p.c), b = w2.evaluate(e.ref, si, p); if (!near(a, b)) return fail(mistakeFor(task, f, a) || `This works for today’s numbers, but gives the wrong answer when the data changes. Point at the cells instead of typing numbers into the formula.`); }
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
        const ref = shiftFormula(task.ref, r - first.r, c - first.c); const want = wb.evaluate(ref, si, { r, c });
        const ok = same(task.exact, v, want); if (r === first.r && c === first.c) firstOk = ok;
        if (!ok && !bad) { bad = addr(r, c); badMsg = null; }
      }
      if (bad) { const f = wb.formula(si, parseAddr(bad).r, parseAddr(bad).c); return fail(badMsg || mistakeFor(task, f, wb.value(si, parseAddr(bad).r, parseAddr(bad).c), { firstOnly: firstOk && empty === 0 }) || (firstOk ? `The first cell is right, but ${bad} is wrong. When you copied the formula, a reference moved that should have stayed fixed. Lock it with $ (press F4).` : `${bad} doesn’t give the right answer yet.`)); }
      for (const fn of task.mustUse || []) if (!cells.every(({ r, c }) => String(wb.formula(si, r, c) || '').toUpperCase().includes(fn.toUpperCase()))) return fail(`Use ${fn} in every cell of ${task.range}.`);
      if (task.regex) { const f0 = wb.formula(si, g.r1, g.c1) || ''; if (!new RegExp(task.regex, 'i').test(f0.replace(/\s/g, ''))) return fail(task.regexMsg || 'Check the formula in the first cell.'); }
      if (ctx.inputs && ctx.inputs.length) for (let k = 1; k <= 2; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); for (const { r, c } of cells) { const ref = shiftFormula(task.ref, r - first.r, c - first.c); if (!same(task.exact, w2.value(si, r, c), w2.evaluate(ref, si, { r, c }))) return fail(mistakeFor(task, w2.formula(si, r, c), w2.value(si, r, c)) || `${addr(r, c)} gives the right answer today, but not when the data changes. Point at cells rather than typing values into formulas.`); } }
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
    case 'chart': { // the lab chart builder: data, type and the design choices the task asks for
      const cfg = ctx.answer; if (!cfg || !cfg.range) return fail('Choose the data first: type a range like A1:B13 in the chart box, or select the cells and click Use selection.');
      const e = task.expect || {}; const got = chartOf(wb, si, cfg); if (!got.data.series.length) return fail(`${cfg.range} has no numbers to chart. Include the column of values.`);
      const words = s => String(s || '').toLowerCase(); const types = String(e.type || '').split('|').filter(Boolean);
      if (e.range) { const want = chartOf(wb, si, { ...cfg, range: e.range, byRow: !!e.byRow }).data; const d = got.data;
        const nm = x => x.series.map(s => s.name).join('|'); const vals = x => JSON.stringify(x.series.map(s => s.values.map(v => v === null ? null : Math.round(v * 1e6) / 1e6)));
        if (d.series.length > want.series.length) return fail(`The chart shows ${d.series.length} series; it needs ${want.series.length === 1 ? 'just one: ' + want.series[0].name : want.series.length + ': ' + want.series.map(s => s.name).join(', ')}. Check the range, or Switch Row/Column.`);
        if (d.cats.length !== want.cats.length || d.cats.some((c, i) => words(c) !== words(want.cats[i]))) return fail(e.dataMsg || `The categories along the axis should be ${want.cats.slice(0, 4).join(', ')}${want.cats.length > 4 ? '…' : ''}. Check the range${want.cats.length && d.series.length && d.cats.length === d.series.length ? ', or Switch Row/Column' : ''}.`);
        if (vals(d) !== vals(want)) return fail(e.dataMsg || 'The chart shows the wrong numbers. Check which columns the range covers.');
        if (e.names !== false && nm(d) !== nm(want) && want.series.some(s => !/^Series\d/.test(s.name))) return fail(`Include the heading row in the range, so the series is called ${want.series.map(s => s.name).join(' and ')}.`); }
      if (types.length && !types.includes(cfg.type)) return fail(e.typeMsg || `This data needs a different chart type. ${task.typeHint || ''}`.trim());
      if (e.labels !== undefined && !!cfg.labels !== !!e.labels) return fail(e.labels ? 'Turn on data labels, so readers see the exact values.' : 'Turn off data labels here: they crowd this chart.');
      if (e.legend === 'none' && cfg.legend !== 'none' && !(cfg.legend === 'auto' && got.data.series.length < 2 && !['pie', 'doughnut'].includes(cfg.type))) return fail('Remove the legend: with labels or one series, it only repeats what’s on the chart.');
      if (e.legend === 'shown' && (cfg.legend === 'none' || (cfg.legend === 'auto' && got.data.series.length < 2))) return fail('Show a legend, so readers can tell the series apart.');
      if (e.grid === false && cfg.grid !== false) return fail('Remove the gridlines: the labels already give the values.');
      if (e.yZero && axisStart(got.spec) !== 0) return fail(`The value axis starts at ${(+axisStart(got.spec)).toLocaleString('en-IN')}, which exaggerates the differences. Set Axis starts at: 0.`);
      if (e.endLabels && !cfg.endLabels) return fail('Label each line at its end, so readers don’t have to match colours to a legend.');
      if (e.highlight && words(cfg.highlight) !== words(e.highlight)) return fail(e.highlightMsg || `Colour ${e.highlight} in the accent colour and leave the rest grey, so the eye goes straight to it.`);
      if (e.noHighlight && cfg.highlight) return fail('Keep every bar the same colour here: no single bar is the point.');
      if (e.secondary) { if (cfg.type !== 'combo') return fail('Use a combo chart: columns for one measure, a line for the other.'); if (words(cfg.secondary) !== words(e.secondary)) return fail(`Put ${e.secondary} on the secondary axis as a line.`); }
      if (e.stacked !== undefined && (cfg.stacked || '') !== e.stacked) return fail(e.stacked === 'percent' ? 'Use 100% stacked, so each bar shows shares that add up to 100%.' : e.stacked === 'stacked' ? 'Stack the series, so each bar shows the total and its parts.' : 'Don’t stack the series here: put them side by side.');
      if (e.largestFirst) { const v = got.data.series[0].values; const ordered = cfg.type === 'bar' ? (cfg.reverse ? v : v.slice().reverse()) : v; if (ordered.some((x, i) => i && x > ordered[i - 1])) return fail(cfg.type === 'bar' ? 'Put the biggest bar at the top: sort the data biggest first and tick Categories in reverse order (or sort it smallest first).' : 'Sort the bars biggest first.'); }
      if (e.reverse !== undefined && !!cfg.reverse !== !!e.reverse) return fail(e.reverse ? 'Tick Categories in reverse order, so the first row of data is at the top.' : 'Untick Categories in reverse order.');
      const nv = v => v === '' || v === null || v === undefined || v === 'auto' || !isFinite(+v) ? null : +v;
      if (e.binWidth !== undefined && !(cfg.bins === 'width' && Math.abs((nv(cfg.binVal) ?? -1) - e.binWidth) < 1e-9)) return fail(e.binMsg || `Set Bins to Bin width, and give it a width of ${e.binWidth.toLocaleString('en-IN')}.`);
      if (e.binCount !== undefined && !(cfg.bins === 'count' && nv(cfg.binVal) === e.binCount)) return fail(e.binMsg || `Set Bins to Number of bins: ${e.binCount}.`);
      if (e.underflow !== undefined && nv(cfg.underflow) !== e.underflow) return fail(e.flowMsg || `Set the underflow bin to ${e.underflow.toLocaleString('en-IN')}, so the small values share one bar.`);
      if (e.overflow !== undefined && nv(cfg.overflow) !== e.overflow) return fail(e.flowMsg || `Set the overflow bin to ${e.overflow.toLocaleString('en-IN')}, so the big values share one bar.`);
      if (e.totals) { const have = (cfg.totals || []).map(words).sort().join('|'), want = e.totals.map(words).sort().join('|'); if (have !== want) return fail(e.totalsMsg || `Set the points that are totals, not changes, as totals: ${e.totals.join(', ')}. The others should float.`); }
      if (e.hideFirst && !cfg.hideFirst) return fail(e.hideMsg || 'Make the first series invisible (no fill), so the bars float at their start.');
      if (e.invert && !cfg.invert) return fail('Tick Invert if negative, so bars below zero get their own colour.');
      if (e.yMin !== undefined && axisStart(got.spec) !== e.yMin) return fail(e.yMinMsg || `Set the axis to start at ${e.yMin.toLocaleString('en-IN')}.`);
      if (e.dynamic) { const rg = String(cfg.range).trim().replace(/^=/, ''); const okDyn = /#$/.test(rg) || wb.names[rg.toLowerCase()] || (wb.tableByName && (wb.tableByName(rg) || (/\[/.test(rg) && wb.tableByName(rg.split('[')[0])))); if (!okDyn || (e.dynamic === 'table' && !wb.tableByName(rg.split('[')[0])) || (e.dynamic === 'spill' && !/#$/.test(rg))) return fail(e.dynamicMsg || 'Point the chart at something that grows: the spilled range itself (like D2#), a Table, or a name for it.'); }
      if (e.titleCell) { const tc = String(cfg.title || '').replace(/\s|\$/g, '').toUpperCase(); if (!new RegExp(`^=('?[A-Z0-9 _]+'?!)?${e.titleCell.toUpperCase()}$`).test(tc)) return fail(e.titleMsg || `Link the title to ${e.titleCell}: type =${e.titleCell} in the Title box, so it changes when the cell does.`); }
      if (e.title) { const t = words(got.spec.title).trim(); if (!t) return fail('Give the chart a title that says what it shows.');
        if ((e.title.not || ['chart title']).some(x => t === words(x))) return fail(e.title.notMsg || 'That title just names the topic. Say the finding: what should the reader notice?');
        if (e.title.has && !e.title.has.some(h => String(h).split('+').every(p => t.includes(words(p))))) return fail(e.title.msg || `Make the title say the finding: mention ${e.title.has[0].split('+').join(' and ')}.`);
        if (e.title.min && t.split(/\s+/).length < e.title.min) return fail(e.title.msg || 'Make the title a short sentence that states the finding, not one or two words.'); }
      return pass();
    }
    case 'lambda': { // a named LAMBDA: it exists, and gives the model answer on the test calls (and on fresh numbers)
      const d = wb.names[task.name.toLowerCase()];
      if (!d) return fail(`There’s no function called ${task.name} yet. Write the LAMBDA in ${task.cell || 'a cell'}, then with that cell selected, type ${task.name} in the Name Box and press Enter.`);
      if (!/^=?\s*(_xlfn\.)?LAMBDA\s*\(/i.test(d.ref)) return fail(`${task.name} points to cells, not a LAMBDA. Select the cell with your =LAMBDA(…) formula, type ${task.name} in the Name Box and press Enter.`);
      const model = w => { const w2 = new Workbook(w.toSpec()); w2.defineName(task.name, task.ref); return w2; };
      const tryAll = w => { const m = model(w); for (const pr of task.probes || []) { const a = w.evaluate(pr, si), b = m.evaluate(pr, si); if (!same(task.exact, a, b)) return `${pr} gives ${show(a)}; it should give ${show(b)}.`; } return null; };
      let bad = tryAll(wb); if (bad) return fail(mistakeFor(task, d.ref, null) || bad);
      if (ctx.inputs && ctx.inputs.length) for (let k = 1; k <= 2; k++) { const w2 = new Workbook(wb.toSpec()); randomize(w2, si, ctx.inputs, k); bad = tryAll(w2); if (bad) return fail(`It works for these numbers, but not for others: ${bad}`); }
      return pass();
    }
    case 'cf': { // conditional formatting on task.range. Visual kinds are matched by type; highlight kinds by which cells light up (task.marks: a TRUE/FALSE formula for the first cell)
      const g = parseRange(task.range); const rules = (wb.sheets[si].cf || []).map((r, k) => ({ r, k })); const covers = r => String(r.sqref).split(/\s+/).some(p => { const q = parseRange(p.replace(/\$/g, '')); return q && !(q.r2 < g.r1 || q.r1 > g.r2 || q.c2 < g.c1 || q.c1 > g.c2); });
      const here = rules.filter(x => covers(x.r)); const VIS = ['dataBar', 'colorScale', 'iconSet']; const NAMES = { dataBar: 'Data Bars', colorScale: 'a Colour Scale', iconSet: 'an Icon Set' };
      if (!here.length) return fail(task.noneMsg || `${task.range} has no conditional formatting yet. Select it, right-click › Conditional Formatting.`);
      if (VIS.includes(task.type)) { const m = here.filter(x => x.r.type === task.type); if (!m.length) return fail(task.typeMsg || `Use ${NAMES[task.type]} on ${task.range}.`);
        const exact = m.find(x => String(x.r.sqref).replace(/\$/g, '') === task.range); if (!exact) return fail(`The rule covers ${m[0].r.sqref}; it should cover exactly ${task.range}, so the cells are compared with each other and nothing else.`);
        if (task.set && exact.r.set !== task.set) return fail(task.setMsg || 'Pick the icon set the task asks for.');
        if (task.low) { const c0 = (exact.r.colors || [])[0] || ''; const red = parseInt(c0.slice(1, 3), 16) > parseInt(c0.slice(3, 5), 16); if ((task.low === 'red') !== red) return fail(task.lowMsg || `Use the scale that makes the lowest values ${task.low}.`); }
        if (task.only && here.some(x => x.r.type !== task.type)) return fail(task.onlyMsg || `Keep just the ${NAMES[task.type].replace(/^an? /, '')} on ${task.range}: clear the other rules first.`);
        return pass(); }
      const hiAll = here.filter(x => !VIS.includes(x.r.type)); if (!hiAll.length) return fail(task.typeMsg || `Add a highlight rule to ${task.range}.`);
      const hi = hiAll.filter(x => String(x.r.sqref).replace(/\$/g, '') === task.range); if (!hi.length) return fail(`Your rule is on ${hiAll[hiAll.length - 1].r.sqref}. Select exactly ${task.range} before adding the rule (clear the other one first).`);
      if (task.mustType && !hi.some(x => x.r.type === task.mustType)) return fail(task.mustTypeMsg || 'Use the rule type the task asks for.');
      let bad = null; eachCell(g, (r, c) => { if (bad) return; const f = shiftFormula(task.marks, r - g.r1, c - g.c1); const x = wb.evaluate(f, si); const want = x === true || (typeof x === 'number' && x !== 0); const got = hi.some(q => { const h = wb.cfHit(si, r, c, q.r, q.k); return !!h; }); if (want !== got) bad = { a: addr(r, c), want }; });
      if (bad) return fail((task.mistakes && mistakeFor(task, hi.map(x => x.r.formula || '').join(' '), null)) || task.wrongMsg || (bad.want ? `${bad.a} should be highlighted, but isn’t.` : `${bad.a} is highlighted, but shouldn’t be.`) + (hi.some(x => x.r.type === 'expression') ? ' Check the $ signs: write the formula for the first cell of the selection.' : ''));
      return pass(); }
    case 'sparkline': { // sparklines in task.at drawn from task.data
      const list = wb.sheets[si].sparklines || []; const ag = parseRange(task.at); const sp = list.find(x => x.at === task.at) || list.find(x => { const h = parseRange(x.at); return !(h.r2 < ag.r1 || h.r1 > ag.r2 || h.c2 < ag.c1 || h.c1 > ag.c2); });
      if (!sp) return fail(task.noneMsg || `There are no sparklines in ${task.at} yet. Select ${task.at}, right-click › Sparklines.`);
      if (sp.at !== task.at) return fail(`The sparklines are in ${sp.at}; put them in ${task.at}.`); if (sp.data !== task.data) return fail(`The Data Range is ${sp.data}; it should be ${task.data}: the numbers only, without the headings.`);
      if (task.type && !task.type.split('|').includes(sp.type)) return fail(task.typeMsg || `Use ${task.type.split('|').join(' or ')} sparklines here.`);
      for (const [k, label] of [['high', 'High Point'], ['low', 'Low Point'], ['negative', 'Negative Points'], ['markers', 'Markers'], ['sameAxis', 'Same axis for all sparklines']]) if (task[k] !== undefined && !!sp[k] !== !!task[k]) return fail(`${task[k] ? 'Tick' : 'Untick'} ${label}.`);
      return pass(); }
    case 'validation': { // a drop-down list (Data › Data Validation › List) on task.cell with these items
      const p = parseAddr(task.cell); const items = wb.listItems(si, p.r, p.c);
      if (!items) return fail(task.noneMsg || `${task.cell} has no drop-down list yet. Select it, right-click › Data Validation, choose List, and type the Source.`);
      const want = task.items || []; const low = a => a.map(x => String(x).trim().toLowerCase());
      const miss = want.filter(w => !low(items).includes(String(w).toLowerCase())); if (miss.length) return fail(`The list is missing ${miss.join(', ')}. ${task.sourceHint || 'Point the Source at the cells that hold the names.'}`);
      const extra = items.filter(x => !low(want).includes(String(x).trim().toLowerCase())); if (want.length && extra.length) return fail(`The list has extra items: ${extra.slice(0, 3).join(', ')}. ${task.sourceHint || 'Point the Source at just the cells that hold the names.'}`);
      if (task.alsoWhen) { /* the list must follow another cell: set it to something else and look again */ const w2 = new Workbook(wb.toSpec()); for (const [a, val] of Object.entries(task.alsoWhen.set || {})) { const q = parseAddr(a); w2.setInput(si, q.r, q.c, val); } const it2 = low(w2.listItems(si, p.r, p.c) || []).sort().join('|'); if (it2 !== low(task.alsoWhen.items).sort().join('|')) return fail(task.alsoWhen.msg || 'The list doesn’t change when the cell it depends on changes.'); }
      if (task.range) { const v = (wb.sheets[si].validations || []).find(d => String(d.sqref).split(/\s+/).some(q => { const g = parseRange(q.replace(/\$/g, '')); return g && p.r >= g.r1 && p.r <= g.r2 && p.c >= g.c1 && p.c <= g.c2; })); if (!v || /^"/.test(String(v.formula))) return fail(task.rangeMsg || 'Typing the names works, but point the Source at the cells instead (like =$B$1:$F$1), so the list changes when they do.'); }
      return pass(); }
    case 'table': { // an Excel Table: it exists (by name), covers task.ref, and (optionally) holds these values
      const all = wb.sheets[si].tables || []; const t = task.name ? all.find(x => x.name.toLowerCase() === task.name.toLowerCase()) : all[0];
      if (!t) return fail(all.length && task.name ? `Name the table ${task.name}: the Create Table box has a Table Name field (in Excel, Table Design › Table Name).` : 'There’s no Table on this sheet yet. Click in the data and press Ctrl+T.');
      for (const [a, want] of Object.entries(task.cells || {})) { const q = parseAddr(a); const v = wb.value(si, q.r, q.c); if (v === null) return fail(`${a} is still empty.`); if (!near(v, want)) return fail(`${a} shows ${show(v)}; it should be ${show(want)}.`); }
      if (task.ref && t.ref !== task.ref) { const g = parseRange(t.ref), w = parseRange(task.ref); return fail(task.refMsg || (g.r2 < w.r2 ? `The table still ends at row ${g.r2 + 1}. Type the new row directly below the table, in row ${g.r2 + 2}, so it grows.` : `The table covers ${t.ref}; it should cover ${task.ref}.`)); }
      return pass(); }
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
  if (task.kind === 'pq' || task.kind === 'pqload') { if (ctx.pq) pqApply(task, ctx.pq); return; }
  if (MODEL_KINDS.includes(task.kind)) { if (ctx.model) modelApply(task, ctx.model.state); return; }
  if (VBA_KINDS.includes(task.kind)) { if (ctx.vba) vbaApply(task, ctx.vba.state, ctx.wb, ctx.vba.lab); return; }
  if (OS_KINDS.includes(task.kind)) { if (ctx.os) osApply(task, ctx.os.state); return; }
  const wb = ctx.wb, grid = ctx.grid; const si = sheetIdx(wb, task.sheet);
  switch (task.kind) {
    case 'cell': case 'spill': { for (const a of task.clear || []) { const q = parseAddr(a); wb.setInput(si, q.r, q.c, ''); } const p = parseAddr(task.cell); wb.setInput(si, p.r, p.c, task.solution); break; }
    case 'range': { const g = parseRange(task.range); eachCell(g, (r, c) => wb.setInput(si, r, c, shiftFormula(task.solutionFormula || task.ref, r - g.r1, c - g.c1))); break; }
    case 'count': case 'values': for (const [a, v] of Object.entries(task.solutionValues || task.cells)) { const p = parseAddr(a); wb.setInput(si, p.r, p.c, typeof v === 'object' && v && v.date ? v.date : String(v)); } break;
    case 'format': { const g = parseRange(task.range); const st = {}; if (task.fmt) st.fmt = task.fmt[0] === 'General' ? null : task.fmt[0]; if (task.solutionFmt) st.fmt = task.solutionFmt; if (task.bold !== undefined) st.bold = task.bold || null; if (task.fill) st.fill = '#FFF2CC'; if (task.align) st.align = task.align; if (task.wrap) st.wrap = true; eachCell(g, (r, c) => wb.setStyle(si, r, c, st)); if (task.minWidth) wb.sheets[si].widths[colIndex(task.minWidth.col)] = task.minWidth.px + 10; break; }
    case 'width': wb.sheets[si].widths[colIndex(task.col)] = task.min + 10; wb.invalidate(); break;
    case 'select': if (task.sheet) grid.setSheet(sheetIdx(wb, task.sheet)); grid.select(task.range); break;
    case 'sorted': { const g = parseRange(task.region || task.range); wb.sortRegion(si, g, [{ col: colIndex(task.col), desc: !!task.desc }], !!task.region); break; }
    case 'chart': break; // the builder's settings come from task.solutionChart (see lesson.js)
    case 'cf': { const sh = wb.sheets[si]; sh.cf = (sh.cf || []).filter(r => !(task.only && String(r.sqref) === task.range)); sh.cf.push({ sqref: task.range, ...task.solutionRule }); wb.invalidate(); break; }
    case 'sparkline': { const sh = wb.sheets[si]; sh.sparklines = (sh.sparklines || []).filter(x => x.at !== task.at); const sp = { at: task.at, data: task.data, type: (task.type || 'line').split('|')[0] }; for (const k of ['high', 'low', 'negative', 'markers', 'sameAxis']) if (task[k]) sp[k] = true; sh.sparklines.push(sp); wb.invalidate(); break; }
    case 'validation': { const v = wb.sheets[si].validations || (wb.sheets[si].validations = []); for (let k = v.length - 1; k >= 0; k--) if (v[k].sqref === task.cell) v.splice(k, 1); v.push({ sqref: task.cell, type: 'list', formula: task.solutionSource || ('"' + task.items.join(',') + '"') }); wb.invalidate(); break; }
    case 'table': { let t = (wb.sheets[si].tables || []).find(x => x.name.toLowerCase() === String(task.name || '').toLowerCase()); for (const [a, v] of Object.entries(task.cells || {})) { const q = parseAddr(a); wb.setInput(si, q.r, q.c, String(v)); } t = t || (wb.sheets[si].tables || [])[0]; if (!t && task.ref) wb.addTable(si, task.ref, task.name); else if (t && task.ref) t.ref = task.ref; wb.invalidate(); break; }
    case 'lambda': { if (task.cell) { const p = parseAddr(task.cell); wb.setInput(si, p.r, p.c, task.ref); } wb.defineName(task.name, task.ref); break; }
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
    case 'pq': out.push(task.fnArgs ? `the function “${task.query}” gives the model answer’s result for ${task.fnArgs.length} test values: ${task.fnArgs.join(', ')}` : task.valueCheck ? `the query “${task.query}” gives the model answer’s value` : `the query “${task.query}” gives the same table as the model answer: the same columns${(task.compare || {}).types ? ', data types' : ''} and rows`); if (task.alsoWith && task.alsoWith.length) out.push(`and it still does on other data: ${task.alsoWith.map(a => a.say || 'changed data').join('; ')}`); if (task.mustUse) out.push(`the query uses ${task.mustUse.join(', ')}`); break;
    case 'model': case 'mpivot': case 'measure': case 'calccol': case 'kpi': out.push(...describeModel(task)); break;
    case 'vba': case 'vbafn': case 'vbacode': case 'vbaran': case 'vbacell': out.push(...describeVBA(task)); break;
    case 'os': case 'osfn': case 'oscode': out.push(...describeOS(task)); break;
    case 'pqload': out.push(`“${task.query}” is loaded ${task.to === 'connection' ? 'as a connection only' : 'to a table on a sheet'}${task.model ? ', and added to the Data Model' : ''}`); break;
    case 'cell':
      if (e.formula !== false && (e.formula || e.ref)) out.push(`${task.cell} holds a formula (it starts with =), not a typed-in number`);
      if (e.ref) out.push(`it gives the same answer as ${f(e.ref)}`); else if (e.value !== undefined) out.push(`it shows ${e.value}`);
      (e.mustUse || []).forEach(fn => out.push(`it uses ${fn}`)); (e.mustNotUse || []).forEach(fn => out.push(`it doesn’t use ${fn}`)); if (e.contains) out.push(e.containsSays || 'the text includes the picked name and the share, worked out by formulas');
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
    case 'chart': { const e = task.expect || {}; if (e.range) out.push(`the chart shows the data in ${e.range}`); if (e.type) out.push(`it is a ${e.type.split('|').join(' or ')} chart`); if (e.title) out.push('the title states the finding, not just the topic'); if (e.labels !== undefined) out.push(e.labels ? 'data labels are on' : 'data labels are off'); if (e.legend) out.push(e.legend === 'none' ? 'there is no legend' : 'there is a legend'); if (e.grid === false) out.push('gridlines are off'); if (e.yZero) out.push('the value axis starts at 0'); if (e.highlight) out.push(`${e.highlight} is the only highlighted category`); if (e.secondary) out.push(`${e.secondary} is a line on the secondary axis`); if (e.stacked) out.push(`the series are ${e.stacked === 'percent' ? '100% stacked' : 'stacked'}`); if (e.largestFirst) out.push('the biggest bar comes first'); if (e.endLabels) out.push('each line is labelled at its end'); if (e.binWidth !== undefined) out.push(`the bin width is ${e.binWidth}`); if (e.underflow !== undefined) out.push(`the underflow bin is ${e.underflow}`); if (e.overflow !== undefined) out.push(`the overflow bin is ${e.overflow}`); if (e.totals) out.push(`${e.totals.join(', ')} are set as totals`); if (e.hideFirst) out.push('the first series has no fill'); if (e.invert) out.push('negative bars have their own colour'); if (e.yMin !== undefined) out.push(`the axis starts at ${e.yMin}`); if (e.titleCell) out.push(`the title is linked to ${e.titleCell}`); break; }
    case 'lambda': out.push(`a function called ${task.name} is saved in the Name Box from a =LAMBDA(…) formula`, `it gives the right answer for ${(task.probes || []).join(', ')}`); if (inputs && inputs.length) out.push('and for 2 sets of fresh numbers'); break;
    case 'cf': out.push(task.says || (['dataBar', 'colorScale', 'iconSet'].includes(task.type) ? `${task.range} has ${{ dataBar: 'Data Bars', colorScale: 'a Colour Scale', iconSet: 'an Icon Set' }[task.type]}, on exactly that range` : `the rule on ${task.range} highlights the right cells, and only those`)); if (task.mustType === 'expression') out.push('the rule uses a formula'); break;
    case 'sparkline': out.push(`${task.at} holds sparklines of ${task.data}`); if (task.type) out.push(`they are ${task.type.split('|').join(' or ')} sparklines`); for (const [k, label] of [['high', 'the high point'], ['low', 'the low point'], ['negative', 'negative points']]) if (task[k]) out.push(`${label} is marked`); break;
    case 'validation': out.push(`${task.cell} has a drop-down list`, `the list holds ${(task.items || []).join(', ')}`); if (task.range) out.push('the list comes from cells on the sheet, not typed names'); break;
    case 'table': out.push(task.name ? `a Table called ${task.name} exists` : 'the data is an Excel Table'); if (task.cells) out.push(...Object.entries(task.cells).map(([a, v]) => `${a} holds ${v}`)); if (task.ref) out.push(`the table covers ${task.ref}`); break;
    case 'name': out.push(`a name called ${task.name} exists`, `it points to ${task.ref}${task.sheet ? ' on ' + task.sheet : ''}`); break;
    case 'sheet': out.push(`a sheet called “${task.name}” exists`); if (task.index !== undefined) out.push(`it is sheet number ${task.index + 1} from the left`); break;
    case 'noBlankRows': out.push(`there are no empty rows between the first and last record in ${task.range}`); break;
    case 'count': out.push(`${task.range} has ${task.equals} filled cells`); break;
    case 'keys': out.push(`you pressed ${task.label || task.combo}`); break;
    case 'mc': out.push('the option you chose'); break; case 'hotspot': out.push('the part of the screenshot you clicked'); break; case 'order': out.push('the order of the steps'); break;
  }
  if (task.alsoWith) { const cells = [...new Set(task.alsoWith.flatMap(o => Object.keys(o)))]; out.push(`it still gives the right answer when ${cells.join(' or ')} change${cells.length === 1 ? 's' : ''} (Sheetah tries ${task.alsoWith.length} other pick${task.alsoWith.length === 1 ? '' : 's'})`); }
  return out;
}
