// Study tools adapted from the CAT app: the mistakes log with causes, the study plan with missed-day reasons,
// journal check-ins and weekly reviews, and insights. Empty pages explain what they're for, with an example.
import { register, getCourse, go } from '../app.js';
import { store, today } from '../store.js';
import { esc, md, $, $$, mascot, toast, heatmap, bars, lineChart, download, fmtDate, fmtDay, dateField, readDate, qhelp } from '../ui.js';
import { dayNum, dayStr, recallIn } from '../srs.js';
import { lessonList, isDone, canOpenLesson, lessonLabel } from '../access.js';
import { t } from '../i18n.js';
import { buildPlan, planStatus, setBaseline, addDays, fmtD, iso, dayMinutes } from '../plan.js';
const t_ = t;

const CAUSES = ['Didn’t know it yet', 'Careless slip', 'Wrong reference', 'Misread the task', 'Forgot the steps', 'Other'];
const lessonName = (c, id) => { const l = lessonList(c).find(x => x.id === id); return l ? `${lessonLabel(l)} · ${l.title}` : ''; };
const example = (title, html) => `<div class="example"><span class="tag">Example</span><p class="small muted">${title}</p>${html}</div>`;
register('errors', async (main, filter = 'open') => {
  const c = await getCourse();
  const list = store.d.errors.filter(e => filter === 'all' || (filter === 'open' ? !e.resolved : e.resolved));
  const byCause = {}; store.d.errors.forEach(e => { const k = e.cause || 'Not tagged'; byCause[k] = (byCause[k] || 0) + 1; });
  const byLesson = {}; store.d.errors.forEach(e => { if (e.lesson) byLesson[e.lesson] = (byLesson[e.lesson] || 0) + 1; });
  const open = store.d.errors.filter(e => !e.resolved).length, fixed = store.d.errors.filter(e => e.resolved).length;
  if (!store.d.errors.length) {
    main.innerHTML = `<h1>${t('Mistakes')}</h1><section class="card empty">${mascot('encourage')}<div><h2>Nothing here yet, and that’s fine</h2><p>Every wrong lab or quiz answer lands here, next to the right one. Each comes back in your daily review after 1, 3 and 7 days until it sticks, and tagging a cause shows your patterns.</p>
      ${example('A mistake looks like this:', `<div class="err-item"><div class="em"><span class="tag">Module 3 · Lesson 2 · lab</span></div><p>Lock the price in B1 so the formula can be filled down.</p><div class="cmp"><div><small>You</small><code>=A4*B1</code></div><div><small>Right</small><code>=A4*$B$1</code></div></div><p class="muted small">B1 moved when you filled down. Press F4 to make it $B$1.</p></div>`)}<a class="btn" href="#/home">Back to studying</a></div></section>`; return;
  }
  main.innerHTML = `<h1>${t('Mistakes')}</h1><p class="lead">Every wrong lab and quiz answer lands here, beside the right one. Each comes back in your daily review after 1, 3 and 7 days until it sticks.</p>
  <div class="grid3"><div class="card stat"><b>${open}</b><span>open</span></div><div class="card stat"><b>${fixed}</b><span>fixed for good</span></div><div class="card stat"><b>${store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length}</b><span>due today</span></div></div>
  <div class="grid2"><section class="card"><h2>Causes so far</h2>${Object.keys(byCause).some(k => k !== 'Not tagged') ? bars(Object.entries(byCause).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 150, label: 'Mistakes by cause' }) + `<p class="muted small">${Object.entries(byCause).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ')}</p>` : '<p class="muted">Pick “Why did it happen?” on a mistake below. After a few, a chart here shows your patterns: careless slips need a slower check; “didn’t know it yet” needs the lesson again.</p>'}</section>
  <section class="card"><h2>By lesson</h2><ul class="plist">${Object.entries(byLesson).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([id, n]) => `<li><span>${esc(lessonName(c, id) || id)}</span><small>${n} mistake${n === 1 ? '' : 's'} · <a href="#/practice/c1/${id}">practise</a></small></li>`).join('')}</ul></section></div>
  <section class="card"><div class="row between"><h2>All mistakes</h2><div class="seg">${[['open', 'Open'], ['resolved', 'Fixed'], ['all', 'All']].map(([f, l]) => `<a class="segb ${f === filter ? 'on' : ''}" href="#/errors/${f}">${l}</a>`).join('')}</div></div>
  ${list.length ? `<ul class="errs">${list.map(e => `<li class="err-item ${e.resolved ? 'res' : ''}"><div class="em"><span class="tag">${esc(lessonName(c, e.lesson) || '')} · ${{ lab: 'lab', quiz: 'quiz', check: 'lecture check', test: 'test' }[e.kind] || esc(e.kind)}</span><span class="muted small">${fmtDate(e.at)}</span></div><p>${md(e.prompt)}</p><div class="cmp"><div><small>You</small><code>${esc(e.yours)}</code></div><div><small>Right</small><code>${esc(e.right)}</code></div></div>${e.why ? `<p class="muted small">${md(e.why)}</p>` : ''}
    <div class="row"><select data-id="${e.id}" class="cause" aria-label="Why did it happen?"><option value="">Why did it happen?</option>${CAUSES.map(cc => `<option ${e.cause === cc ? 'selected' : ''}>${cc}</option>`).join('')}</select>${e.lesson ? `<a class="btn ghost sm" href="#/practice/${e.course || 'c1'}/${e.lesson}">Practise this lesson</a>` : ''}${e.resolved ? '' : `<button class="btn ghost sm" data-fix="${e.id}" type="button">Mark fixed</button>`}</div></li>`).join('')}</ul>` : `<div class="empty-in">${mascot('encourage', 'sm')}<p>${filter === 'open' ? 'No open mistakes. Lovely.' : 'Nothing here yet.'}</p></div>`}</section>`;
  $$('.cause').forEach(s => s.onchange = () => { const e = store.d.errors.find(x => x.id === s.dataset.id); e.cause = s.value; store.save(); toast('Saved'); });
  $$('[data-fix]').forEach(b => b.onclick = () => { const e = store.d.errors.find(x => x.id === b.dataset.fix); e.resolved = true; e.resolvedAt = new Date().toISOString(); store.save(); toast('Marked as fixed'); b.closest('li').classList.add('res'); b.remove(); });
});

// the study plan: pick dates, and the plan builds itself (see plan.js)
register('planner', async (main, view) => {
  const c = await getCourse(); const p = store.d.plan; p.off = p.off || [];
  if (!p.start) p.start = today();
  const D = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']; const order = [1, 2, 3, 4, 5, 6, 0];
  const P = buildPlan(c); const S = planStatus(c); const t = today();
  const months = []; if (P.dates.length || P.items.length) { let m = (P.start || t).slice(0, 7); const last = (P.last || P.end || t).slice(0, 7); for (let g = 0; m <= last && g < 13; g++) { months.push(m); const [y, mm] = m.split('-').map(Number); m = mm === 12 ? `${y + 1}-01` : `${y}-${String(mm + 1).padStart(2, '0')}`; } }
  let mi = Math.max(0, months.indexOf(t.slice(0, 7))); if (view && months.includes(view)) mi = months.indexOf(view);
  const kinds = { lesson: 0, test: 0, project: 0, exam: 0 }; P.items.forEach(x => kinds[x.kind]++);
  const load = P.perDay; const heavy = load > 120, light = load && load < 15;
  const later = heavy && P.dates.length ? (() => { const need = Math.ceil(P.total / 60); let d = P.start, k = 0, g = 0; while (k < need && g++ < 1500) { if (p.days.includes(new Date(d + 'T00:00:00').getDay()) && !p.off.includes(d)) k++; d = addDays(d, 1); } return addDays(d, -1); })() : null;
  const statusHTML = !P.items.length ? '<span class="pstat ok">🎉 Everything is done</span>' : P.error ? `<span class="pstat bad">⚠ ${esc(P.error)}</span>` : S && S.behind ? `<span class="pstat warn">⏳ ${S.behind} item${S.behind === 1 ? '' : 's'} behind: the plan has moved them to the days ahead</span>` : S && S.ahead ? `<span class="pstat ok">🚀 ${S.ahead} item${S.ahead === 1 ? '' : 's'} ahead of plan</span>` : '<span class="pstat ok">✓ On track</span>';
  const chip = x => `<a class="pchip k-${x.kind}" href="${x.href}" title="${esc(x.title)} · about ${x.min} min">${x.kind === 'lesson' ? `<b>${esc(x.short)}</b> ${esc(x.title)}` : `<b>${esc(x.short)}</b>`}</a>`;
  const month = m => { const [y, mm] = m.split('-').map(Number); const first = new Date(y, mm - 1, 1); const pad = (first.getDay() + 6) % 7; const n = new Date(y, mm, 0).getDate();
    let cells = '<div class="cal-h">' + order.map(k => `<span>${D[k]}</span>`).join('') + '</div><div class="cal-g">' + '<span class="cal-d empty"></span>'.repeat(pad);
    for (let d = 1; d <= n; d++) { const s = `${m}-${String(d).padStart(2, '0')}`; const items = P.days[s] || []; const inRange = s >= P.start && s <= (P.end || s); const study = p.days.includes(new Date(s + 'T00:00:00').getDay()); const off = p.off.includes(s);
      const cls = ['cal-d', s === t ? 'today' : '', s < t ? 'past' : '', off ? 'off' : '', !study ? 'rest' : '', inRange && study && !off && s >= t ? 'can' : '', items.length ? 'has' : '', s === P.end ? 'end' : ''].filter(Boolean).join(' ');
      cells += `<div class="${cls}" data-day="${s}" ${inRange && study && s >= t ? `title="${off ? 'Day off: click to study this day' : 'Click to take this day off'}"` : ''}><span class="cal-n">${d}${s === P.end ? ' 🏁' : ''}</span>${off ? '<span class="cal-off">Day off</span>' : ''}${items.slice(0, 3).map(chip).join('')}${items.length > 3 ? `<span class="pmore">+${items.length - 3} more</span>` : ''}${items.length ? `<span class="cal-min">${dayMinutes(items)} min</span>` : ''}</div>`; }
    return cells + '</div>'; };
  const wk = []; for (let k = 0; k < 7; k++) { const s = addDays(t, k); if (P.days[s]) wk.push([s, P.days[s]]); }
  main.innerHTML = `<h1>${t_('Study plan')}</h1><p class="lead">Pick your dates. Sheetah spreads every lesson, test, the project and the exam across them, and moves things along by itself if you miss a day.</p>
  <section class="card plan-setup"><div class="ps-grid">
    <div class="ps-step"><span class="ps-n">1</span><div><b>Start</b>${dateField('start', p.start, { label: 'Start date' })}<button class="linkbtn small" type="button" id="stToday">Today</button></div></div>
    <div class="ps-step"><span class="ps-n">2</span><div><b>Finish by</b>${dateField('end', p.end || '', { label: 'Finish date' })}<div class="chips quick">${[[1, '1 month'], [2, '2 months'], [3, '3 months'], [6, '6 months']].map(([n, l]) => `<button type="button" class="chip" data-months="${n}">${l}</button>`).join('')}</div></div></div>
    <div class="ps-step"><span class="ps-n">3</span><div><b>Study days</b><div class="days">${order.map(k => `<label class="pill"><input type="checkbox" data-wd="${k}" ${p.days.includes(k) ? 'checked' : ''}> ${D[k]}</label>`).join('')}</div><label class="inline small">Time <input type="time" id="ptime" value="${esc(p.time)}"></label></div></div>
    <div class="ps-step"><span class="ps-n">4</span><div><b>Days off</b><p class="small muted">Holidays, exams, a wedding? Click those dates on the calendar below.</p>${p.off.length ? `<p class="small">${p.off.length} day${p.off.length === 1 ? '' : 's'} off · <button class="linkbtn" type="button" id="clrOff">clear them</button></p>` : ''}</div></div>
  </div></section>
  <section class="card plan-sum"><div class="d-big">
    <div class="bn"><b>${P.items.length}</b><span>things left</span><small>${kinds.lesson} lessons · ${kinds.test} tests${kinds.project ? ' · project' : ''}${kinds.exam ? ' · exam' : ''}</small></div>
    <div class="bn"><b>${P.dates.length}</b><span>study days</span><small>${P.dates.length ? `${fmtD(P.dates[0])} → ${fmtD(P.dates.at(-1), { day: 'numeric', month: 'short', year: 'numeric' })}` : 'pick your dates'}</small></div>
    <div class="bn"><b>${load ? `≈ ${load}` : '—'}</b><span>minutes a day</span><small>${Math.round(P.total / 6) / 10} hours in all</small></div></div>
    <div class="row between ps-status">${statusHTML}${p.end ? '' : '<span class="small muted">No finish date yet, so this uses ' + p.minutesGoal + ' minutes a day.</span>'}</div>
    ${heavy ? `<p class="warnbox small">That's over two hours a day. ${later ? `For about an hour a day, finish around <b>${fmtD(later, { day: 'numeric', month: 'short', year: 'numeric' })}</b>. <button class="btn sm" type="button" id="useLater">Use that date</button>` : ''}</p>` : light ? '<p class="small muted">Under 15 minutes a day: you could finish sooner.</p>' : ''}</section>
  ${P.items.length && !P.error ? `<section class="card"><div class="row between"><h2>This week</h2><a class="btn ghost sm" href="#/home">Today's list on Home</a></div>${wk.length ? `<div class="week">${wk.map(([s, items]) => `<div class="wday ${s === t ? 'today' : ''}"><span class="wd">${s === t ? 'Today' : fmtD(s, { weekday: 'short', day: 'numeric', month: 'short' })}</span><div class="wl">${items.map(chip).join('')}</div><span class="cal-min">${dayMinutes(items)} min</span></div>`).join('')}</div>` : '<p class="muted">Nothing planned in the next 7 days.</p>'}</section>
  <section class="card cal"><div class="row between"><button class="btn ghost sm" type="button" id="mPrev" ${mi <= 0 ? 'disabled' : ''}>←</button><h2>${months.length ? new Date(months[mi] + '-01T00:00:00').toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }) : ''}</h2><button class="btn ghost sm" type="button" id="mNext" ${mi >= months.length - 1 ? 'disabled' : ''}>→</button></div>
    ${months.length ? month(months[mi]) : ''}<div class="legend cal-leg"><span class="pchip k-lesson"><b>1.2</b> Lesson</span><span class="pchip k-test"><b>Test</b></span><span class="pchip k-project"><b>Project</b></span><span class="pchip k-exam"><b>Exam</b></span><span class="cal-sw off"></span>Day off<span class="cal-sw rest"></span>Not a study day</div></section>` : ''}
  <div class="row"><button type="button" class="btn ghost" id="ics">📅 Add to my calendar</button><button type="button" class="btn ghost" id="notif">🔔 Turn on reminders</button></div>
  <details class="card habits"><summary><b>Your study habits</b> <span class="muted small">missed days and when you really study</span></summary><div class="grid2">${habits()}</div></details>`;
  const save = (rebase = true) => { if (rebase) setBaseline(c); else store.save(); go('#/planner' + (view ? '/' + view : '')); };
  const readEnd = () => readDate(main, 'end'), readStart = () => readDate(main, 'start');
  $$('[data-date="start"] select').forEach(x => x.onchange = () => { const v = readStart(); if (v) { p.start = v; save(); } });
  $$('[data-date="end"] select').forEach(x => x.onchange = () => { const v = readEnd(); if (v) { p.end = v; save(); } });
  $('#stToday').onclick = () => { p.start = today(); save(); };
  $$('[data-months]').forEach(b => b.onclick = () => { const d = new Date((p.start > t ? p.start : t) + 'T00:00:00'); d.setMonth(d.getMonth() + +b.dataset.months); p.end = iso(d); save(); });
  $$('[data-wd]').forEach(x => x.onchange = () => { p.days = $$('[data-wd]').filter(y => y.checked).map(y => +y.dataset.wd).sort(); save(); });
  $('#ptime').onchange = e => { p.time = e.target.value; store.save(); toast('Saved'); };
  const co = $('#clrOff'); if (co) co.onclick = () => { p.off = []; save(); };
  const ul = $('#useLater'); if (ul) ul.onclick = () => { p.end = later; save(); };
  $$('.cal-d.can, .cal-d.off').forEach(d => d.addEventListener('click', e => { if (e.target.closest('a')) return; const s = d.dataset.day; p.off = p.off.includes(s) ? p.off.filter(x => x !== s) : p.off.concat(s); save(false); }));
  const mp = $('#mPrev'), mn = $('#mNext'); if (mp) mp.onclick = () => go('#/planner/' + months[mi - 1]); if (mn) mn.onclick = () => go('#/planner/' + months[mi + 1]);
  $('#ics').onclick = () => { const f = s => s.replace(/-/g, ''); const [hh, mm] = (p.time || '20:00').split(':');
    const ev = Object.entries(P.days).map(([d, items], k) => `BEGIN:VEVENT\r\nUID:sheetah-${d}-${k}@sheetah\r\nDTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').slice(0, 15)}Z\r\nDTSTART:${f(d)}T${hh}${mm}00\r\nDURATION:PT${dayMinutes(items)}M\r\nSUMMARY:Sheetah: ${items.map(x => x.kind === 'lesson' ? 'Lesson ' + x.short : x.short).join(', ')}\r\nDESCRIPTION:${items.map(x => x.title).join(' · ')}\r\nBEGIN:VALARM\r\nTRIGGER:-PT10M\r\nACTION:DISPLAY\r\nDESCRIPTION:Excel study in 10 minutes\r\nEND:VALARM\r\nEND:VEVENT\r\n`).join('');
    download('sheetah-study-plan.ics', `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Sheetah//Study plan//EN\r\n${ev}END:VCALENDAR\r\n`, 'text/calendar'); toast('Open the file to add every study session to your calendar.'); };
  $('#notif').onclick = async () => { if (!('Notification' in window)) return toast('This browser has no notifications; Sheetah will remind you inside the app.'); const r = await Notification.requestPermission(); store.d.settings.remind = true; store.save(); toast(r === 'granted' ? 'Reminders on. You’ll get one at your study time while Sheetah is open.' : 'Notifications are blocked; Sheetah will remind you inside the app instead.'); };
  function habits() {
    const missed = Object.entries(p.missed).sort().reverse(); const reasons = {}; missed.forEach(([, r]) => reasons[r] = (reasons[r] || 0) + 1);
    const w = [0, 0, 0, 0, 0, 0, 0]; Object.entries(store.d.activity).forEach(([d, a]) => { w[new Date(d + 'T00:00:00').getDay()] += a.minutes || 0; });
    return `<div><h3>Missed days and why</h3>${missed.length ? `<ul class="missed-list">${missed.slice(0, 10).map(([d, r]) => `<li><span>${fmtDate(d)}</span><b>${esc(r)}</b></li>`).join('')}</ul>${bars(Object.entries(reasons).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 120, label: 'Reasons for missed days' })}` : '<p class="muted small">When you miss a planned day, Home asks what got in the way. Patterns show up here.</p>'}</div>
      <div><h3>When you actually study</h3>${Object.keys(store.d.activity).length ? bars(order.map(k => ({ k: D[k], v: Math.round(w[k]), hl: p.days.includes(k) })), { h: 130, fmt: v => v + 'm', label: 'Minutes by weekday' }) + '<p class="muted small">Highlighted bars are your study days.</p>' : '<p class="muted small">After a week or two, this compares your planned days with when you really study.</p>'}</div>`;
  }
});

// weeks are keyed Monday-first; the heading reads “This week: 2–8 Nov”
const weekKey = d => { const m = new Date(d); m.setHours(0, 0, 0, 0); m.setDate(m.getDate() - ((m.getDay() + 6) % 7)); return m.toISOString().slice(0, 10); };
const weekLabel = k => { const a = new Date(k + 'T00:00:00'), b = new Date(a); b.setDate(a.getDate() + 6); const sameM = a.getMonth() === b.getMonth(); return `${a.getDate()}${sameM ? '' : ' ' + a.toLocaleString('en-IN', { month: 'short' })}–${b.getDate()} ${b.toLocaleString('en-IN', { month: 'short' })}`; };
const scale = (name, val, labels) => `<div class="scale" role="radiogroup" aria-label="${esc(name)}">${labels.map((l, i) => `<label class="pill"><input type="radio" name="${name}" value="${i + 1}" ${+val === i + 1 ? 'checked' : ''}> ${l}</label>`).join('')}</div>`;
register('journal', async main => {
  const tdy = today(); const j = store.d.journal[tdy] || {}; const wk = weekKey(new Date()); const w = store.d.weekly[wk] || {};
  const past = Object.entries(store.d.journal).filter(([d]) => d !== tdy).sort().reverse();
  const ago = [7, 30].map(n => [n, store.d.journal[dayStr(dayNum() - n)]]).filter(([, e]) => e);
  const a = store.d.activity[tdy] || {};
  main.innerHTML = `<h1>${t('Journal')}</h1><p class="lead">Two minutes a day: what you learned, what was hard, how sure you feel. Writing it down helps it stick, and the patterns show you what to review.</p>
  <div class="grid2"><section class="card"><h2>Today · ${fmtDay(tdy)}</h2><p class="muted small">So far today: ${Math.round(a.minutes || 0)} minutes, ${a.lessons || 0} lesson${a.lessons === 1 ? '' : 's'}, ${a.reviews || 0} review${a.reviews === 1 ? '' : 's'}.</p><form id="jf" class="form">
   <label>What did I learn? <textarea name="learned" rows="3" placeholder="e.g. $B$1 stays put when I copy a formula">${esc(j.learned || '')}</textarea></label>
   <label>What was hard? <textarea name="hard" rows="2" placeholder="e.g. remembering when to use A$1 instead of $A$1">${esc(j.hard || '')}</textarea></label>
   <fieldset><legend>How sure do I feel about today’s topic?</legend>${scale('conf', j.conf, ['Lost', 'Shaky', 'Okay', 'Good', 'Very sure'])}</fieldset>
   <fieldset><legend>Energy</legend>${scale('energy', j.energy, ['Drained', 'Low', 'Okay', 'Good', 'Fresh'])}</fieldset>
   <button class="btn">Save check-in</button></form></section>
  <section class="card"><h2>This week: ${weekLabel(wk)}</h2><form id="wf" class="form"><label>What went well this week? <textarea name="good" rows="2">${esc(w.good || '')}</textarea></label><label>What was tough? <textarea name="bad" rows="2">${esc(w.bad || '')}</textarea></label><label>One thing I’ll change next week <textarea name="change" rows="2">${esc(w.change || '')}</textarea></label>
   <fieldset><legend>How was the week overall?</legend>${scale('rate', w.rate, ['Rough', 'Meh', 'Okay', 'Good', 'Great'])}</fieldset><button class="btn">Save weekly review</button></form>
   ${ago.length ? `<h3>On this day</h3>${ago.map(([n, e]) => `<div class="onthisday"><small>${n} days ago</small><p>${esc(e.learned || '—')}</p></div>`).join('')}` : ''}</section></div>
  <section class="card"><h2>All entries</h2>${past.length ? `<ul class="journal-list">${past.slice(0, 60).map(([d, e]) => `<li><span class="jd">${fmtDate(d)}</span><div><p>${esc(e.learned || '')}</p>${e.hard ? `<p class="muted small">Hard: ${esc(e.hard)}</p>` : ''}</div><span class="conf" title="How sure: ${e.conf || '—'} of 5">${'●'.repeat(e.conf || 0)}${'○'.repeat(5 - (e.conf || 0))}</span></li>`).join('')}</ul>` : example('Past check-ins appear here, like this:', `<ul class="journal-list"><li><span class="jd">3 Nov 2026</span><div><p>F4 cycles $A$1, A$1, $A1 and A1.</p><p class="muted small">Hard: knowing which one I need.</p></div><span class="conf">●●●○○</span></li></ul>`)}</section>`;
  $('#jf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); store.d.journal[tdy] = { learned: f.get('learned'), hard: f.get('hard'), conf: +f.get('conf') || 0, energy: +f.get('energy') || 0, at: new Date().toISOString() }; const ck = store.d.checklist[tdy] || (store.d.checklist[tdy] = {}); ck.journal = true; store.save(); toast('Check-in saved'); };
  $('#wf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); store.d.weekly[wk] = { good: f.get('good'), bad: f.get('bad'), change: f.get('change'), rate: +f.get('rate') || 0 }; store.save(); toast('Weekly review saved'); };
});

export function mastery(lessonId) {
  const l = store.d.lessons[lessonId]; if (!l || l.status !== 'done') return l && l.status === 'started' ? 1 : 0;
  const cards = Object.values(store.d.cards).filter(c => c.lesson === lessonId && !c.isNew); if (!cards.length) return 1;
  const r = cards.reduce((a, c) => a + recallIn(c, 30), 0) / cards.length; const first = l.firstTries || 0;
  return r >= .9 && first >= 3 ? 4 : r >= .8 ? 3 : r >= .6 ? 2 : 1;
}
const LEVELS = ['New', 'Learning', 'Familiar', 'Proficient', 'Mastered'];
register('insights', async main => {
  const c = await getCourse(); const d = store.d; const t0 = dayNum();
  const hasData = Object.keys(d.activity).length > 0;
  const last30 = Array.from({ length: 30 }, (_, k) => { const ds = dayStr(t0 - 29 + k); const a = d.activity[ds] || {}; return { k: ds.slice(8), v: Math.round(a.minutes || 0), r: a.reviews || 0 }; });
  const totalMin = Object.values(d.activity).reduce((a, x) => a + (x.minutes || 0), 0);
  const tests = d.tests.filter(x => x.kind === 'mock' || x.kind === 'module').slice(-15).map(x => ({ k: fmtDate(x.at), v: x.pct }));
  const causes = {}; d.errors.forEach(e => { const k = e.cause || 'Not tagged'; causes[k] = (causes[k] || 0) + 1; });
  const confs = Object.entries(d.journal).sort().slice(-30).filter(([, e]) => e.conf).map(([k, e]) => ({ k, v: e.conf * 20 }));
  const all = lessonList(c); const weak = all.filter(l => { const lv = mastery(l.id); return lv >= 1 && lv <= 2 && isDone(l.id); }).slice(0, 4);
  main.innerHTML = `<h1>${t('Insights')}</h1>${hasData ? '' : `<section class="card empty">${mascot('teaching', 'sm')}<div><h2>Your charts fill in as you study</h2><p>Minutes per day, review counts, test scores, how sure you feel, and a mastery map of every lesson. Finish your first lesson and come back: the mastery map below already shows where you’re heading.</p></div></section>`}
  <div class="grid3"><div class="card stat"><b>${Math.round(totalMin / 60 * 10) / 10}</b><span>hours studied</span></div><div class="card stat"><b>${Object.values(d.lessons).filter(l => l.status === 'done').length}</b><span>lessons done</span></div><div class="card stat"><b>${Object.values(d.cards).reduce((a, x) => a + (x.reps || 0), 0)}</b><span>cards reviewed</span></div></div>
  <section class="card"><h2>Mastery map ${qhelp('mastery')}</h2><p class="muted small">Each lesson brightens as you review its cards and dims when they’re due again. Mastered means at least a 90% chance you’d remember it in 30 days, and three lab tasks right first time.</p><div class="mastery">${c.modules.map(m => `<div class="mrow"><span class="mlabel">Module ${m.n}</span>${m.lessons.map(l => { const lv = mastery(l.id); const open = canOpenLesson(c, { ...l, module: m }); return `<a class="mcell lv${lv}" href="${open ? (lv && lv <= 2 ? `#/practice/c1/${l.id}` : `#/lesson/c1/${l.id}`) : `#/module/c1/${m.n}`}" title="Lesson ${l.num.split('.')[1]}: ${esc(l.title)} · ${LEVELS[lv]}">${l.num.split('.')[1]}</a>`; }).join('')}</div>`).join('')}</div><div class="legend">${LEVELS.map((n, i) => `<span class="mcell lv${i} sm"></span>${n}`).join(' ')}</div>
   ${weak.length ? `<p class="small"><b>Worth practising:</b> ${weak.map(l => `<a href="#/practice/c1/${l.id}">${esc(lessonLabel(l))} · ${esc(l.title)}</a>`).join(' · ')}</p>` : ''}</section>
  ${hasData ? `<section class="card"><h2>Calendar</h2>${heatmap(d.activity, 26)}</section>
  <div class="grid2"><section class="card"><h2>Minutes · last 30 days</h2>${bars(last30.map(x => ({ k: x.k, v: x.v })), { h: 160, label: 'Minutes per day' })}</section><section class="card"><h2>Reviews · last 30 days</h2>${bars(last30.map(x => ({ k: x.k, v: x.r })), { h: 160, label: 'Reviews per day' })}</section></div>
  <div class="grid2"><section class="card"><h2>Test scores</h2>${tests.length ? lineChart(tests, { target: 75, label: 'Test scores' }) : '<p class="muted">Module and practice test scores appear here as a line, with the 75% target marked.</p>'}</section><section class="card"><h2>How sure you feel</h2>${confs.length ? lineChart(confs, { label: 'Confidence', max: 100 }) : '<p class="muted">From your <a href="#/journal">journal</a> check-ins. Two minutes after a lesson is enough.</p>'}</section></div>
  <section class="card"><h2>Mistakes by cause</h2>${Object.keys(causes).length ? bars(Object.entries(causes).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 140, label: 'Mistakes by cause' }) : '<p class="muted">No mistakes logged yet.</p>'}</section>` : ''}`;
});
