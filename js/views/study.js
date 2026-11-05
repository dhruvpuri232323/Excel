// Study tools adapted from the CAT app: the mistakes log with causes, the study plan with missed-day reasons,
// journal check-ins and weekly reviews, and insights. Empty pages explain what they're for, with an example.
import { register, getCourse, go } from '../app.js';
import { store, today } from '../store.js';
import { esc, md, $, $$, mascot, toast, heatmap, bars, lineChart, download, fmtDate, fmtDay, dateField, readDate, qhelp } from '../ui.js';
import { dayNum, dayStr, recallIn } from '../srs.js';
import { lessonList, isDone, canOpenLesson, lessonLabel } from '../access.js';
import { t } from '../i18n.js';

const CAUSES = ['Didn’t know it yet', 'Careless slip', 'Wrong reference', 'Misread the task', 'Forgot the steps', 'Other'];
const lessonName = (c, id) => { const l = lessonList(c).find(x => x.id === id); return l ? `${lessonLabel(l)} · ${l.title}` : ''; };
const example = (title, html) => `<div class="example"><span class="tag">Example</span><p class="small muted">${title}</p>${html}</div>`;
register('errors', async (main, filter = 'open') => {
  const c = await getCourse('c1');
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

register('planner', async main => {
  const c = await getCourse('c1'); const p = store.d.plan; const left = lessonList(c).filter(l => !isDone(l.id));
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  // the study days + daily minutes laid out over the remaining lessons (2 lessons on a 60+ minute day)
  const sched = []; if (p.days.length) { let d = new Date(); d.setHours(0, 0, 0, 0); const perDay = p.minutesGoal >= 60 ? 2 : 1; let k = 0; let guard = 0; while (k < left.length && guard++ < 800) { if (p.days.includes(d.getDay())) { sched.push({ date: new Date(d), lessons: left.slice(k, k + perDay) }); k += perDay; } d.setDate(d.getDate() + 1); } }
  const finish = sched.length ? sched[sched.length - 1].date : null; const target = p.target ? new Date(p.target + 'T00:00:00') : null;
  const onTrack = finish && target ? finish <= target : null;
  const missed = Object.entries(p.missed).sort().reverse();
  const reasons = {}; missed.forEach(([, r]) => reasons[r] = (reasons[r] || 0) + 1);
  const wk = [0, 0, 0, 0, 0, 0, 0]; Object.entries(store.d.activity).forEach(([d, a]) => { const wd = new Date(d + 'T00:00:00').getDay(); wk[wd] += a.minutes || 0; });
  const fd = d => d.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  main.innerHTML = `<h1>${t('Study plan')}</h1><p class="lead">Deciding exactly when you’ll study makes you far more likely to do it. Pick your days and time; Sheetah lays out every remaining lesson.</p>
  <section class="card"><form id="pf" class="form"><fieldset><legend>Study days</legend><div class="days">${days.map((d, i) => `<label class="pill"><input type="checkbox" name="d" value="${i}" ${p.days.includes(i) ? 'checked' : ''}> ${d}</label>`).join('')}</div></fieldset>
   <div class="grid3"><label>Time <input type="time" name="time" value="${p.time}"></label><label>Minutes a day <select name="goal">${[15, 30, 45, 60, 90].map(m => `<option ${m === p.minutesGoal ? 'selected' : ''} value="${m}">${m} minutes</option>`).join('')}</select></label><label>Finish ${esc(c.title)} by (optional) ${dateField('target', p.target || '', { label: 'Target date' })}</label></div>
   <div class="row"><button class="btn">Save plan</button><button type="button" class="btn ghost" id="ics">Add to my calendar (.ics)</button><button type="button" class="btn ghost" id="notif">Turn on reminders</button></div></form>
   ${finish ? `<p class="plan-sum ${onTrack === false ? 'warn' : ''}">At this pace you finish the ${left.length} remaining lesson${left.length === 1 ? '' : 's'} on <b>${fd(finish)}</b>${target ? (onTrack ? ', before your target. 👍' : `, after your target of ${fmtDate(p.target)}. Add a day or raise the minutes.`) : '.'} Then the course project and the final exam.</p>` : left.length ? '<p class="warn">Pick at least one study day.</p>' : '<p>Every lesson is done. Time for the project and the final exam.</p>'}</section>
  <section class="card"><h2>Day by day</h2>${sched.length ? `<table class="tbl plan"><thead><tr><th>Date</th><th>Lesson</th><th></th></tr></thead><tbody>${sched.slice(0, 40).map(s => `<tr><td>${fmtDay(s.date.toISOString().slice(0, 10))}</td><td>${s.lessons.map(l => `${esc(lessonLabel(l))} · ${esc(l.title)}`).join('<br>')}</td><td>${s.lessons.map(l => canOpenLesson(c, l) ? `<a href="#/lesson/c1/${l.id}">Open</a>` : '<span title="Opens when you pass the module test before it">🔒</span>').join('<br>')}</td></tr>`).join('')}</tbody></table>` : '<p class="muted">Nothing left to schedule.</p>'}</section>
  <div class="grid2"><section class="card"><h2>Missed days and why</h2>${missed.length ? `<ul class="missed-list">${missed.slice(0, 12).map(([d, r]) => `<li><span>${fmtDate(d)}</span><b>${esc(r)}</b></li>`).join('')}</ul><h3>Reasons you noted</h3>${bars(Object.entries(reasons).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 130, label: 'Reasons for missed days' })}` : `<p class="muted">When you miss a planned day, Home asks what got in the way (“work ran late”, “too tired”…). Patterns show up here, so you can move a study day to when it actually works.</p>`}</section>
  <section class="card"><h2>When you actually study</h2>${Object.keys(store.d.activity).length ? bars(days.map((d, i) => ({ k: d, v: Math.round(wk[i]), hl: p.days.includes(i) })), { h: 150, fmt: v => v + 'm', label: 'Minutes by weekday' }) + '<p class="muted small">Highlighted bars are your planned days.</p>' : '<p class="muted">After a week or two, this shows the minutes you study on each weekday next to the days you planned.</p>'}</section></div>`;
  $('#pf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); p.days = f.getAll('d').map(Number); p.time = f.get('time'); p.minutesGoal = +f.get('goal'); p.target = readDate(e.target, 'target'); store.save(); toast('Plan saved'); go('#/planner'); };
  $('#ics').onclick = () => { const [h, m] = p.time.split(':'); const byday = p.days.map(i => ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'][i]).join(','); const st = new Date(); st.setHours(+h, +m, 0, 0); const f = d => d.toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z'; const en = new Date(st.getTime() + p.minutesGoal * 6e4);
    download('sheetah-study-plan.ics', `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Sheetah//Study plan//EN\r\nBEGIN:VEVENT\r\nUID:sheetah-${Date.now()}@sheetah\r\nDTSTAMP:${f(new Date())}\r\nDTSTART:${f(st)}\r\nDTEND:${f(en)}\r\nRRULE:FREQ=WEEKLY;BYDAY=${byday}\r\nSUMMARY:Sheetah: Excel study session\r\nDESCRIPTION:One lesson plus the daily review.\r\nBEGIN:VALARM\r\nTRIGGER:-PT10M\r\nACTION:DISPLAY\r\nDESCRIPTION:Excel study in 10 minutes\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`, 'text/calendar'); };
  $('#notif').onclick = async () => { if (!('Notification' in window)) return toast('This browser has no notifications; Sheetah will remind you inside the app.'); const r = await Notification.requestPermission(); store.d.settings.remind = true; store.save(); toast(r === 'granted' ? 'Reminders on. You’ll get one at your study time while Sheetah is open.' : 'Notifications are blocked; Sheetah will remind you inside the app instead.'); };
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
  const c = await getCourse('c1'); const d = store.d; const t0 = dayNum();
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
