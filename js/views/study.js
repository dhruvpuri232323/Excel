// Study tools adapted from the CAT app: error log with causes, study plan + daily checklist + missed-day
// reasons, journal check-ins and weekly reviews, and insights.
import { register, getCourse, go } from '../app.js';
import { store, today } from '../store.js';
import { esc, md, $, $$, mascot, toast, heatmap, bars, lineChart, download, fmtDate, confirmBox } from '../ui.js';
import { dayNum, dayStr, recallIn } from '../srs.js';
import { lessonList, isDone } from './home.js';
import { canOpen } from '../content.js';

const CAUSES = ['Didn’t know it yet', 'Careless slip', 'Wrong reference', 'Misread the task', 'Forgot the steps', 'Other'];
register('errors', async (main, filter = 'open') => {
  const list = store.d.errors.filter(e => filter === 'all' || (filter === 'open' ? !e.resolved : e.resolved));
  const byCause = {}; store.d.errors.forEach(e => { const k = e.cause || 'Not tagged'; byCause[k] = (byCause[k] || 0) + 1; });
  const byLesson = {}; store.d.errors.forEach(e => { if (e.lessonNum) byLesson[e.lessonNum] = (byLesson[e.lessonNum] || 0) + 1; });
  const open = store.d.errors.filter(e => !e.resolved).length, fixed = store.d.errors.filter(e => e.resolved).length;
  main.innerHTML = `<h1>Error log</h1><p class="lead">Every wrong lab answer and quiz answer lands here, beside the right one. Each comes back in your daily review after 1, 3 and 7 days until it sticks.</p>
  <div class="grid3"><div class="card stat"><b>${open}</b><span>open</span></div><div class="card stat"><b>${fixed}</b><span>fixed for good</span></div><div class="card stat"><b>${store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length}</b><span>due today</span></div></div>
  <div class="grid2"><section class="card"><h2>Causes so far</h2>${Object.keys(byCause).length ? bars(Object.entries(byCause).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 150, label: 'Mistakes by cause' }) + `<p class="muted small">${Object.entries(byCause).map(([k, v]) => `${esc(k)}: ${v}`).join(' · ')}</p>` : '<p class="muted">Tag each mistake with a cause to see your patterns.</p>'}</section>
  <section class="card"><h2>By lesson</h2>${Object.keys(byLesson).length ? bars(Object.entries(byLesson).sort().map(([k, v]) => ({ k, v })), { h: 150, label: 'Mistakes by lesson' }) : '<p class="muted">No mistakes logged yet.</p>'}</section></div>
  <section class="card"><div class="row between"><h2>All mistakes</h2><div class="seg">${['open', 'resolved', 'all'].map(f => `<a class="segb ${f === filter ? 'on' : ''}" href="#/errors/${f}">${f[0].toUpperCase() + f.slice(1)}</a>`).join('')}</div></div>
  ${list.length ? `<ul class="errs">${list.map(e => `<li class="err-item ${e.resolved ? 'res' : ''}"><div class="em"><span class="tag">${esc(e.lessonNum || '')} · ${e.kind}</span><span class="muted small">${fmtDate(e.at)}</span></div><p>${md(e.prompt)}</p><div class="cmp"><div><small>You</small><code>${esc(e.yours)}</code></div><div><small>Right</small><code>${esc(e.right)}</code></div></div>${e.why ? `<p class="muted small">${md(e.why)}</p>` : ''}
    <div class="row"><select data-id="${e.id}" class="cause" aria-label="Cause"><option value="">Why did it happen?</option>${CAUSES.map(c => `<option ${e.cause === c ? 'selected' : ''}>${c}</option>`).join('')}</select><a class="btn ghost sm" href="#/lesson/${e.course}/${e.lesson}">Practise in the lesson</a>${e.resolved ? '' : `<button class="btn ghost sm" data-fix="${e.id}">Mark fixed</button>`}</div></li>`).join('')}</ul>` : `<div class="empty-in">${mascot('encourage', 'sm')}<p>${filter === 'open' ? 'No open mistakes. Lovely.' : 'Nothing here yet.'}</p></div>`}</section>`;
  $$('.cause').forEach(s => s.onchange = () => { const e = store.d.errors.find(x => x.id === s.dataset.id); e.cause = s.value; store.save(); });
  $$('[data-fix]').forEach(b => b.onclick = () => { const e = store.d.errors.find(x => x.id === b.dataset.fix); e.resolved = true; e.resolvedAt = new Date().toISOString(); store.save(); toast('Marked as fixed'); b.closest('li').classList.add('res'); b.remove(); });
});

register('planner', async main => {
  const c = await getCourse('c1'); const p = store.d.plan; const left = lessonList(c).filter(l => !isDone(l.id));
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  // turn the target date + study days into a schedule of the remaining lessons (2 lessons on a 60+ minute day)
  const sched = []; if (p.days.length) { let d = new Date(); d.setHours(0, 0, 0, 0); const perDay = p.minutesGoal >= 60 ? 2 : 1; let k = 0; let guard = 0; while (k < left.length && guard++ < 800) { if (p.days.includes(d.getDay())) { sched.push({ date: new Date(d), lessons: left.slice(k, k + perDay) }); k += perDay; } d.setDate(d.getDate() + 1); } }
  const finish = sched.length ? sched[sched.length - 1].date : null; const target = p.target ? new Date(p.target) : null;
  const onTrack = finish && target ? finish <= target : null;
  const missed = Object.entries(p.missed).sort().reverse();
  const reasons = {}; missed.forEach(([, r]) => reasons[r] = (reasons[r] || 0) + 1);
  const wk = [0, 0, 0, 0, 0, 0, 0], wkPlan = [0, 0, 0, 0, 0, 0, 0]; Object.entries(store.d.activity).forEach(([d, a]) => { const wd = new Date(d + 'T00:00:00').getDay(); wk[wd] += a.minutes || 0; });
  main.innerHTML = `<h1>Study plan</h1><p class="lead">Planning exactly when you’ll study roughly doubles the chance you follow through. Pick your days and time; Sheetah lays out every remaining lesson.</p>
  <section class="card"><form id="pf" class="form"><fieldset><legend>Study days</legend><div class="days">${days.map((d, i) => `<label class="pill"><input type="checkbox" name="d" value="${i}" ${p.days.includes(i) ? 'checked' : ''}> ${d}</label>`).join('')}</div></fieldset>
   <div class="grid3"><label>Time <input type="time" name="time" value="${p.time}"></label><label>Minutes a day <select name="goal">${[15, 30, 45, 60, 90].map(m => `<option ${m === p.minutesGoal ? 'selected' : ''} value="${m}">${m} minutes</option>`).join('')}</select></label><label>Finish C1 by <input type="date" name="target" value="${p.target || ''}"></label></div>
   <div class="row"><button class="btn">Save plan</button><button type="button" class="btn ghost" id="ics">Add to my calendar (.ics)</button><button type="button" class="btn ghost" id="notif">Turn on reminders</button></div></form>
   ${finish ? `<p class="plan-sum ${onTrack === false ? 'warn' : ''}">At this pace you finish the ${left.length} remaining lesson${left.length === 1 ? '' : 's'} on <b>${finish.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' })}</b>${target ? (onTrack ? ', before your target. 👍' : `, after your target of ${target.toLocaleDateString('en-IN', { day: 'numeric', month: 'short' })}. Add a day or raise the minutes.`) : '.'} Then the project and final exam.</p>` : left.length ? '<p class="warn">Pick at least one study day.</p>' : '<p>Every lesson is done. Time for the project and the final exam.</p>'}</section>
  <section class="card"><h2>Day by day</h2>${sched.length ? `<table class="tbl plan"><thead><tr><th>Date</th><th>Lesson</th><th></th></tr></thead><tbody>${sched.slice(0, 40).map(s => `<tr><td>${s.date.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' })}</td><td>${s.lessons.map(l => `${esc(l.num)} ${esc(l.title)}`).join('<br>')}</td><td>${s.lessons.map(l => canOpen(c, l) ? `<a href="#/lesson/c1/${l.id}">Open</a>` : '🔒').join('<br>')}</td></tr>`).join('')}</tbody></table>` : '<p class="muted">Nothing left to schedule.</p>'}</section>
  <div class="grid2"><section class="card"><h2>Missed days and why</h2>${missed.length ? `<ul class="missed-list">${missed.slice(0, 12).map(([d, r]) => `<li><span>${fmtDate(d)}</span><b>${esc(r)}</b></li>`).join('')}</ul><h3>Reasons you noted</h3>${bars(Object.entries(reasons).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 130, label: 'Reasons for missed days' })}` : '<p class="muted">When you miss a planned day, Home asks what got in the way. Patterns show up here.</p>'}</section>
  <section class="card"><h2>When you actually study</h2>${bars(days.map((d, i) => ({ k: d, v: Math.round(wk[i]), hl: p.days.includes(i) })), { h: 150, fmt: v => v + 'm', label: 'Minutes by weekday' })}<p class="muted small">Highlighted bars are your planned days.</p></section></div>`;
  $('#pf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); p.days = f.getAll('d').map(Number); p.time = f.get('time'); p.minutesGoal = +f.get('goal'); p.target = f.get('target'); store.save(); toast('Plan saved'); go('#/planner'); };
  $('#ics').onclick = () => { const [h, m] = p.time.split(':'); const byday = p.days.map(i => ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'][i]).join(','); const st = new Date(); st.setHours(+h, +m, 0, 0); const f = d => d.toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z'; const en = new Date(st.getTime() + p.minutesGoal * 6e4);
    download('sheetah-study-plan.ics', `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Sheetah//Study plan//EN\r\nBEGIN:VEVENT\r\nUID:sheetah-${Date.now()}@sheetah\r\nDTSTAMP:${f(new Date())}\r\nDTSTART:${f(st)}\r\nDTEND:${f(en)}\r\nRRULE:FREQ=WEEKLY;BYDAY=${byday}\r\nSUMMARY:Sheetah: Excel study session\r\nDESCRIPTION:One lesson plus the daily review.\r\nBEGIN:VALARM\r\nTRIGGER:-PT10M\r\nACTION:DISPLAY\r\nDESCRIPTION:Excel study in 10 minutes\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`, 'text/calendar'); };
  $('#notif').onclick = async () => { if (!('Notification' in window)) return toast('This browser has no notifications; Sheetah will remind you inside the app.'); const r = await Notification.requestPermission(); store.d.settings.remind = true; store.save(); toast(r === 'granted' ? 'Reminders on. You’ll get one at your study time while Sheetah is open.' : 'Notifications are blocked; Sheetah will remind you inside the app instead.'); };
});

const isoWeek = d => { const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day); const y = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return t.getUTCFullYear() + '-W' + String(Math.ceil(((t - y) / 864e5 + 1) / 7)).padStart(2, '0'); };
register('journal', async main => {
  const t = today(); const j = store.d.journal[t] || {}; const wk = isoWeek(new Date()); const w = store.d.weekly[wk] || {};
  const past = Object.entries(store.d.journal).filter(([d]) => d !== t).sort().reverse();
  const ago = [7, 30].map(n => [n, store.d.journal[dayStr(dayNum() - n)]]).filter(([, e]) => e);
  const a = store.d.activity[t] || {};
  main.innerHTML = `<h1>Journal</h1><p class="lead">Two minutes a day: what you learned, what was hard, how sure you feel. Writing it down helps it stick, and the patterns show you what to review.</p>
  <div class="grid2"><section class="card"><h2>Today’s check-in</h2><p class="muted small">Today so far: ${Math.round(a.minutes || 0)} minutes, ${a.lessons || 0} lesson(s), ${a.reviews || 0} review(s).</p><form id="jf" class="form">
   <label>What did I learn? <textarea name="learned" rows="3" placeholder="e.g. $B$1 stays put when I copy a formula">${esc(j.learned || '')}</textarea></label>
   <label>What was hard? <textarea name="hard" rows="2">${esc(j.hard || '')}</textarea></label>
   <div class="grid2"><label>How confident do I feel? <input type="range" name="conf" min="1" max="5" value="${j.conf || 3}"><span class="rl"><small>Lost</small><small>Very sure</small></span></label><label>Energy <input type="range" name="energy" min="1" max="5" value="${j.energy || 3}"><span class="rl"><small>Drained</small><small>Fresh</small></span></label></div>
   <fieldset><legend>Mood</legend><div class="days">${['😫', '😕', '😐', '🙂', '😄'].map((m, i) => `<label class="pill"><input type="radio" name="mood" value="${i + 1}" ${+j.mood === i + 1 ? 'checked' : ''}> ${m}</label>`).join('')}</div></fieldset>
   <button class="btn">Save check-in</button></form></section>
  <section class="card"><h2>Weekly review · ${wk}</h2><form id="wf" class="form"><label>What went well this week? <textarea name="good" rows="2">${esc(w.good || '')}</textarea></label><label>What was tough? <textarea name="bad" rows="2">${esc(w.bad || '')}</textarea></label><label>One thing I’ll change next week <textarea name="change" rows="2">${esc(w.change || '')}</textarea></label><label>Rate the week <select name="rate">${[1, 2, 3, 4, 5].map(n => `<option ${+w.rate === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label><button class="btn">Save weekly review</button></form>
   ${ago.length ? `<h3>On this day</h3>${ago.map(([n, e]) => `<div class="onthisday"><small>${n} days ago</small><p>${esc(e.learned || '—')}</p></div>`).join('')}` : ''}</section></div>
  <section class="card"><h2>All entries</h2>${past.length ? `<ul class="journal-list">${past.slice(0, 60).map(([d, e]) => `<li><span class="jd">${fmtDate(d)}</span><span class="mood">${['', '😫', '😕', '😐', '🙂', '😄'][e.mood || 0] || ''}</span><div><p>${esc(e.learned || '')}</p>${e.hard ? `<p class="muted small">Hard: ${esc(e.hard)}</p>` : ''}</div><span class="conf" title="Confidence">${'●'.repeat(e.conf || 0)}${'○'.repeat(5 - (e.conf || 0))}</span></li>`).join('')}</ul>` : '<p class="muted">Your past check-ins appear here.</p>'}</section>`;
  $('#jf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); store.d.journal[t] = { learned: f.get('learned'), hard: f.get('hard'), conf: +f.get('conf'), energy: +f.get('energy'), mood: +f.get('mood') || 0, at: new Date().toISOString() }; const ck = store.d.checklist[t] || (store.d.checklist[t] = {}); ck.journal = true; store.save(); toast('Check-in saved'); };
  $('#wf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); store.d.weekly[wk] = { good: f.get('good'), bad: f.get('bad'), change: f.get('change'), rate: +f.get('rate') }; store.save(); toast('Weekly review saved'); };
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
  const last30 = Array.from({ length: 30 }, (_, k) => { const ds = dayStr(t0 - 29 + k); const a = d.activity[ds] || {}; return { k: ds.slice(8), v: Math.round(a.minutes || 0), r: a.reviews || 0 }; });
  const totalMin = Object.values(d.activity).reduce((a, x) => a + (x.minutes || 0), 0);
  const tests = d.tests.filter(x => x.kind === 'mock' || x.kind === 'module').slice(-15).map((x, i) => ({ k: fmtDate(x.at), v: x.pct }));
  const causes = {}; d.errors.forEach(e => { const k = e.cause || 'Not tagged'; causes[k] = (causes[k] || 0) + 1; });
  const moods = Object.entries(d.journal).sort().slice(-30).map(([k, e]) => ({ k, v: (e.conf || 0) * 20 }));
  main.innerHTML = `<h1>Insights</h1><div class="grid3"><div class="card stat"><b>${Math.round(totalMin / 60 * 10) / 10}</b><span>hours studied</span></div><div class="card stat"><b>${Object.values(d.lessons).filter(l => l.status === 'done').length}</b><span>lessons done</span></div><div class="card stat"><b>${Object.values(d.cards).reduce((a, x) => a + (x.reps || 0), 0)}</b><span>cards reviewed</span></div></div>
  <section class="card"><h2>Calendar</h2>${heatmap(d.activity, 26)}</section>
  <div class="grid2"><section class="card"><h2>Minutes · last 30 days</h2>${bars(last30.map(x => ({ k: x.k, v: x.v })), { h: 160, label: 'Minutes per day' })}</section><section class="card"><h2>Reviews · last 30 days</h2>${bars(last30.map(x => ({ k: x.k, v: x.r })), { h: 160, label: 'Reviews per day' })}</section></div>
  <div class="grid2"><section class="card"><h2>Test scores</h2>${lineChart(tests, { target: 75, label: 'Test scores' })}</section><section class="card"><h2>Confidence from your journal</h2>${lineChart(moods.map(x => ({ k: x.k, v: x.v })), { label: 'Confidence', max: 100 })}</section></div>
  <section class="card"><h2>Mastery map</h2><p class="muted small">Each skill brightens as you review it and dims when it’s due again. Mastered means at least a 90% chance of recall in 30 days and three lab tasks right first time.</p><div class="mastery">${c.modules.map(m => `<div class="mrow"><span class="mlabel">M${m.n}</span>${m.lessons.map(l => { const lv = mastery(l.id); return `<a class="mcell lv${lv}" href="#/lesson/c1/${l.id}" title="${esc(l.num + ' ' + l.title)}: ${LEVELS[lv]}">${l.num}</a>`; }).join('')}</div>`).join('')}</div><div class="legend">${LEVELS.map((n, i) => `<span class="mcell lv${i} sm"></span>${n}`).join(' ')}</div></section>
  <section class="card"><h2>Mistakes by cause</h2>${Object.keys(causes).length ? bars(Object.entries(causes).map(([k, v]) => ({ k: k.split(' ')[0], v })), { h: 140, label: 'Mistakes by cause' }) : '<p class="muted">No mistakes logged yet.</p>'}</section>`;
});
