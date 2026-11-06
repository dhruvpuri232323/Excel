// Home (what's next today), the Courses page, the course's learning map and each module's page.
import { register, app, getCourse, go, lockedPanel, liveCourses, setActiveCourse, activeCourseId } from '../app.js';
import { store, today } from '../store.js';
import { session, assetURL } from '../content.js';
import { esc, $, $$, mascot, line, ring, heatmap, bar, toast, md, crumbs, qhelp } from '../ui.js';
import { dayNum, dayStr } from '../srs.js';
import { freeModulesText, lessonList, isDone, moduleInfo, nextStep, stepLabel, stepHref, PART_NAMES, partMinutes, lessonMinutesLeft, lessonLabel, hasKey, who, projectOpen, examReqs, fmtMin, TEST_MINUTES, testsFor, lessonNo, PARTS } from '../access.js';
import { mountMap, miniMap } from '../map.js';
import { site } from '../site.js';
import { finishDate } from './setup.js';
import { todayItems, planStatus } from '../plan.js';

export { lessonList, isDone };
export function courseProgress(c) { const all = lessonList(c); const done = all.filter(l => isDone(l.id)).length; const tests = c.modules.filter(m => moduleInfo(c, m).passed).length; return { done, total: all.length, tests, pct: (done + tests) / (all.length + c.modules.length) }; }
const MILESTONES = [
  ['first', 'First lesson', d => Object.values(d.lessons).some(l => l.status === 'done')], ['five', 'Five lessons', d => Object.values(d.lessons).filter(l => l.status === 'done').length >= 5],
  ['streak3', '3-day streak', () => store.bestStreak() >= 3], ['streak7', '7-day streak', () => store.bestStreak() >= 7], ['cards50', '50 reviews', d => Object.values(d.cards).reduce((a, c) => a + (c.reps || 0), 0) >= 50],
  ['fix10', '10 mistakes fixed', d => d.errors.filter(e => e.resolved).length >= 10], ['test', 'First module test passed', d => d.tests.some(t => t.kind === 'module' && t.passed)], ['cert', 'First certificate', d => d.certs.length > 0]];
const SUGGEST = {
  job: ['Interview prep', '20 common Excel interview questions with model answers, and a timed mock round.', '#/interview'],
  college: ['Practice tests', 'Mixed, timed questions from the lessons you’ve done, with a target score.', '#/tests'],
  business: ['This week’s challenge', 'A short real-world case, like a GST invoice that adds up.', '#/challenge'],
  work: ['Shortcut trainer', 'Press the real keys and Sheetah checks them. Ten a round.', '#/shortcuts'],
  curious: ['Playground', 'A free sheet with sample data from a kirana shop, a payroll and a cricket season.', '#/playground'],
};

register('home', async main => {
  const c = await getCourse(); const pr = courseProgress(c); const step = nextStep(c);
  const t = store.d.activity[today()] || { minutes: 0, lessons: 0, reviews: 0 }; const goal = store.d.plan.minutesGoal || 30;
  const due = Object.values(store.d.cards).filter(x => x.due <= dayNum() && !x.isNew).length; const errDue = store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length;
  const journalDone = !!store.d.journal[today()]; const st = store.streak(); const anyDone = Object.values(store.d.lessons).some(l => l.status === 'done');
  const y = dayStr(dayNum() - 1); const yDay = new Date(Date.now() - 864e5).getDay(); const askMissed = store.d.plan.days.includes(yDay) && !(store.d.plan.off || []).includes(y) && !store.studied(y) && !store.d.plan.missed[y] && Object.keys(store.d.activity).some(d => d < y) && (store.d.createdAt || '').slice(0, 10) < y;
  // the next step, with the five parts of the lesson as a little track
  const ls = step.kind === 'lesson' ? store.d.lessons[step.lesson.id] || {} : null; const reach = ls ? (ls.status === 'done' ? 5 : Math.max(ls.maxPart || 0, ls.part || 0)) : 0;
  const track = ls ? `<div class="nx-track">${PART_NAMES.map((n, k) => `<span class="${k < reach ? 'done' : k === reach && step.continue ? 'on' : ''}">${k < reach ? '✓' : k + 1} ${n}</span>`).join('')}</div>` : '';
  const nextTitle = step.kind === 'lesson' ? esc(step.lesson.title) : esc(stepLabel(step));
  const nextLabel = step.kind === 'lesson' ? `${esc(lessonLabel(step.lesson))} · ${step.continue ? `continue at the ${PART_NAMES[step.part].toLowerCase()}` : 'new lesson'}` : step.kind === 'test' ? 'Every lesson in this module is done' : '';
  // today, from the study plan: what's planned, plus what's already been done today
  const { P, list, next } = todayItems(c); const S = planStatus(c); const td = today();
  const doneToday = lessonList(c).filter(l => (store.d.lessons[l.id] || {}).doneAt && store.d.lessons[l.id].doneAt.slice(0, 10) === td).map(l => ({ kind: 'lesson', short: `${l.module.n}.${l.num.split('.')[1]}`, title: l.title, href: `#/lesson/${c.id}/${l.id}`, min: l.minutes, done: true }));
  const row = (icon, label, sub, href, done, cls = '') => `<li class="${done ? 'ok' : ''} ${cls}"><a href="${href}"><span class="tk" aria-hidden="true">${done ? '✓' : icon}</span><span class="grow">${label}</span>${sub ? `<small>${sub}</small>` : ''}</a></li>`;
  const planRows = [...doneToday, ...list].map(x => row(x.kind === 'test' ? '📝' : x.kind === 'project' ? '📁' : x.kind === 'exam' ? '🎓' : '▶', x.kind === 'lesson' ? `<b>${esc(x.short)}</b> ${esc(x.title)}` : `<b>${esc(x.title)}</b>`, x.done ? 'done' : `${x.min} min`, x.href, x.done, 'k-' + x.kind)).join('');
  const studyDay = store.d.plan.days.includes(new Date().getDay()) && !(store.d.plan.off || []).includes(td);
  const planNote = !store.d.plan.end ? `<a class="plan-cta" href="#/planner">📅 <b>Set your finish date</b> and Sheetah plans every day for you →</a>` : P.error ? `<a class="plan-cta" href="#/planner">⚠ ${esc(P.error)} Fix it in your study plan →</a>` : !list.length && !doneToday.length ? `<p class="muted small">${studyDay ? 'Nothing planned for today.' : 'Rest day: not a study day on your plan.'}${next ? ` Next: <b>${new Date(next + 'T00:00:00').toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' })}</b>, ${(P.days[next] || []).map(x => x.kind === 'lesson' ? 'Lesson ' + x.short : x.short).join(', ')}.` : ''}</p>` : '';
  const pill = S && P.items.length && !P.error ? (S.behind ? `<span class="pstat warn">⏳ ${S.behind} behind</span>` : S.ahead ? `<span class="pstat ok">🚀 ${S.ahead} ahead</span>` : '<span class="pstat ok">✓ On track</span>') : '';
  const sug = SUGGEST[store.d.profile.why];
  const lives = liveCourses(); const switcher = lives.length > 1 ? `<div class="cswitch" role="group" aria-label="Which course">${lives.map(x => `<button type="button" class="cs ${x.id === c.id ? 'on' : ''}" data-course="${x.id}" aria-pressed="${x.id === c.id}">${esc(x.code || '')} ${esc(x.short || x.title)}</button>`).join('')}</div>` : '';
  main.innerHTML = `
  ${!store.d.tourDone ? `<section class="card tourcall"><span>👋</span><p><b>New here?</b> A one-minute tour shows where everything lives.</p><div class="row"><a class="btn sm" href="#/tour">Take the tour</a><button class="linkbtn" id="noTour" type="button">No thanks</button></div></section>` : ''}
  <div class="home-top">
    <section class="card nextcard">${switcher}<p class="eyebrow">${esc(line('hi', store.d.name || 'there'))}</p><span class="nx-kicker">Next up</span><h1>${nextTitle}</h1><p class="nx-label">${nextLabel}</p>${track}
      <div class="row"><a class="btn big" href="${stepHref(c, step)}">${step.kind === 'lesson' ? (step.continue ? '▶ Continue' : '▶ Start') : esc(stepLabel(step))}${step.minutes ? ` · ${step.minutes} min` : ''}</a>${due + errDue ? `<a class="btn ghost" href="#/review">↻ Review ${due + errDue}</a>` : ''}</div></section>
    <section class="card stats3"><div class="st-ring">${ring(Math.min(1, t.minutes / goal), 'Today’s minutes')}<span><b>${Math.round(t.minutes)}</b>/${goal} min today</span></div><div class="st-row"><div><b>🔥 ${st}</b><span>day streak</span></div><div><b>${pr.done}/${pr.total}</b><span>lessons</span></div><div><b>${pr.tests}/${c.modules.length}</b><span>tests</span></div></div>${!store.studied(y) && store.freezesLeft() > 0 && st === 0 && Object.keys(store.d.activity).length ? '<button class="btn ghost sm" id="freeze" type="button">Use a streak freeze for yesterday</button>' : ''}</section>
  </div>
  ${askMissed ? `<section class="card missed"><b>Yesterday was a study day and it didn’t happen.</b> What got in the way? <span class="muted">It helps you plan better.</span><div class="row">${['Work ran late', 'Too tired', 'Forgot', 'Travel', 'Exams or college', 'Other'].map(r => `<button class="pill" data-r="${r}" type="button">${r}</button>`).join('')}</div></section>` : ''}
  <div class="grid2">
   <section class="card" id="todayCard"><div class="row between"><h2>Today</h2>${pill}</div><ul class="todo">${planRows}${row('↻', 'Daily review', due + errDue ? `${due + errDue} due · 10 min` : 'all done', '#/review', t.reviews > 0 || !(due + errDue))}${row('✐', 'Journal check-in', '2 min', '#/journal', journalDone)}</ul>${planNote}<a class="small" href="#/planner">Open the study plan →</a></section>
   <section class="card"><div class="row between"><h2>${esc(c.title)}</h2>${qhelp('map')}</div>${miniMap(c)}${bar(pr.pct)}<p class="small muted">${Math.round(pr.pct * 100)}% done · ${pr.done} of ${pr.total} lessons · ${pr.tests} of ${c.modules.length} module tests</p><a href="#/courses">See the whole journey →</a></section>
  </div>
  ${sug ? `<section class="card suggest"><div><p class="eyebrow">Picked for you</p><h2>${esc(sug[0])}</h2><p>${esc(sug[1])}</p></div><a class="btn ghost" href="${sug[2]}">Open</a></section>` : ''}
  ${anyDone ? `<details class="card more-home"><summary><b>Study calendar and milestones</b></summary><div class="row between"><h3>Study calendar</h3><a href="#/insights">All insights →</a></div>${heatmap(store.d.activity, 20)}<h3>Milestones</h3><div class="badges">${MILESTONES.map(([k, label, test]) => { const got = test(store.d); return `<div class="badge ${got ? 'got' : ''}" title="${got ? 'Earned' : 'Not yet'}"><span>${got ? '★' : '☆'}</span>${esc(label)}</div>`; }).join('')}</div></details>` : ''}
  ${hasKey(c) ? '' : `<section class="card free-note">${mascot('thinking', 'sm')}<p>${session.signedIn ? `<b>Your login doesn’t include ${esc(c.title)} yet.</b> ${freeModulesText(c)} ${/–/.test(freeModulesText(c)) ? 'are' : 'is'} free for everyone; ask your teacher to add the course to your login for the rest. <button class="linkbtn" type="button" data-ask="login">Ask your teacher</button>` : `<b>Free access:</b> ${freeModulesText(c)} of ${esc(c.title)}, and every tool. A student login from your teacher opens the rest, module by module. <a href="#/login">Student sign-in</a> · <button class="linkbtn" type="button" data-ask="login">Ask for a login</button>`}</p></section>`}`;
  $$('[data-course]').forEach(b => b.onclick = () => { setActiveCourse(b.dataset.course); go('#/home'); });
  $$('[data-r]').forEach(b => b.onclick = () => { store.d.plan.missed[y] = b.dataset.r; store.save(); toast('Noted. Your study plan shows the patterns.'); b.closest('.missed').remove(); });
  const fz = $('#freeze'); if (fz) fz.onclick = () => { if (store.useFreeze()) { toast('Freeze used. Your streak is safe.'); go('#/home'); } };
  const nt = $('#noTour'); if (nt) nt.onclick = () => { store.d.tourDone = true; store.save(); nt.closest('.tourcall').remove(); };
});

// My courses: the whole plan on one page. The journey across all eight courses with "You are here", then the
// current course's map, next step, time left and finish date, then every course card.
const grp = x => x.group || (/^p/i.test(x.id) ? 'specialist' : 'core');
register('courses', async main => {
  let courses = (app.catalog && app.catalog.courses) || [];
  if (!courses.length) { const c1 = await getCourse('c1'); courses = [{ id: 'c1', title: c1.title, summary: c1.summary, modules: c1.modules.length, lessons: lessonList(c1).length, hours: c1.hours, status: 'live', group: 'core' }]; }
  const core = courses.filter(x => grp(x) === 'core'), spec = courses.filter(x => grp(x) === 'specialist');
  const info = {}; for (const x of courses.filter(x => x.status === 'live')) { const c = await getCourse(x.id); const pr = courseProgress(c); const R = examReqs(c); const certified = store.d.certs.some(ct => ct.course === c.id); info[x.id] = { c, pr, step: nextStep(c), certified, R, left: c.modules.reduce((a, m) => a + moduleInfo(c, m).minutesLeft, 0) }; }
  const current = (info[activeCourseId()] && courses.find(x => x.id === activeCourseId())) || courses.find(x => info[x.id] && !info[x.id].certified) || courses.find(x => info[x.id]);
  const doneCount = courses.filter(x => info[x.id] && info[x.id].certified).length;
  const state = x => !info[x.id] ? 'soon' : info[x.id].certified ? 'done' : info[x.id].pr.done || info[x.id].pr.tests ? 'progress' : 'open';
  const STATE = { soon: 'Coming soon', done: 'Certificate earned', progress: 'In progress', open: 'Ready to start' };
  // the journey: Foundations → Formulas & Functions → one direction → Master Certificate
  const stage = (x, n) => { const st = state(x), here = current && x.id === current.id, i = info[x.id]; return `<a class="jstage st-${st} ${here ? 'here' : ''}" href="${i ? '#/course/' + x.id : '#/courses'}" ${i ? '' : 'aria-disabled="true"'}>${here ? '<span class="here-tag">You are here ▼</span>' : ''}<span class="jn">${st === 'done' ? '✓' : n}</span><b>${esc(x.title)}</b><small>${i ? `${Math.round(i.pr.pct * 100)}% · ${i.pr.done} of ${i.pr.total} lessons` : STATE[st]}</small>${i ? bar(i.pr.pct, 'thin') : ''}</a>`; };
  const journey = `<div class="journeymap">${core.map((x, k) => stage(x, k + 1)).join('<span class="jarrow" aria-hidden="true"></span>')}<span class="jarrow" aria-hidden="true"></span>
    <div class="jstage st-soon jdir"><span class="jn">${core.length + 1}</span><b>Choose a direction</b><small>One of ${spec.length} specialist courses</small><div class="jchips">${spec.map(x => `<span class="${state(x) === 'soon' ? 'soon' : ''}" title="${esc(x.summary)}">${esc(x.title)}</span>`).join('')}</div></div><span class="jarrow" aria-hidden="true"></span>
    <div class="jstage st-${doneCount === courses.length ? 'done' : 'soon'} jmaster"><span class="jn">★</span><b>Master Certificate</b><small>${doneCount} of ${courses.length} courses done</small></div></div>`;
  // where you are now, in the current course
  let now = '';
  if (current && info[current.id]) { const { c, pr, step, R, left } = info[current.id]; const fd = finishDate(c); const p = store.d.plan; const D = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const mod = step.lesson ? step.lesson.module : step.module || null;
    const where = step.kind === 'lesson' ? `Module ${mod.n} of ${c.modules.length}: <b>${esc(mod.title)}</b> · Lesson ${lessonNo(step.lesson)}: ${esc(step.lesson.title)}${step.continue ? ` · ${PART_NAMES[step.part]}` : ''}` : step.kind === 'test' ? `Module ${mod.n} of ${c.modules.length}: <b>${esc(mod.title)}</b> · every lesson done, the module test is next` : step.kind === 'locked' ? `Module ${mod.n} of ${c.modules.length}: <b>${esc(mod.title)}</b> · ${step.info.reason === 'login' ? 'needs a student login' : `opens when you pass the Module ${step.info.prev.n} test`}` : step.kind === 'project' ? 'Every module test passed · the <b>course project</b> is next' : step.kind === 'exam' ? 'Project passed · the <b>final exam</b> is next' : 'Course complete';
    now = `<section class="card nowcard"><div class="row between"><div><p class="eyebrow">Where you are now</p><h2>${esc(c.title)}</h2></div><a class="btn" href="${stepHref(c, step)}">${esc(stepLabel(step))}${step.minutes ? ` · ${step.minutes} min` : ''}</a></div>
      <p class="lead">${where}</p>
      <div class="nowstats"><div><b>${pr.done}<small>/${pr.total}</small></b><span>lessons done</span></div><div><b>${pr.tests}<small>/${c.modules.length}</small></b><span>module tests passed</span></div><div><b>${R.projOk ? '✓' : '—'}</b><span>course project</span></div><div><b>${(store.d.examAttempts[c.id] || []).some(a => a.passed) ? '✓' : '—'}</b><span>final exam</span></div><div><b>${left >= 60 ? Math.round(left / 60) + '<small> h</small>' : left + '<small> min</small>'}</b><span>left in the modules</span></div></div>
      <div id="nowmap" class="mapbox"></div><div class="legend maplegend"><span class="lg st-done"></span>Done <span class="lg st-progress"></span>In progress <span class="lg st-open"></span>Open <span class="lg st-locked"></span>Locked <span class="lg gate"></span>Module test (pass 70%) ${qhelp('map')}</div>
      <div class="planline">${p.days.length ? `📅 You study <b>${p.days.map(d => D[d]).join(', ')}</b> at ${esc(p.time)}, ${p.minutesGoal} minutes a day.${fd ? ` Your plan finishes everything, project and exam included, by <b>${fd.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}</b>.` : ''}` : '📅 Pick a finish date to get your day-by-day plan.'} <a href="#/planner">Change the plan</a> · <a href="#/course/${c.id}">Course page</a></div></section>`; }
  const card = x => { const st = state(x), i = info[x.id]; return i ? `<a class="card course-card live st-${st}" href="#/course/${x.id}"><p class="eyebrow">${grp(x) === 'core' ? 'Core course' : 'Specialist'} · ${STATE[st]}</p><h2>${esc(x.title)}</h2><p>${esc(x.summary)}</p><p class="meta">${x.modules} modules · ${x.lessons} lessons · about ${x.hours} hours</p>${bar(i.pr.pct)}<span class="btn sm">${st === 'open' ? 'Open the map' : st === 'done' ? 'See the course' : 'Continue'} →</span></a>`
    : `<div class="card course-card soon"><p class="eyebrow">${grp(x) === 'core' ? 'Core course' : 'Specialist'} · Coming soon</p><h2>${esc(x.title)}</h2><p>${esc(x.summary)}</p><p class="meta">${x.modules} modules · about ${x.hours} hours</p></div>`; };
  main.innerHTML = `<h1>My courses</h1><p class="lead">Your whole plan in one place: the route through every course, where you are now, and what comes next.</p>
  <section class="card"><div class="row between"><h2>Your journey</h2><span class="muted small">${doneCount} of ${courses.length} courses complete</span></div>${journey}</section>
  ${now}
  <section class="journey"><h2>1. Start here: the core courses</h2><div class="core-row">${core.map(card).join('<span class="arrow" aria-hidden="true">→</span>')}</div>
   <h2>2. Then choose a direction</h2><p class="muted">Pick the specialist course that fits your work. Each one opens when it’s ready.</p><div class="spec-grid">${spec.map(card).join('')}</div>
   <h2>3. The Master Certificate</h2><div class="card master-line"><p>★ Finish all ${courses.length} courses to unlock a capstone business case and a three-hour exam.</p>${bar(doneCount / Math.max(1, courses.length))}<p class="small muted">${doneCount} of ${courses.length} done</p></div></section>`;
  if (current && info[current.id] && $('#nowmap')) mountMap($('#nowmap'), info[current.id].c);
});

register('course', async (main, id = 'c1') => {
  const c = await getCourse(id); const pr = courseProgress(c); const step = nextStep(c);
  main.innerHTML = `${crumbs([['My courses', '#/courses'], [c.title]])}
  <section class="card course-head"><div class="grow"><p class="eyebrow">${esc(c.level)} · about ${c.hours} hours · ${c.modules.length + (c.upcoming || []).length} modules${c.upcoming && c.upcoming.length ? ` · ${c.modules.length} out now` : ''}</p><h1>${esc(c.title)}</h1>${bar(pr.pct)}<p>${pr.done} of ${pr.total} lessons · ${pr.tests} of ${c.modules.length} module tests passed</p>
   <div class="row"><a class="btn big" href="${stepHref(c, step)}">${esc(stepLabel(step))}</a>${c.downloads.map(d => `<button class="btn ghost" data-dl="${d.file}" data-name="${esc(d.name)}" type="button">⬇ ${esc(d.label)}</button>`).join('')}<a class="btn ghost" href="#/placement/${c.id}">Placement check</a></div></div></section>
  <section class="card"><div class="row between"><h2>Your learning map</h2>${qhelp('map')}</div><div id="lmap" class="mapbox"></div><div class="legend maplegend"><span class="lg st-done"></span>Done <span class="lg st-progress"></span>In progress <span class="lg st-open"></span>Open <span class="lg st-locked"></span>Locked <span class="lg gate"></span>Module test (pass 70%)</div></section>
  <section class="card"><h2>By the end, you can</h2><ol class="outcomes">${c.outcomes.map(o => `<li>${esc(o)}</li>`).join('')}</ol></section>
  <div class="modgrid">${c.modules.map(m => { const i = moduleInfo(c, m); return `<a class="card modcard st-${i.status}" href="#/module/${c.id}/${m.n}"><p class="eyebrow">Module ${m.n}${m.free ? ' · Free' : ''}</p><h3>${esc(m.title)}</h3><p class="small muted">${esc(m.summary)}</p>${bar(i.done / i.total)}<p class="small"><b>${{ done: '✓ Done', testedout: '✓ Tested out', progress: 'In progress', open: 'Open', locked: '🔒 Locked' }[i.status]}</b> · ${i.done}/${i.total} lessons${i.open && !i.passed ? ` · ${fmtMin(i.minutesLeft)} left` : ''}</p></a>`; }).join('')}${(c.upcoming || []).map(u => `<div class="card modcard st-locked soon"><p class="eyebrow">Module ${u.n} · Coming soon</p><h3>${esc(u.title)}</h3><p class="small muted">${esc(u.summary)}</p><p class="small"><b>${u.lessons} lessons</b> · being written now</p></div>`).join('')}</div>`;
  mountMap($('#lmap'), c);
  $$('[data-dl]').forEach(b => b.onclick = async () => { const url = await assetURL(c.id, b.dataset.dl); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); });
});

register('module', async (main, id = 'c1', n = '1') => {
  const c = await getCourse(id); const m = c.modules.find(x => x.n === +n); if (!m) throw new Error('No such module');
  const i = moduleInfo(c, m); const idx = c.modules.indexOf(m); const prevM = c.modules[idx - 1], nextM = c.modules[idx + 1]; const tries = testsFor(c, m.n);
  const nextL = m.lessons.find(l => !isDone(l.id)); const key = hasKey(c) || m.free;
  const lessonRow = l => { const s = store.d.lessons[l.id]; const stt = s ? s.status : 'new'; const left = lessonMinutesLeft(l);
    return `<li class="${stt}"><a href="${i.open ? `#/lesson/${c.id}/${l.id}` : '#'}" ${i.open ? '' : 'aria-disabled="true" class="off"'}><span class="ls ${stt}">${stt === 'done' ? '✓' : stt === 'started' ? '◐' : i.open ? '' : '🔒'}</span><span class="ln">Lesson ${lessonNo(l)}</span><span class="lt">${esc(l.title)}<small>${esc(l.summary)}</small></span><span class="lm">${stt === 'done' ? 'Done' : stt === 'started' ? `${left} min left` : `${l.minutes} min`}</span></a>${stt === 'done' ? `<a class="btn ghost sm prac" href="#/practice/${c.id}/${l.id}">Practise</a>` : ''}</li>`; };
  const lockBox = !i.open ? (i.reason === 'login' ? `<div class="lockbox">${mascot('thinking', 'sm')}<div><b>${session.signedIn ? `Your login doesn’t include ${esc(c.title)} yet.` : 'This module needs a student login.'}</b><p>${freeModulesText(c)} ${/–/.test(freeModulesText(c)) ? 'are' : 'is'} free. ${session.signedIn ? 'Ask your teacher to add this course to your login; you’ll get a new password.' : 'With a login from your teacher, the rest opens module by module.'}</p><div class="row"><a class="btn sm" href="#/login">Student sign-in</a><button class="btn ghost sm" type="button" data-ask="login">Ask for a login</button></div></div></div>`
    : `<div class="lockbox">${mascot('thinking', 'sm')}<div><b>Pass the Module ${i.prev.n} test to open this.</b><p>${i.prevLeft ? `You’re ${i.prevLeft} lesson${i.prevLeft === 1 ? '' : 's'} away from that test.` : 'That test is ready for you now.'} Already know this module? Take its test now to test out. ${qhelp('testout')}</p><div class="row"><a class="btn sm" href="#/${i.prevLeft ? `module/${c.id}/${i.prev.n}` : `mtest/${c.id}/${i.prev.n}`}">${i.prevLeft ? `Go to Module ${i.prev.n}` : `Take the Module ${i.prev.n} test`}</a><a class="btn ghost sm" href="#/mtest/${c.id}/${m.n}">Test out of Module ${m.n}</a></div></div></div>`) : '';
  main.innerHTML = `${crumbs([['My courses', '#/courses'], [c.title, '#/course/' + c.id], [`Module ${m.n}`]])}
  <section class="card module-head st-${i.status}"><div class="grow"><p class="eyebrow">Module ${m.n} of ${c.modules.length}${m.free ? ' · Free' : ''} · ${{ done: 'Done', testedout: 'Tested out', progress: 'In progress', open: 'Open', locked: 'Locked' }[i.status]}</p><h1>${esc(m.title)}</h1><p class="lead">${esc(m.summary)}</p>${bar(i.done / i.total)}
    <p>${i.done} of ${i.total} lessons${i.passed ? ' · module test passed' : i.open ? ` · ${fmtMin(i.minutesLeft)} left, including the test` : ''}</p>
    ${i.open ? `<div class="row">${nextL ? `<a class="btn big" href="#/lesson/${c.id}/${nextL.id}">${store.d.lessons[nextL.id] && store.d.lessons[nextL.id].status === 'started' ? 'Continue' : 'Start'}: Lesson ${lessonNo(nextL)} · ${esc(nextL.title)}</a>` : ''}<a class="btn ${nextL ? 'ghost' : 'big'}" href="#/mtest/${c.id}/${m.n}">${i.passed ? 'Retake the test for practice' : 'Take the module test'}</a></div>` : lockBox}</div></section>
  <div class="module-cols"><div>
   <section class="card"><h2>Lessons</h2><p class="muted small">Take them in any order. The suggested order is shown. ${qhelp('parts')}</p><ol class="lessons">${m.lessons.map(lessonRow).join('')}</ol></section>
   <section class="card test-card"><div class="row between"><h2>Module test</h2>${qhelp('testrules')}</div><p>${esc(m.testInfo)} · about ${TEST_MINUTES} minutes · pass mark 70%</p><p class="small">${tries.length ? `Attempts: ${tries.length} · best ${Math.max(...tries.map(t => t.pct))}%${i.passed ? (i.testedOut ? ' · tested out ✓' : ' · passed ✓') : ''}` : 'Not taken yet.'}</p>${key ? `<a class="btn ${i.open ? '' : 'ghost'}" href="#/mtest/${c.id}/${m.n}">${i.open ? 'Open the test' : 'Test out of this module'}</a>` : ''}</section>
  </div><div>
   <section class="card"><h2>What you’ll learn</h2><ul class="ticks">${m.lessons.flatMap(l => (l.objectives || []).slice(0, 2)).map(o => `<li>${md(o)}</li>`).join('')}</ul></section>
   <section class="card"><h2>Topics in this module</h2>${m.lessons.map(l => `<h3 class="small-h">Lesson ${lessonNo(l)}: ${esc(l.title)}</h3><div class="topic-chips">${(l.topics || []).map(([si, k, h]) => i.open ? `<a class="tchip" href="#/lesson/${c.id}/${l.id}/lecture/${si + 1}-${k + 1}">${esc(h)}</a>` : `<span class="tchip off">${esc(h)}</span>`).join('')}</div>`).join('')}</section>
   <section class="card"><h2>Downloads</h2>${key ? '' : '<p class="muted small">🔒 These open with a student login. Every free lesson’s files are on its own page.</p>'}${m.lessons.filter(l => (l.downloads || []).length).map(l => `<h3 class="small-h">Lesson ${lessonNo(l)}: ${esc(l.title)}</h3><div class="dl-row">${l.downloads.map(d => `<button class="dl sm" data-file="${d.file}" data-name="${esc(d.name)}" type="button" ${key ? '' : 'disabled'}><span class="ext ${d.file.split('.').pop()}">${d.file.split('.').pop().toUpperCase()}</span>${esc(d.label)}</button>`).join('')}</div>`).join('')}<p class="small"><a href="#/downloads">Every download in the course →</a></p></section>
  </div></div>
  <div class="row between">${prevM ? `<a class="btn ghost" href="#/module/${c.id}/${prevM.n}">← Module ${prevM.n}</a>` : '<span></span>'}${nextM ? `<a class="btn ghost" href="#/module/${c.id}/${nextM.n}">Module ${nextM.n} →</a>` : c.upcoming && c.upcoming.length ? `<span class="muted small">Module ${c.upcoming[0].n}: ${esc(c.upcoming[0].title)} is coming soon</span>` : `<a class="btn ghost" href="#/project/${c.id}">Course project →</a>`}</div>`;
  $$('[data-file]').forEach(b => b.onclick = async () => { try { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); } catch (e) { toast('Sign in to download this file.'); } });
  $$('a[aria-disabled="true"]').forEach(a => a.onclick = e => { e.preventDefault(); toast(i.reason === 'login' ? 'This module needs a student login.' : `Pass the Module ${i.prev.n} test to open this, or test out of it.`); });
});
