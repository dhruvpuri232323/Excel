import { register, app, getCourse, go, lockedPanel } from '../app.js';
import { store, today } from '../store.js';
import { session, canOpen, assetURL } from '../content.js';
import { esc, $, $$, mascot, line, ring, heatmap, bar, toast, md } from '../ui.js';
import { dayNum, dayStr } from '../srs.js';

export function lessonList(c) { const out = []; for (const m of c.modules) for (const l of m.lessons) out.push({ ...l, module: m }); return out; }
export function nextLesson(c) { const all = lessonList(c); return all.find(l => store.d.lessons[l.id] && store.d.lessons[l.id].status === 'started' && canOpen(c, l)) || all.find(l => (!store.d.lessons[l.id] || store.d.lessons[l.id].status !== 'done') && canOpen(c, l)) || null; }
export function courseProgress(c) { const all = lessonList(c); const done = all.filter(l => store.d.lessons[l.id] && store.d.lessons[l.id].status === 'done').length; const tests = c.modules.filter(m => (store.d.tests.find(t => t.kind === 'module' && t.module === c.id + '-m' + m.n && t.passed))).length; return { done, total: all.length, tests, pct: (done + tests) / (all.length + c.modules.length) }; }
export const isDone = id => store.d.lessons[id] && store.d.lessons[id].status === 'done';
const MILESTONES = [
  ['first', 'First lesson', d => Object.values(d.lessons).some(l => l.status === 'done')], ['five', 'Five lessons', d => Object.values(d.lessons).filter(l => l.status === 'done').length >= 5],
  ['streak3', '3-day streak', () => store.bestStreak() >= 3], ['streak7', '7-day streak', () => store.bestStreak() >= 7], ['cards50', '50 reviews', d => Object.values(d.cards).reduce((a, c) => a + (c.reps || 0), 0) >= 50],
  ['fix10', '10 mistakes fixed', d => d.errors.filter(e => e.resolved).length >= 10], ['test', 'First module test passed', d => d.tests.some(t => t.kind === 'module' && t.passed)], ['cert', 'First certificate', d => d.certs.length > 0]];

register('home', async main => {
  const c = await getCourse('c1'); const nx = nextLesson(c); const pr = courseProgress(c);
  const t = store.d.activity[today()] || { minutes: 0, lessons: 0, reviews: 0 }; const goal = store.d.plan.minutesGoal || 30;
  const due = Object.values(store.d.cards).filter(x => x.due <= dayNum()).length; const errDue = store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length;
  const ck = store.d.checklist[today()] || {}; const journalDone = !!store.d.journal[today()];
  const items = [['review', `Daily review${due + errDue ? ` · ${due + errDue} due` : ''}`, t.reviews > 0 || (due + errDue === 0 && Object.keys(store.d.cards).length > 0), '#/review'], ['lesson', nx ? `Lesson ${nx.num}: ${nx.title}` : 'All open lessons done', t.lessons > 0, nx ? `#/lesson/c1/${nx.id}` : '#/course/c1'], ['journal', 'Journal check-in (2 minutes)', journalDone, '#/journal']];
  const left = pr.total - pr.done; const st = store.streak();
  // missed-day prompt (from the CAT app's "missed days and why")
  const y = dayStr(dayNum() - 1); const yDay = new Date(Date.now() - 864e5).getDay(); const askMissed = store.d.plan.days.includes(yDay) && !store.studied(y) && !store.d.plan.missed[y] && Object.keys(store.d.activity).length > 0;
  main.innerHTML = `
  <section class="hero card">${mascot(st >= 3 ? 'encourage' : 'hello')}<div class="grow"><p class="eyebrow">${new Date().toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long' })}</p><h1>${esc(line('hi', store.d.name || 'there'))}</h1>
    <p class="lead">${st >= 2 ? esc(line('streak', st)) : due ? esc(line('review', due)) : 'One lesson is about 30 minutes: a short lecture, a demo, a hands-on lab, a quiz and a recap.'}</p>
    <div class="row">${nx ? `<a class="btn big" href="#/lesson/c1/${nx.id}">${store.d.lessons[nx.id] && store.d.lessons[nx.id].status === 'started' ? 'Continue' : 'Start'} lesson ${nx.num}</a>` : `<a class="btn big" href="#/course/c1">Open the course</a>`}${due + errDue ? `<a class="btn ghost" href="#/review">Review ${due + errDue} cards</a>` : ''}</div></div>
    <div class="ringbox">${ring(Math.min(1, t.minutes / goal), 'Today’s minutes')}<small>${Math.round(t.minutes)} of ${goal} min today</small></div></section>
  ${askMissed ? `<section class="card missed"><b>Yesterday was a study day on your plan, and it didn’t happen.</b> What got in the way? <span class="muted">No guilt; this helps you plan better.</span><div class="row">${['Work ran late', 'Too tired', 'Forgot', 'Travel', 'Exams or college', 'Other'].map(r => `<button class="pill btn-pill" data-r="${r}">${r}</button>`).join('')}</div></section>` : ''}
  <div class="grid3">
   <section class="card"><h2>Today’s checklist</h2><ul class="checklist">${items.map(([k, label, auto, href]) => `<li><label><input type="checkbox" data-k="${k}" ${auto || ck[k] ? 'checked' : ''}> <a href="${href}">${esc(label)}</a></label></li>`).join('')}</ul></section>
   <section class="card"><h2>Excel Foundations</h2>${bar(pr.pct)}<p><b>${pr.done}</b> of ${pr.total} lessons done · <b>${pr.tests}</b> of ${c.modules.length} module tests passed</p><p class="countdown">${left > 0 ? `<b>${left}</b> lesson${left === 1 ? '' : 's'} to go before the course project and final exam` : 'All lessons done. The project and final exam are open.'}</p><a href="#/course/c1">Open the syllabus →</a></section>
   <section class="card"><h2>Your streak</h2><p class="bignum">🔥 ${st}<small> day${st === 1 ? '' : 's'}</small></p><p class="muted">Best ${store.bestStreak()} · ${store.freezesLeft()} streak freeze${store.freezesLeft() === 1 ? '' : 's'} left this month (free, never sold)</p>${!store.studied(y) && store.freezesLeft() > 0 && store.streak() === 0 && Object.keys(store.d.activity).length ? '<button class="btn ghost sm" id="freeze">Use a freeze for yesterday</button>' : ''}</section>
  </div>
  <section class="card"><div class="row between"><h2>Study calendar</h2><a href="#/insights">All insights →</a></div>${heatmap(store.d.activity, 20)}<div class="legend"><span>Less</span>${[0, 1, 2, 3, 4].map(i => `<i class="hm${i}"></i>`).join('')}<span>More</span></div></section>
  <section class="card"><h2>Milestones</h2><div class="badges">${MILESTONES.map(([k, label, test]) => { const got = test(store.d); return `<div class="badge ${got ? 'got' : ''}" title="${got ? 'Earned' : 'Not yet'}"><span>${got ? '★' : '☆'}</span>${esc(label)}</div>`; }).join('')}</div></section>
  ${session.signedIn ? '' : `<section class="card free-note">${mascot('thinking', 'sm')}<p><b>You’re using free access.</b> Modules 1–2, the review, error log, planner, journal, function library, shortcuts and playground are all yours. Students with a login from their teacher get Modules 3–8, the project, the final exam and the certificate. <a href="#/login">Student sign-in</a></p></section>`}`;
  $$('[data-k]').forEach(cb => cb.onchange = () => { const d = store.d.checklist[today()] || (store.d.checklist[today()] = {}); d[cb.dataset.k] = cb.checked; store.save(); });
  $$('[data-r]').forEach(b => b.onclick = () => { store.d.plan.missed[y] = b.dataset.r; store.save(); toast('Noted. The planner will show your patterns.'); b.closest('.missed').remove(); });
  const fz = $('#freeze'); if (fz) fz.onclick = () => { if (store.useFreeze()) { toast('Freeze used. Your streak is safe.'); go('#/home'); } };
});

register('courses', async main => {
  const cat = app.catalog; main.innerHTML = `<h1>Courses</h1><p class="lead">Two core courses, six pathways that answer “Excel for what?”, and the Master Certificate for finishing all eight.</p>
  <div class="courses">${cat.courses.map(c => `<a class="card course-card ${c.status}" href="${c.status === 'live' ? '#/course/' + c.id : '#/courses'}" ${c.status !== 'live' ? 'aria-disabled="true"' : ''}><span class="code">${esc(c.code)}</span><h2>${esc(c.title)}</h2><p>${esc(c.summary)}</p><p class="meta">${c.modules} modules · ${c.lessons} lessons · about ${c.hours} hours</p><span class="tag">${c.status === 'live' ? (c.free ? 'Live · Modules 1–2 free' : 'Live') : 'Coming next'}</span>${c.status === 'live' ? bar(courseProgressQuick(c.id)) : ''}</a>`).join('')}
  <div class="card course-card master"><span class="code">★</span><h2>Master Certificate in Excel</h2><p>Opens once all eight course certificates are earned: a capstone business case and a three-hour exam.</p><p class="meta">about 15 hours</p><span class="tag">Opens when all pathways are live</span></div></div>`;
});
function courseProgressQuick(id) { const ls = Object.entries(store.d.lessons).filter(([k, v]) => v.status === 'done').length; return id === 'c1' ? ls / 28 : 0; }

register('course', async (main, id = 'c1') => {
  const c = await getCourse(id); const pr = courseProgress(c); const nx = nextLesson(c);
  const unlocked = session.has(c.id); const allDone = pr.done === pr.total;
  const mt = m => store.d.tests.filter(t => t.kind === 'module' && t.module === c.id + '-m' + m.n); const passed = m => mt(m).some(t => t.passed);
  const proj = store.d.projects[c.id]; const cert = store.d.certs.find(x => x.course === c.id);
  main.innerHTML = `
  <section class="card course-head"><div class="grow"><p class="eyebrow">${esc(c.code)} · ${esc(c.level)} · about ${c.hours} hours</p><h1>${esc(c.title)}</h1><p class="lead">${esc(c.summary)}</p>${bar(pr.pct)}<p>${pr.done}/${pr.total} lessons · ${pr.tests}/${c.modules.length} module tests${cert ? ' · <b>certificate earned</b>' : ''}</p>
   <div class="row">${nx ? `<a class="btn" href="#/lesson/${c.id}/${nx.id}">${store.d.lessons[nx.id] && store.d.lessons[nx.id].status === 'started' ? 'Continue' : 'Start'} ${nx.num} · ${esc(nx.title)}</a>` : ''}${c.downloads.map(d => `<button class="btn ghost" data-dl="${d.file}" data-name="${esc(d.name)}">⬇ ${esc(d.label)}</button>`).join('')}</div></div>${mascot('teaching')}</section>
  <section class="card"><h2>By the end, you can</h2><ol class="outcomes">${c.outcomes.map(o => `<li>${esc(o)}</li>`).join('')}</ol></section>
  ${c.modules.map(m => { const open = m.free || unlocked; const mdone = m.lessons.filter(l => isDone(l.id)).length; return `
   <section class="card module ${open ? '' : 'is-locked'}"><div class="row between"><div><p class="eyebrow">Module ${m.n}${m.free ? ' · Free' : ''}</p><h2>${esc(m.title)}</h2><p class="muted">${esc(m.summary)}</p></div><div class="mod-stat">${mdone}/${m.lessons.length}${open ? '' : ' 🔒'}</div></div>
   <ol class="lessons">${m.lessons.map(l => { const s = store.d.lessons[l.id]; const st = s ? s.status : 'new'; const can = canOpen(c, l); return `<li class="${st}"><a href="${can ? `#/lesson/${c.id}/${l.id}` : '#/login'}"><span class="ls ${st}">${st === 'done' ? '✓' : st === 'started' ? '◐' : can ? '' : '🔒'}</span><span class="ln">${l.num}</span><span class="lt">${esc(l.title)}<small>${esc(l.summary)}</small></span><span class="lm">${l.minutes} min</span></a></li>`; }).join('')}
    <li class="test ${passed(m) ? 'done' : ''}"><a href="${open ? `#/mtest/${c.id}/${m.n}` : '#/login'}"><span class="ls">${passed(m) ? '✓' : '✎'}</span><span class="ln">Test</span><span class="lt">Module ${m.n} test<small>${m.testInfo} · pass mark 70%${mt(m).length ? ` · best ${Math.max(...mt(m).map(t => t.pct))}%` : ''}</small></span><span class="lm">25 min</span></a></li></ol></section>`; }).join('')}
  <section class="card module ${unlocked ? '' : 'is-locked'}"><p class="eyebrow">Course project · 3 hours</p><h2>${esc(c.project.title)}</h2><p>${esc(c.project.brief)}</p><p class="muted">Pass mark 70% · two submissions · counts 40% of the course grade${proj && proj.best !== undefined ? ` · your best ${proj.best}%` : ''}</p><a class="btn ${unlocked ? '' : 'ghost'}" href="${unlocked ? `#/project/${c.id}` : '#/login'}">${unlocked ? 'Open the project' : '🔒 Sign in to open'}</a></section>
  <section class="card module ${unlocked ? '' : 'is-locked'}"><p class="eyebrow">Final exam · ${c.exam.minutes} minutes</p><h2>${esc(c.code)} final exam</h2><p>${c.exam.questions} questions and ${c.exam.tasks} practical tasks on a live grid, timed. Pass mark ${c.exam.pass}%. Up to ${c.exam.attempts} attempts, ${c.exam.gapHours} hours apart, each with new questions. It opens once every lesson and module test is done and the project is passed.</p>
   <a class="btn ${unlocked ? '' : 'ghost'}" href="${unlocked ? `#/exam/${c.id}` : '#/login'}">${unlocked ? 'Exam details' : '🔒 Sign in to open'}</a></section>`;
  $$('[data-dl]').forEach(b => b.onclick = async () => { const url = await assetURL(c.id, b.dataset.dl); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); });
});
