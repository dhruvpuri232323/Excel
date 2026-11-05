// Module tests (and testing out), practice tests, the course project, the final exam and certificates.
import { register, getCourse, go, lockedPanel, setFocus } from '../app.js';
import { store, today } from '../store.js';
import { session, lesson as loadLesson, getSecure, assetURL } from '../content.js';
import { esc, md, $, $$, mascot, toast, confetti, bar, shuffle, modal, lineChart, fmtDate, confirmBox, crumbs, qhelp, download } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask } from '../checker.js';
import { readXlsx, grade } from '../xlsxcheck.js';
import { lessonList, isDone, moduleInfo, canTakeTest, canOpenLesson, hasKey, who, testsFor, projectOpen, examReqs, lastModule, lessonLabel, TEST_MINUTES, fmtMin } from '../access.js';
import { resultsHTML, hydrate } from './lesson.js';
import { site, loadSite } from '../site.js';
import { askTeacher } from '../help.js';

export async function questionsFrom(c, lessons) { const out = []; for (const l of lessons) { try { const L = await loadLesson(c, l.id); (L.quiz || []).forEach(q => out.push({ ...q, from: L.num, lesson: l.id, lessonTitle: L.title })); } catch (e) { } } return out; }
export async function labTasksFrom(c, lessons) { const out = []; for (const l of lessons) { try { const L = await loadLesson(c, l.id); for (const t of (L.lab && L.lab.tasks) || []) if (t.standalone) out.push({ sheet: L.lab.sheet, editable: L.lab.editable, inputs: L.lab.inputs, task: t, from: L.num, lesson: l.id, lessonTitle: L.title, height: 300 }); } catch (e) { } } return out; }
export const shuffleOpts = q => { if (!q.options) return q; const idx = shuffle(q.options.map((_, i) => i)); return { ...q, options: idx.map(i => q.options[i]), answer: idx.indexOf(q.answer) }; };

// A timed test: questions, then practical tasks on a live grid. Each task counts as two questions. Tests open in focus mode.
export function runTest(main, { title, subtitle, questions, tasks = [], minutes, pass, onFinish, kind, after, record = true, course = 'c1' }) {
  const answers = {}, flags = new Set(); let i = 0; const n = questions.length + tasks.length; const ends = Date.now() + minutes * 60000; const grids = {}; let timer; let finished = false;
  const pages = [...questions.map((q, k) => ({ q, k })), ...tasks.map((t, k) => ({ t, k }))];
  setFocus(true);
  const leave = e => { if (!finished) { e.preventDefault(); e.returnValue = ''; } }; addEventListener('beforeunload', leave);
  const draw = () => {
    const p = pages[i];
    main.innerHTML = `<div class="test"><div class="test-head card"><div><p class="eyebrow">${esc(subtitle || '')}</p><h1>${esc(title)}</h1></div><div class="timer" id="tm" role="timer"></div></div>
     <div class="qnav" aria-label="Questions">${pages.map((pp, k) => `<button class="qn ${k === i ? 'on' : ''} ${answered(k) ? 'ans' : ''} ${flags.has(k) ? 'flag' : ''}" data-k="${k}" aria-label="${pp.t ? 'Task ' + (pp.k + 1) : 'Question ' + (k + 1)}">${pp.t ? 'T' + (pp.k + 1) : k + 1}</button>`).join('')}</div>
     <div class="card tpage" id="tp"></div>
     <div class="row between"><button class="btn ghost" id="pv" ${i ? '' : 'disabled'}>← Previous</button><button class="btn ghost" id="fl">${flags.has(i) ? '⚑ Flagged' : '⚐ Flag for later'}</button>${i < n - 1 ? '<button class="btn" id="nx">Next →</button>' : '<button class="btn" id="sub">Submit</button>'}</div></div>`;
    const tp = $('#tp');
    if (p.q) { const q = p.q; tp.innerHTML = `<p class="eyebrow">Question ${i + 1} of ${questions.length}</p><h2>${md(q.q)}</h2>${q.code ? `<p><code class="big">${esc(q.code)}</code></p>` : ''}${q.grid ? `<div class="mini-grid" data-grid='${esc(JSON.stringify(q.grid))}'></div>` : ''}${q.answers ? `<input class="ans big-input" value="${esc(answers[i] ?? '')}" placeholder="Type your answer" spellcheck="false">` : `<div class="opts">${q.options.map((o, k) => `<button class="opt ${answers[i] === k ? 'sel' : ''}" data-o="${k}">${md(o)}</button>`).join('')}</div>`}`;
      hydrate(tp, course); $$('.opt', tp).forEach(b => b.onclick = () => { answers[i] = +b.dataset.o; $$('.opt', tp).forEach(x => x.classList.toggle('sel', x === b)); markNav(); }); const inp = $('.ans', tp); if (inp) inp.oninput = () => { answers[i] = inp.value; markNav(); };
    } else { const t = p.t; tp.innerHTML = `<p class="eyebrow">Practical task ${p.k + 1} of ${tasks.length}${t.lessonTitle ? ' · from ' + esc(t.lessonTitle) : ''}</p><h2>${md(t.task.prompt)}</h2><div class="tg"></div><p class="muted small">Your work here is kept as you move between questions. It’s checked when you submit.</p>`;
      if (!grids[i]) { const host = document.createElement('div'); grids[i] = new Grid(host, new Workbook(t.sheet), { height: t.height || 300, editable: t.editable, highlight: t.editable, allowSort: true, allowStructure: t.task.kind === 'noBlankRows', onChange: () => { answers[i] = true; markNav(); } }); }
      $('.tg', tp).appendChild(grids[i].host); grids[i].render(); }
    $$('.qn').forEach(b => b.onclick = () => { i = +b.dataset.k; draw(); }); $('#pv').onclick = () => { i--; draw(); }; const nx = $('#nx'); if (nx) nx.onclick = () => { i++; draw(); };
    $('#fl').onclick = () => { flags.has(i) ? flags.delete(i) : flags.add(i); draw(); }; const sb = $('#sub'); if (sb) sb.onclick = async () => { const blank = pages.filter((pp, k) => !answered(k)).length; if (blank && !await confirmBox(`${blank} question${blank === 1 ? ' is' : 's are'} unanswered. Submit anyway?`, 'Submit', 'Keep going')) return; finish(); };
    tick();
  };
  const answered = k => answers[k] !== undefined && answers[k] !== '';
  const markNav = () => $$('.qn').forEach((b, k) => b.classList.toggle('ans', answered(k)));
  const tick = () => { clearInterval(timer); const el = $('#tm'); const up = () => { const left = Math.max(0, ends - Date.now()); const m = Math.floor(left / 60000), s = Math.floor(left / 1000) % 60; if (!document.body.contains(el)) { clearInterval(timer); return; } el.textContent = `${m}:${String(s).padStart(2, '0')} left`; el.classList.toggle('low', left < 5 * 60000); if (!left) { clearInterval(timer); toast('Time is up. Your answers have been submitted.'); finish(); } }; up(); timer = setInterval(up, 1000); };
  const finish = () => {
    finished = true; removeEventListener('beforeunload', leave); clearInterval(timer); setFocus(false);
    let qRight = 0, tRight = 0; const review = []; const byLesson = {};
    const tally = (lid, ok, w = 1) => { if (!lid) return; const b = byLesson[lid] || (byLesson[lid] = [0, 0]); b[1] += w; if (ok) b[0] += w; };
    questions.forEach((q, k) => { const a = answers[k]; const ok = q.answers ? !!(a && q.answers.some(x => x.replace(/\s/g, '').toUpperCase() === String(a).replace(/\s/g, '').toUpperCase())) : a === q.answer; if (ok) qRight++; tally(q.lesson, ok); review.push({ q, ok, given: q.answers ? a : (a !== undefined ? q.options[a] : '(blank)') }); if (!ok && record && kind !== 'placement') store.logError({ course, lesson: q.lesson, lessonNum: q.from, task: 'test', kind, prompt: q.q, yours: review.at(-1).given || '(blank)', right: q.answers ? q.answers[0] : q.options[q.answer], why: q.explain }); });
    tasks.forEach((t, k) => { const g = grids[questions.length + k]; const res = g ? checkTask(t.task, { wb: g.wb, grid: g, inputs: t.inputs }) : { ok: false, msg: 'Not attempted.' }; if (res.ok) tRight++; tally(t.lesson, res.ok, 2); review.push({ task: t.task, ok: res.ok, msg: res.msg, lesson: t.lesson }); });
    const total = questions.length + tasks.length * 2; const score = qRight + tRight * 2; const pct = total ? Math.round(score / total * 100) : 0; const passed = pct >= pass;
    const revisit = Object.entries(byLesson).filter(([, [r, nn]]) => r < nn).sort((a, b) => (a[1][0] / a[1][1]) - (b[1][0] / b[1][1])).slice(0, 3).map(([l]) => l);
    const result = { kind, at: new Date().toISOString(), pct, qRight, nq: questions.length, tRight, nt: tasks.length, passed, minutes: Math.round((minutes * 60000 - Math.max(0, ends - Date.now())) / 60000), byLesson, revisit };
    if (onFinish) onFinish(result); store.save();
    if (after) return after({ ...result, review });
    main.innerHTML = resultPage(title, result, review, pass);
    if (passed) confetti(); scrollTo(0, 0);
  };
  draw();
}
const titleOf = (c, id) => { for (const m of c.modules) for (const l of m.lessons) if (l.id === id) return { l, m }; return null; };
let COURSE = null;
// the results card (X9): what went up or down since last time, and the 2-3 lessons worth revisiting
function resultPage(title, r, review, pass, prev, extra = '') {
  const c = COURSE; const name = id => { const x = c && titleOf(c, id); return x ? `${lessonLabel(x.l, x.m)} · ${esc(x.l.title)}` : esc(id); };
  const moves = prev && prev.byLesson ? Object.keys(r.byLesson).map(id => { const a = r.byLesson[id], b = prev.byLesson[id]; if (!b) return null; const d = Math.round(a[0] / a[1] * 100) - Math.round(b[0] / b[1] * 100); return d ? [id, d] : null; }).filter(Boolean) : [];
  return `<div class="card test-end">${mascot(r.passed ? 'celebrate' : 'encourage')}<div class="grow"><p class="eyebrow">${esc(title)}</p><h1>${r.pct}%${pass ? ` · ${r.passed ? 'Passed' : 'Not passed yet'}` : ''}</h1><p class="lead">${r.qRight} of ${r.nq} questions${r.nt ? ` and ${r.tRight} of ${r.nt} practical tasks` : ''} right.${pass ? ` The pass mark is ${pass}%.` : ''}</p>${extra}</div></div>
   <div class="grid2"><section class="card"><h2>Worth revisiting</h2>${r.revisit.length ? `<ul class="revisit">${r.revisit.map(id => { const x = c && titleOf(c, id); const b = r.byLesson[id]; return `<li><a href="${x ? `#/lesson/${c.id}/${id}/lecture` : '#/home'}">${name(id)}</a><small>${b[0]} of ${b[1]} right</small></li>`; }).join('')}</ul><p class="muted small">Every question you missed is also in Mistakes and comes back in your daily review.</p>` : '<p>Nothing: every lesson in this test went well.</p>'}</section>
   <section class="card"><h2>Since last time</h2>${moves.length ? `<ul class="moves">${moves.map(([id, d]) => `<li class="${d > 0 ? 'up' : 'down'}"><span>${d > 0 ? '▲' : '▼'} ${Math.abs(d)} points</span> ${name(id)}</li>`).join('')}</ul>` : `<p class="muted">${prev ? 'No change by lesson since your last attempt.' : 'This is your first attempt. Next time, this shows which topics went up or down.'}</p>`}</section></div>
   <section class="card"><details><summary><b>All answers</b></summary><ol class="review-list">${review.map(x => x.q ? `<li class="${x.ok ? 'ok' : 'no'}"><p>${md(x.q.q)}</p><p class="small">${x.ok ? '✓' : '✗ You: ' + esc(x.given ?? '(blank)') + ' · Right: ' + esc(x.q.answers ? x.q.answers[0] : x.q.options[x.q.answer])}</p><p class="muted small">${md(x.q.explain || '')}</p></li>` : `<li class="${x.ok ? 'ok' : 'no'}"><p><b>Task:</b> ${md(x.task.prompt)}</p><p class="small">${x.ok ? '✓ Correct' : '✗ ' + esc(x.msg)}</p></li>`).join('')}</ol></details></section>`;
}
// retakes (U6): 24 hours after a failed attempt, or straight away once the suggested lessons are redone
export function retakeState(c, m) {
  const tries = testsFor(c, m.n); const last = tries.at(-1); if (!last || last.passed) return { open: true };
  const wait = Math.max(0, Date.parse(last.at) + 24 * 3600e3 - Date.now());
  const redo = (last.revisit || []).filter(id => { const s = store.d.lessons[id]; return !(s && s.quizAt && s.quizAt > last.at); });
  return { open: !wait || !redo.length, wait, redo, last };
}

register('mtest', async (main, courseId, n) => {
  const c = await getCourse(courseId); COURSE = c; const m = c.modules.find(x => x.n === +n); if (!m) throw new Error('No such module');
  if (!canTakeTest(c, m)) { main.innerHTML = lockedPanel(); return; }
  const info = moduleInfo(c, m); const tries = testsFor(c, m.n); const passed = info.passed; const rt = retakeState(c, m); const testOut = !info.open;
  main.innerHTML = `${crumbs([[c.title, '#/course/' + c.id], [`Module ${m.n}`, `#/module/${c.id}/${m.n}`], ['Module test']])}
  <section class="card"><div class="row between"><div><p class="eyebrow">Module ${m.n} · about ${TEST_MINUTES} minutes ${qhelp('testrules')}</p><h1>${testOut ? `Test out of Module ${m.n}: ${esc(m.title)}` : `Module ${m.n} test: ${esc(m.title)}`}</h1></div>${mascot('focus')}</div>
   <p class="lead">${testOut ? `Already know this module? Pass this test (70% or more) and Sheetah marks Module ${m.n} as <b>tested out</b> and opens the next module. The lessons stay open for review. ${qhelp('testout')}` : `${esc(m.testInfo)}. 15 questions and 2 practical tasks, pass mark 70%. Questions come from every lesson in this module, in a new order each time.`}</p>
   <ul class="facts"><li>${info.done}/${info.total} lessons in this module done${info.done < info.total && !testOut ? ' <span class="muted">(you can take the test whenever you feel ready)</span>' : ''}</li><li>${tries.length ? `Attempts so far: ${tries.length} · best ${Math.max(...tries.map(t => t.pct))}%${passed ? ' · passed ✓' : ''}` : 'No attempts yet'}</li>
   ${!rt.open ? `<li>The retake opens in ${Math.ceil(rt.wait / 3600e3)} hour(s), or straight away once you redo: ${rt.redo.map(id => { const x = titleOf(c, id); return `<a href="#/lesson/${c.id}/${id}/quiz">${x ? esc(x.l.title) : id}</a>`; }).join(', ')} (finish each lesson’s quiz again).</li>` : ''}</ul>
   <div class="row"><button class="btn big" id="go" ${rt.open || who() === 'admin' ? '' : 'disabled'}>${passed ? 'Take it again for practice' : testOut ? 'Start the test-out' : 'Start the test'}</button><a class="btn ghost" href="#/module/${c.id}/${m.n}">Back</a></div></section>`;
  $('#go').onclick = async () => {
    const qs = shuffle(await questionsFrom(c, m.lessons)).slice(0, 15).map(shuffleOpts); const ts = shuffle(await labTasksFrom(c, m.lessons)).slice(0, 2);
    const prev = tries.at(-1); const idx = c.modules.indexOf(m); const nextM = c.modules[idx + 1];
    runTest(main, { title: `Module ${m.n} test`, subtitle: `${c.title} · ${m.title}`, questions: qs, tasks: ts, minutes: TEST_MINUTES, pass: 70, kind: 'module', course: c.id,
      onFinish: r => { store.d.tests.push({ ...r, module: `${c.id}-m${m.n}`, testOut }); },
      after: r => {
        let extra = '';
        if (r.passed && nextM) extra = who() === 'visitor' && !nextM.free ? `<div class="tip">${mascot('celebrate', 'xs')}<p><b>Great result.</b> Ask your teacher for a login to continue with Module ${nextM.n}: ${esc(nextM.title)}. <button class="btn sm" type="button" data-ask="login">Ask for a login</button></p></div>` : `<p class="ok">Module ${nextM.n}: ${esc(nextM.title)} is open. <a class="btn sm" href="#/module/${c.id}/${nextM.n}">Go to Module ${nextM.n}</a></p>`;
        if (r.passed && !nextM) extra = `<p class="ok">That was the last module test. The course project is open. <a class="btn sm" href="#/project/${c.id}">Open the project</a></p>`;
        if (!r.passed) extra = `<p>The retake opens in 24 hours, or straight away once you redo the lessons below (finish each one’s quiz again).</p>`;
        main.innerHTML = resultPage(`Module ${m.n} test`, r, r.review, 70, prev, extra) + `<div class="row"><a class="btn" href="#/course/${c.id}">Back to the map</a><a class="btn ghost" href="#/errors">Open Mistakes</a></div>`; if (r.passed) confetti(); scrollTo(0, 0);
      } });
  };
});

register('tests', async main => {
  const c = await getCourse('c1'); COURSE = c; const open = lessonList(c).filter(l => canOpenLesson(c, l)); const doneL = open.filter(l => isDone(l.id));
  const hist = store.d.tests.filter(t => t.kind === 'mock'); const target = store.d.mockTarget || 75;
  const weak = {}; store.d.errors.filter(e => e.kind === 'mock').forEach(e => { weak[e.lesson] = (weak[e.lesson] || 0) + 1; });
  main.innerHTML = `<h1>Tests</h1><p class="lead">Practice tests mix questions from lessons you’ve done, which is how real exams feel. Module tests open the next module. All results stay on this laptop.</p>
  <section class="card"><div class="row between"><div><h2>Practice test</h2><p class="muted">${doneL.length >= 2 ? `Drawn from your ${doneL.length} finished lessons` : doneL.length ? 'Finish one more lesson for a mixed test. For now it uses every lesson open to you.' : 'Finish a lesson or two first: practice tests work best on what you’ve already learned.'}</p></div>${mascot('focus', 'sm')}</div>
   <div class="row"><button class="btn" data-n="10" ${open.length ? '' : 'disabled'}>10 questions · 10 min</button><button class="btn" data-n="20" ${open.length ? '' : 'disabled'}>20 questions · 20 min</button><button class="btn ghost" data-n="30" ${open.length ? '' : 'disabled'}>30 questions · 30 min</button></div>
   <label class="inline">Your target score <input type="number" id="tg" min="40" max="100" value="${target}">%</label></section>
  <section class="card"><h2>Module tests</h2><ul class="mtests">${c.modules.map(m => { const inf = moduleInfo(c, m); const best = testsFor(c, m.n); return `<li class="${inf.passed ? 'done' : ''}"><a href="#/mtest/${c.id}/${m.n}"><span class="ls">${inf.passed ? '✓' : inf.open ? '✎' : '🔒'}</span><span><b>Module ${m.n}: ${esc(m.title)}</b><small>${inf.passed ? (inf.testedOut ? 'Tested out' : 'Passed') : inf.open ? 'Open' : inf.reason === 'login' ? 'Needs a login' : 'Can be taken to test out'}${best.length ? ` · best ${Math.max(...best.map(t => t.pct))}%` : ''}</small></span><span class="lm">${TEST_MINUTES} min</span></a></li>`; }).join('')}</ul>
   <div class="row"><a class="btn ghost" href="#/placement/c1">Placement check (10 min)</a><a class="btn ghost" href="#/interview">Interview prep mock</a></div></section>
  <div class="grid2"><section class="card"><h2>Score trend</h2>${lineChart(hist.slice(-20).map(t => ({ k: fmtDate(t.at), v: t.pct })), { target, label: 'Practice test scores' })}</section>
  <section class="card"><h2>Weakest lessons in tests</h2>${Object.keys(weak).length ? `<ul class="weak">${Object.entries(weak).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([k, v]) => { const x = titleOf(c, k); return `<li><a href="#/lesson/c1/${k}">${x ? esc(x.l.title) : esc(k)}</a> · ${v} missed</li>`; }).join('')}</ul>` : '<p class="muted">After a test, the lessons you miss most show up here.</p>'}</section></div>
  <section class="card"><h2>Score sheet</h2>${hist.length ? `<table class="tbl"><thead><tr><th>Date</th><th>Questions</th><th>Score</th><th>Time</th></tr></thead><tbody>${hist.slice().reverse().map(t => `<tr><td>${fmtDate(t.at)}</td><td>${t.nq}</td><td><b>${t.pct}%</b> ${t.pct >= target ? '✓' : ''}</td><td>${t.minutes} min</td></tr>`).join('')}</tbody></table>` : '<p class="muted">No practice tests yet.</p>'}</section>`;
  $('#tg').onchange = e => { store.d.mockTarget = +e.target.value; store.save(); };
  $$('[data-n]').forEach(b => b.onclick = async () => { const n = +b.dataset.n; const pool = await questionsFrom(c, doneL.length >= 2 ? doneL : open); if (!pool.length) return toast('Open a lesson first.'); const prev = hist.at(-1);
    runTest(main, { title: `Practice test · ${n} questions`, subtitle: 'Mixed practice', questions: shuffle(pool).slice(0, n).map(shuffleOpts), minutes: n, pass: target, kind: 'mock', onFinish: r => store.d.tests.push(r), after: r => { main.innerHTML = resultPage(`Practice test · ${n} questions`, r, r.review, target, prev) + '<div class="row"><a class="btn" href="#/tests">Back to tests</a></div>'; if (r.passed) confetti(); scrollTo(0, 0); } }); });
});

// ---------- project ----------
register('project', async (main, courseId) => {
  const c = await getCourse(courseId); COURSE = c;
  if (!hasKey(c)) { main.innerHTML = lockedPanel(); return; }
  if (!projectOpen(c)) { const lm = lastModule(c); main.innerHTML = lockedPanel('test', `<h2>The project opens after the Module ${lm.n} test</h2><p>Pass the Module ${lm.n} test (${esc(lm.title)}) and the course project opens. It’s a real workbook you build in Excel and upload: about 3 hours, pass mark 70%.</p><div class="row"><a class="btn" href="#/course/${c.id}">See the map</a></div>`); return; }
  const P = await getSecure(c.id, 'project'); const ps = store.d.projects[c.id] || (store.d.projects[c.id] = { tries: [] });
  const left = 2 - ps.tries.length;
  main.innerHTML = `${crumbs([[c.title, '#/course/' + c.id], ['Course project']])}<section class="card"><div class="row between"><div><p class="eyebrow">Course project · about 3 hours · pass mark 70% ${qhelp('project')}</p><h1>${esc(P.title)}</h1></div>${mascot('focus')}</div><p class="lead">${md(P.brief)}</p>
   <h2>What to build</h2><ol class="steps">${P.steps.map(s => `<li>${md(s)}</li>`).join('')}</ol>
   <h2>How it’s checked</h2><ul class="ticks">${P.rubricText.map(s => `<li>${md(s)}</li>`).join('')}</ul>
   <div class="row">${P.downloads.map(d => `<button class="btn ghost" data-file="${d.file}" data-name="${esc(d.name)}">⬇ ${esc(d.label)}</button>`).join('')}</div>
   <h2>Submit</h2><p>${left > 0 ? `You have <b>${left}</b> submission${left === 1 ? '' : 's'} left. Your best score counts.` : 'You have used both submissions.'}${ps.best !== undefined ? ` Best so far: <b>${ps.best}%</b>${ps.best >= 70 ? ' · passed ✓' : ''}.` : ''}</p>
   ${left > 0 || who() === 'admin' ? `<label class="btn">⬆ Upload my workbook (.xlsx)<input type="file" accept=".xlsx,.xlsm,.xls,.csv,.ods,.numbers" hidden id="up"></label> <a class="small" href="#/help/upload">Problems uploading?</a>` : ''}<div class="results">${ps.last ? resultsHTML(ps.last) : ''}</div>
   ${ps.tries.length ? `<h3>Your submissions</h3><ul>${ps.tries.map(t => `<li>${fmtDate(t.at)}: <b>${t.pct}%</b></li>`).join('')}</ul>` : ''}</section>`;
  $$('[data-file]').forEach(b => b.onclick = async () => { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); });
  const up = $('#up'); if (up) up.onchange = async e => { const f = e.target.files[0]; if (!f) return; const x = await readXlsx(f); if (x.error) { $('.results').innerHTML = `<p class="no">${esc(x.error)}</p>`; e.target.value = ''; return; } if (!await confirmBox(`Submit “${esc(f.name)}” as submission ${ps.tries.length + 1} of 2?`, 'Submit')) return; $('.results').innerHTML = '<p>Checking…</p>'; const res = grade(x, P.checks); ps.tries.push({ at: new Date().toISOString(), pct: res.pct }); ps.last = res; ps.best = Math.max(ps.best ?? 0, res.pct); store.save(); if (res.pct >= 70) confetti(); go('#/project/' + c.id); };
});

// ---------- final exam ----------
register('exam', async (main, courseId) => {
  const c = await getCourse(courseId); COURSE = c; if (!hasKey(c)) { main.innerHTML = lockedPanel(); return; }
  const E = c.exam; const att = store.d.examAttempts[c.id] || (store.d.examAttempts[c.id] = []); const passed = att.some(a => a.passed);
  const R = examReqs(c); const last = att.at(-1); const wait = last && !last.passed ? Math.max(0, Date.parse(last.at) + E.gapHours * 3600e3 - Date.now()) : 0;
  const ready = R.all && !passed && att.length < E.attempts && !wait;
  main.innerHTML = `${crumbs([[c.title, '#/course/' + c.id], ['Final exam']])}<section class="card"><div class="row between"><div><p class="eyebrow">Final exam ${qhelp('exam')}</p><h1>${esc(c.title)}: final exam</h1></div>${mascot('focus')}</div>
   <p class="lead">${E.minutes} minutes · ${E.questions} questions and ${E.tasks} practical tasks · pass mark ${E.pass}% · up to ${E.attempts} attempts, ${E.gapHours} hours apart, each with new questions and data.</p>
   <ul class="reqs"><li class="${R.lessonsOk ? 'ok' : ''}">${R.lessonsOk ? '✓' : '○'} Every lesson done (or its module tested out)</li><li class="${R.modsOk ? 'ok' : ''}">${R.modsOk ? '✓' : '○'} Every module test passed</li><li class="${R.projOk ? 'ok' : ''}">${R.projOk ? '✓' : '○'} Course project passed${R.proj && R.proj.best !== undefined ? ` (best ${R.proj.best}%)` : ''}</li></ul>
   ${att.length ? `<p>Attempts: ${att.map(a => `${fmtDate(a.at)} <b>${a.pct}%</b>`).join(' · ')}</p>` : ''}${wait ? `<p>Your next attempt opens in ${Math.ceil(wait / 3600e3)} hours.</p>` : ''}${passed ? '<p><b>You passed.</b> Your certificate is on the Certificates page.</p>' : ''}
   <p class="muted small">Before you start: use a laptop, close other tabs and keep 90 minutes free. The exam opens in focus mode and follows an honour code: work alone, with no help from others.</p>
   <div class="row"><button class="btn big" id="go" ${ready ? '' : 'disabled'}>Start the exam</button>${who() === 'admin' ? '<button class="btn ghost" id="prev">Admin preview (doesn’t count)</button>' : ''}<a class="btn ghost" href="#/course/${c.id}">Back</a></div></section>`;
  const start = async preview => {
    if (!preview && !await confirmBox(`Start the exam now? The ${E.minutes}-minute timer can’t be paused.`, 'Start')) return;
    const bank = await getSecure(c.id, 'exam'); const qs = shuffle(bank.questions).slice(0, E.questions).map(shuffleOpts); const ts = shuffle(bank.tasks).slice(0, E.tasks);
    runTest(main, { title: `${c.title} final exam`, subtitle: preview ? 'Admin preview · not recorded' : `Attempt ${att.length + 1} of ${E.attempts}`, questions: qs, tasks: ts, minutes: E.minutes, pass: E.pass, kind: 'exam', course: c.id, record: !preview,
      onFinish: r => { if (preview) return; att.push(r); if (r.passed) award(c, r.pct); },
      after: r => { main.innerHTML = resultPage(`${c.title} final exam`, r, r.review, E.pass, att.at(-2), r.passed && !preview ? `<p class="ok">Your certificate is ready. <a class="btn sm" href="#/certificates">See it</a></p>` : '') + `<div class="row"><a class="btn" href="#/course/${c.id}">Back to the map</a></div>`; if (r.passed) confetti(); } });
  };
  $('#go').onclick = () => start(false); const pv = $('#prev'); if (pv) pv.onclick = () => start(true);
});
function award(c, examPct) {
  const proj = store.d.projects[c.id].best; const gradePct = Math.round(examPct * .6 + proj * .4); const id = 'SH-' + c.code + '-' + new Date().getFullYear() + '-' + Math.random().toString(36).slice(2, 7).toUpperCase();
  store.d.certs.push({ course: c.id, code: c.code, title: c.title, id, at: new Date().toISOString(), grade: gradePct, distinction: gradePct >= 90, name: (session.s && session.s.name) || store.d.name, skills: c.outcomes.slice(0, 4), avatar: store.d.profile.avatar });
}

// ---------- certificates (X6) ----------
const verifyURL = id => location.origin + location.pathname + '#/verify/' + encodeURIComponent(id);
register('certificates', async main => {
  const certs = store.d.certs;
  main.innerHTML = `<h1>Certificates</h1><p class="lead">Each course earns its own certificate once the project and the final exam are passed. Download it as a PDF, add it to LinkedIn, and send its ID to your teacher so anyone can verify it.</p>
  ${certs.length ? certs.map(ct => `<section class="card certrow"><div><h2>${esc(ct.title)}</h2><p>${fmtDate(ct.at)} · grade ${ct.grade}%${ct.distinction ? ' · with Distinction' : ''} · ID ${esc(ct.id)}</p></div><a class="btn" href="#/cert/${ct.id}">Open</a></section>`).join('') : `<section class="card empty">${mascot('thinking')}<div><h2>No certificates yet</h2><p>Excel Foundations: pass the module tests, then the course project and the final exam.</p><a class="btn" href="#/course/c1">See your map</a></div></section>`}`;
});
register('cert', async (main, id) => {
  const ct = store.d.certs.find(x => x.id === id); if (!ct) throw new Error('This certificate isn’t on this laptop. Certificates are saved where they were earned; restore a backup to bring them across.');
  const d = new Date(ct.at); const li = `https://www.linkedin.com/profile/add?startTask=CERTIFICATION_NAME&name=${encodeURIComponent(ct.title)}&organizationName=${encodeURIComponent('Sheetah')}&issueYear=${d.getFullYear()}&issueMonth=${d.getMonth() + 1}&certUrl=${encodeURIComponent(verifyURL(ct.id))}&certId=${encodeURIComponent(ct.id)}`;
  main.innerHTML = `${crumbs([['Certificates', '#/certificates'], [ct.title]])}<div class="row noprint"><button class="btn" id="pdf">⬇ Download PDF</button><button class="btn ghost" onclick="print()">Print</button><a class="btn ghost" href="${li}" target="_blank" rel="noopener">Add to LinkedIn</a><button class="btn ghost" id="cpv">Copy verify link</button><button class="btn ghost" id="snd">Send the ID to my teacher</button></div>
  <div class="certificate" id="certBox">${certHTML(ct)}</div><p class="muted small noprint">Verify link: ${esc(verifyURL(ct.id))}. It shows as verified once your teacher adds the ID to the register.</p>`;
  $('#cpv').onclick = () => { navigator.clipboard && navigator.clipboard.writeText(verifyURL(ct.id)); toast('Verify link copied'); };
  $('#snd').onclick = () => askTeacher({ subject: 'Sheetah certificate', what: `I passed ${ct.title} with ${ct.grade}%${ct.distinction ? ' (Distinction)' : ''}. Certificate ID: ${ct.id}. Could you add it to the certificate register?` });
  $('#pdf').onclick = () => certPDF(ct).catch(e => toast('Couldn’t make the PDF: ' + esc(e.message) + '. Use Print › Save as PDF instead.'));
});
const certHTML = ct => `<div class="cert-inner"><img src="img/mascot/sheetah-celebrate.svg" alt="" class="cert-m"><p class="cert-eyebrow">Certificate of completion</p><h1>${esc(ct.name)}</h1><p>has completed</p><h2>${esc(ct.title)}</h2><p class="grade">Grade ${ct.grade}%${ct.distinction ? ' · with Distinction' : ''}</p>
  <ul class="cert-skills">${(ct.skills || []).map(s => `<li>${esc(s)}</li>`).join('')}</ul><div class="cert-foot"><span>Issued ${fmtDate(ct.at)}</span><span>Certificate ID ${esc(ct.id)}</span></div>
  <p class="cert-legal">Issued by Sheetah. Verify at ${esc(verifyURL(ct.id))}. Not affiliated with or endorsed by Microsoft. Microsoft and Excel are trademarks of the Microsoft group of companies.</p></div>`;
const loadScript = src => new Promise((res, rej) => { if (document.querySelector(`script[src="${src}"]`)) return res(); const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('download failed')); document.head.appendChild(s); });
async function certPDF(ct) {
  toast('Making your PDF…'); await loadScript('vendor/pdf-lib.min.js');
  const W = 1684, H = 1190; const mascotSvg = await (await fetch('img/mascot/sheetah-celebrate.svg')).text(); const mUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(mascotSvg)));
  const x = s => esc(s); const skills = (ct.skills || []).slice(0, 4);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="100%" height="100%" fill="#fff"/><rect x="24" y="24" width="${W - 48}" height="${H - 48}" fill="none" stroke="#23263A" stroke-width="28"/><rect x="62" y="62" width="${W - 124}" height="${H - 124}" fill="none" stroke="#F2A541" stroke-width="6"/>
   <image href="${mUrl}" x="${W / 2 - 90}" y="110" width="180" height="198"/><g font-family="Nunito, Segoe UI, Arial, sans-serif" text-anchor="middle" fill="#23263A">
   <text x="${W / 2}" y="370" font-size="30" letter-spacing="9" font-weight="800" fill="#3557D6">CERTIFICATE OF COMPLETION</text><text x="${W / 2}" y="480" font-size="84" font-weight="800">${x(ct.name)}</text><text x="${W / 2}" y="550" font-size="32" fill="#4A4D62">has completed</text>
   <text x="${W / 2}" y="635" font-size="56" font-weight="800">${x(ct.title)}</text><text x="${W / 2}" y="705" font-size="34" font-weight="800" fill="#9A5B00">Grade ${ct.grade}%${ct.distinction ? ' · with Distinction' : ''}</text>
   ${skills.map((s, k) => `<text x="${W / 2}" y="${775 + k * 40}" font-size="24" fill="#4A4D62">${x(s)}</text>`).join('')}
   <text x="150" y="1030" font-size="26" font-weight="800" text-anchor="start">Issued ${x(fmtDate(ct.at))}</text><text x="${W - 150}" y="1030" font-size="26" font-weight="800" text-anchor="end">Certificate ID ${x(ct.id)}</text>
   <text x="${W / 2}" y="1085" font-size="18" fill="#6E6C64">Issued by Sheetah. Verify at ${x(verifyURL(ct.id))}. Not affiliated with or endorsed by Microsoft. Microsoft and Excel are trademarks of the Microsoft group of companies.</text></g></svg>`;
  const img = new Image(); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg))); await img.decode();
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H; cv.getContext('2d').drawImage(img, 0, 0); const png = await new Promise(r => cv.toBlob(r, 'image/png'));
  const { PDFDocument } = window.PDFLib; const pdf = await PDFDocument.create(); const page = pdf.addPage([842, 595]); const pi = await pdf.embedPng(await png.arrayBuffer()); page.drawImage(pi, { x: 0, y: 0, width: 842, height: 595 });
  pdf.setTitle(`${ct.title} certificate · ${ct.name}`); pdf.setAuthor('Sheetah'); download(`Sheetah certificate - ${ct.title}.pdf`, new Blob([await pdf.save()], { type: 'application/pdf' }));
}
register('verify', async (main, id = '') => {
  await loadSite(); const list = site.data.certificates || []; const hit = list.find(x => x.id.toUpperCase() === String(id).trim().toUpperCase());
  main.innerHTML = `<section class="card verify">${mascot(hit ? 'celebrate' : 'thinking')}<div><p class="eyebrow">Certificate check</p>${hit ? `<h1>✓ Verified</h1><p class="lead"><b>${esc(hit.name)}</b> earned <b>${esc(hit.title || 'Excel Foundations')}</b>${hit.grade ? ` with a grade of ${hit.grade}%` : ''}${hit.date ? `, issued ${fmtDate(hit.date)}` : ''}.</p><p class="muted">Certificate ID ${esc(hit.id)}. Issued by Sheetah.</p>` : `<h1>${id ? 'Not found in the register' : 'Check a certificate'}</h1><p class="lead">${id ? `There’s no certificate with the ID <code>${esc(id)}</code> in this teacher’s register yet. New certificates appear once the teacher adds them.` : 'Type a certificate ID, such as SH-C1-2026-ABCDE.'}</p>`}
   <form id="vf" class="row"><input name="id" placeholder="Certificate ID" value="${esc(id)}" style="max-width:280px"><button class="btn">Check</button></form></div></section>`;
  $('#vf').onsubmit = e => { e.preventDefault(); go('#/verify/' + encodeURIComponent(new FormData(e.target).get('id').trim())); };
});
