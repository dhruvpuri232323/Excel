// Module tests, practice (mock) tests, the course project and the final exam.
import { register, getCourse, go, lockedPanel } from '../app.js';
import { store, today } from '../store.js';
import { session, lesson as loadLesson, canOpen, getSecure, assetURL } from '../content.js';
import { esc, md, $, $$, mascot, toast, confetti, bar, shuffle, modal, lineChart, fmtDate, confirmBox } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask } from '../checker.js';
import { readXlsx, grade } from '../xlsxcheck.js';
import { lessonList, isDone } from './home.js';
import { resultsHTML, hydrate } from './lesson.js';

async function questionsFrom(c, lessons) { const out = []; for (const l of lessons) { try { const L = await loadLesson(c, l.id); (L.quiz || []).forEach((q, i) => out.push({ ...q, from: L.num, lesson: l.id })); } catch (e) { } } return out; }
async function labTasksFrom(c, lessons) { const out = []; for (const l of lessons) { try { const L = await loadLesson(c, l.id); for (const t of (L.lab && L.lab.tasks) || []) if (t.standalone) out.push({ sheet: L.lab.sheet, editable: L.lab.editable, inputs: L.lab.inputs, task: t, from: L.num, lesson: l.id, height: 300 }); } catch (e) { } } return out; }
const shuffleOpts = q => { if (!q.options) return q; const idx = shuffle(q.options.map((_, i) => i)); return { ...q, options: idx.map(i => q.options[i]), answer: idx.indexOf(q.answer) }; };

// A timed test: questions, then practical tasks on a live grid. Each task counts as two questions.
function runTest(main, { title, subtitle, questions, tasks = [], minutes, pass, onFinish, kind }) {
  const answers = {}, flags = new Set(); let i = 0; const n = questions.length + tasks.length; const ends = Date.now() + minutes * 60000; const grids = {}; let timer;
  const pages = [...questions.map((q, k) => ({ q, k })), ...tasks.map((t, k) => ({ t, k }))];
  const draw = () => {
    const p = pages[i]; Object.values(grids).forEach(g => g.host.closest('.tpage') && (g.host.parentElement.style.display = 'none'));
    main.innerHTML = `<div class="test"><div class="test-head card"><div><p class="eyebrow">${esc(subtitle || '')}</p><h1>${esc(title)}</h1></div><div class="timer" id="tm"></div></div>
     <div class="qnav">${pages.map((pp, k) => `<button class="qn ${k === i ? 'on' : ''} ${answered(pp) ? 'ans' : ''} ${flags.has(k) ? 'flag' : ''}" data-k="${k}" aria-label="Question ${k + 1}">${pp.t ? 'T' + (pp.k + 1) : k + 1}</button>`).join('')}</div>
     <div class="card tpage" id="tp"></div>
     <div class="row between"><button class="btn ghost" id="pv" ${i ? '' : 'disabled'}>← Previous</button><button class="btn ghost" id="fl">${flags.has(i) ? '⚑ Flagged' : '⚐ Flag for later'}</button>${i < n - 1 ? '<button class="btn" id="nx">Next →</button>' : '<button class="btn" id="sub">Submit test</button>'}</div></div>`;
    const tp = $('#tp');
    if (p.q) { const q = p.q; tp.innerHTML = `<p class="eyebrow">Question ${i + 1} of ${questions.length}</p><h2>${md(q.q)}</h2>${q.code ? `<p><code class="big">${esc(q.code)}</code></p>` : ''}${q.grid ? `<div class="mini-grid" data-grid='${esc(JSON.stringify(q.grid))}'></div>` : ''}${q.answers ? `<input class="ans big-input" value="${esc(answers[i] ?? '')}" placeholder="Type your answer" spellcheck="false">` : `<div class="opts">${q.options.map((o, k) => `<button class="opt ${answers[i] === k ? 'sel' : ''}" data-o="${k}">${md(o)}</button>`).join('')}</div>`}`;
      hydrate(tp, 'c1'); $$('.opt', tp).forEach(b => b.onclick = () => { answers[i] = +b.dataset.o; $$('.opt', tp).forEach(x => x.classList.toggle('sel', x === b)); markNav(); }); const inp = $('.ans', tp); if (inp) inp.oninput = () => { answers[i] = inp.value; markNav(); };
    } else { const t = p.t; tp.innerHTML = `<p class="eyebrow">Practical task ${p.k + 1} of ${tasks.length} · from lesson ${esc(t.from || '')}</p><h2>${md(t.task.prompt)}</h2><div class="tg"></div><p class="muted small">Your work here is saved as you go. It’s checked when you submit.</p>`;
      if (!grids[i]) { const host = document.createElement('div'); grids[i] = new Grid(host, new Workbook(t.sheet), { height: t.height || 300, editable: t.editable, highlight: t.editable, allowSort: true, onChange: () => { answers[i] = true; markNav(); } }); }
      $('.tg', tp).appendChild(grids[i].host); grids[i].render(); }
    $$('.qn').forEach(b => b.onclick = () => { i = +b.dataset.k; draw(); }); $('#pv').onclick = () => { i--; draw(); }; const nx = $('#nx'); if (nx) nx.onclick = () => { i++; draw(); };
    $('#fl').onclick = () => { flags.has(i) ? flags.delete(i) : flags.add(i); draw(); }; const sb = $('#sub'); if (sb) sb.onclick = async () => { const blank = pages.filter(pp => !answered(pp)).length; if (blank && !await confirmBox(`${blank} question${blank === 1 ? ' is' : 's are'} unanswered. Submit anyway?`, 'Submit', 'Keep going')) return; finish(); };
    tick();
  };
  const answered = pp => { const k = pages.indexOf(pp); return answers[k] !== undefined && answers[k] !== ''; };
  const markNav = () => $$('.qn').forEach((b, k) => b.classList.toggle('ans', answers[k] !== undefined && answers[k] !== ''));
  const tick = () => { clearInterval(timer); const el = $('#tm'); const up = () => { const left = Math.max(0, ends - Date.now()); const m = Math.floor(left / 60000), s = Math.floor(left / 1000) % 60; if (!document.body.contains(el)) { clearInterval(timer); return; } el.textContent = `${m}:${String(s).padStart(2, '0')} left`; el.classList.toggle('low', left < 5 * 60000); if (!left) { clearInterval(timer); toast('Time is up. Your test has been submitted.'); finish(); } }; up(); timer = setInterval(up, 1000); };
  const finish = () => {
    clearInterval(timer); let qRight = 0, tRight = 0; const review = [];
    questions.forEach((q, k) => { const a = answers[k]; const ok = q.answers ? (a && q.answers.some(x => x.replace(/\s/g, '').toUpperCase() === String(a).replace(/\s/g, '').toUpperCase())) : a === q.answer; if (ok) qRight++; review.push({ q, ok, given: q.answers ? a : (a !== undefined ? q.options[a] : '(blank)') }); if (!ok) store.logError({ course: 'c1', lesson: q.lesson, lessonNum: q.from, task: 'test', kind: kind, prompt: q.q, yours: review.at(-1).given || '(blank)', right: q.answers ? q.answers[0] : q.options[q.answer], why: q.explain }); });
    tasks.forEach((t, k) => { const g = grids[questions.length + k]; const res = g ? checkTask(t.task, { wb: g.wb, grid: g, inputs: t.inputs }) : { ok: false, msg: 'Not attempted.' }; if (res.ok) tRight++; review.push({ task: t.task, ok: res.ok, msg: res.msg }); });
    const total = questions.length + tasks.length * 2; const score = qRight + tRight * 2; const pct = Math.round(score / total * 100); const passed = pct >= pass;
    const result = { kind, at: new Date().toISOString(), pct, qRight, nq: questions.length, tRight, nt: tasks.length, passed, minutes: Math.round((minutes * 60000 - Math.max(0, ends - Date.now())) / 60000) };
    onFinish(result); store.save();
    main.innerHTML = `<div class="card test-end">${mascot(passed ? 'celebrate' : 'encourage')}<div><p class="eyebrow">${esc(title)}</p><h1>${pct}% · ${passed ? 'Passed' : 'Not passed yet'}</h1><p class="lead">${qRight} of ${questions.length} questions${tasks.length ? ` and ${tRight} of ${tasks.length} practical tasks` : ''} right. The pass mark is ${pass}%.</p>${passed ? '' : '<p>Every question you missed is in your error log, and comes back in your daily review.</p>'}<div class="row"><a class="btn" href="#/course/c1">Back to the course</a><a class="btn ghost" href="#/errors">Open the error log</a></div></div></div>
     <section class="card"><h2>Answers</h2><ol class="review-list">${review.map(r => r.q ? `<li class="${r.ok ? 'ok' : 'no'}"><p>${md(r.q.q)}</p><p class="small">${r.ok ? '✓' : '✗ You: ' + esc(r.given ?? '(blank)') + ' · Right: ' + esc(r.q.answers ? r.q.answers[0] : r.q.options[r.q.answer])}</p><p class="muted small">${md(r.q.explain || '')}</p></li>` : `<li class="${r.ok ? 'ok' : 'no'}"><p><b>Task:</b> ${md(r.task.prompt)}</p><p class="small">${r.ok ? '✓ Correct' : '✗ ' + esc(r.msg)}</p></li>`).join('')}</ol></section>`;
    if (passed) confetti(); scrollTo(0, 0);
  };
  draw();
}

register('mtest', async (main, courseId, n) => {
  const c = await getCourse(courseId); const m = c.modules.find(x => x.n === +n); if (!m) throw new Error('No such module');
  if (!m.free && !session.has(c.id)) { main.innerHTML = lockedPanel(); return; }
  const key = c.id + '-m' + m.n; const tries = store.d.tests.filter(t => t.kind === 'module' && t.module === key); const passed = tries.some(t => t.passed); const last = tries.at(-1);
  const wait = last && !last.passed ? Math.max(0, Date.parse(last.at) + 24 * 3600e3 - Date.now()) : 0;
  const done = m.lessons.filter(l => isDone(l.id)).length;
  main.innerHTML = `<section class="card">${mascot('focus')}<div><p class="eyebrow">${esc(c.code)} · Module ${m.n}</p><h1>Module ${m.n} test: ${esc(m.title)}</h1><p class="lead">${esc(m.testInfo)}. 25 minutes, pass mark 70%. Questions come from every lesson in this module, in a new order each time; practical tasks run on a live grid.</p>
   <ul class="ticks"><li>${done}/${m.lessons.length} lessons in this module done ${done < m.lessons.length ? '<span class="muted">(you can still take the test)</span>' : ''}</li><li>${tries.length ? `Attempts so far: ${tries.length} · best ${Math.max(...tries.map(t => t.pct))}%${passed ? ' · passed ✓' : ''}` : 'No attempts yet'}</li>${wait ? `<li>A retake opens in ${Math.ceil(wait / 3600e3)} hour(s). Use the time to review the error log.</li>` : ''}</ul>
   <div class="row"><button class="btn big" id="go" ${wait && !session.isAdmin ? 'disabled' : ''}>${passed ? 'Take it again for practice' : 'Start the test'}</button><a class="btn ghost" href="#/course/${c.id}">Back</a></div></div></section>`;
  $('#go').onclick = async () => {
    const qs = shuffle(await questionsFrom(c, m.lessons)).slice(0, 15).map(shuffleOpts); const ts = shuffle(await labTasksFrom(c, m.lessons)).slice(0, 2);
    runTest(main, { title: `Module ${m.n} test`, subtitle: `${c.code} · ${m.title}`, questions: qs, tasks: ts, minutes: 25, pass: 70, kind: 'module', onFinish: r => { store.d.tests.push({ ...r, module: key }); } });
  };
});

register('tests', async main => {
  const c = await getCourse('c1'); const open = lessonList(c).filter(l => canOpen(c, l)); const doneL = open.filter(l => isDone(l.id));
  const hist = store.d.tests.filter(t => t.kind === 'mock'); const target = store.d.mockTarget || 75;
  const weak = {}; store.d.errors.filter(e => e.kind === 'mock').forEach(e => weak[e.lessonNum] = (weak[e.lessonNum] || 0) + 1);
  main.innerHTML = `<h1>Practice tests</h1><p class="lead">Short timed tests that mix questions from every lesson you’ve done. Mixing topics like this is how real exams feel, and it beats revising one topic at a time.</p>
  <section class="card"><div class="row between"><div><h2>Start a test</h2><p class="muted">${doneL.length ? `Drawn from your ${doneL.length} finished lesson${doneL.length === 1 ? '' : 's'}` : `Drawn from the ${open.length} lessons open to you (finish lessons for a sharper test)`}</p></div>${mascot('focus', 'sm')}</div>
   <div class="row"><button class="btn" data-n="10">10 questions · 10 min</button><button class="btn" data-n="20">20 questions · 20 min</button><button class="btn ghost" data-n="30">30 questions · 30 min</button></div>
   <label class="inline">Your target score <input type="number" id="tg" min="40" max="100" value="${target}">%</label></section>
  <div class="grid2"><section class="card"><h2>Score trend</h2>${lineChart(hist.slice(-20).map(t => ({ k: fmtDate(t.at), v: t.pct })), { target, label: 'Practice test scores' })}</section>
  <section class="card"><h2>Weakest lessons in tests</h2>${Object.keys(weak).length ? `<ul class="weak">${Object.entries(weak).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([k, v]) => `<li><b>${esc(k)}</b> ${v} missed</li>`).join('')}</ul>` : '<p class="muted">After a test, the lessons you miss most show up here.</p>'}</section></div>
  <section class="card"><h2>Score sheet</h2>${hist.length ? `<table class="tbl"><thead><tr><th>Date</th><th>Questions</th><th>Score</th><th>Time</th></tr></thead><tbody>${hist.slice().reverse().map(t => `<tr><td>${fmtDate(t.at)}</td><td>${t.nq}</td><td><b>${t.pct}%</b> ${t.pct >= target ? '✓' : ''}</td><td>${t.minutes} min</td></tr>`).join('')}</tbody></table>` : '<p class="muted">No tests yet.</p>'}</section>`;
  $('#tg').onchange = e => { store.d.mockTarget = +e.target.value; store.save(); };
  $$('[data-n]').forEach(b => b.onclick = async () => { const n = +b.dataset.n; const pool = await questionsFrom(c, doneL.length >= 2 ? doneL : open); if (!pool.length) return toast('Open a lesson first.'); runTest(main, { title: `Practice test · ${n} questions`, subtitle: 'Mixed practice', questions: shuffle(pool).slice(0, n).map(shuffleOpts), minutes: n, pass: target, kind: 'mock', onFinish: r => store.d.tests.push(r) }); });
});

// ---------- project ----------
register('project', async (main, courseId) => {
  const c = await getCourse(courseId); if (!session.has(c.id)) { main.innerHTML = lockedPanel(); return; }
  const P = await getSecure(c.id, 'project'); const ps = store.d.projects[c.id] || (store.d.projects[c.id] = { tries: [] });
  const left = 2 - ps.tries.length;
  main.innerHTML = `<section class="card"><div class="row between"><div><p class="eyebrow">${esc(c.code)} course project · about 3 hours · pass mark 70%</p><h1>${esc(P.title)}</h1></div>${mascot('focus')}</div><p class="lead">${md(P.brief)}</p>
   <h2>What to build</h2><ol class="steps">${P.steps.map(s => `<li>${md(s)}</li>`).join('')}</ol>
   <h2>How it’s checked</h2><ul class="ticks">${P.rubricText.map(s => `<li>${md(s)}</li>`).join('')}</ul>
   <div class="row">${P.downloads.map(d => `<button class="btn ghost" data-file="${d.file}" data-name="${esc(d.name)}">⬇ ${esc(d.label)}</button>`).join('')}</div>
   <h2>Submit</h2><p>${left > 0 ? `You have <b>${left}</b> submission${left === 1 ? '' : 's'} left. Your best score counts.` : 'You have used both submissions.'}${ps.best !== undefined ? ` Best so far: <b>${ps.best}%</b>${ps.best >= 70 ? ' · passed ✓' : ''}.` : ''}</p>
   ${left > 0 || session.isAdmin ? `<label class="btn">⬆ Upload my workbook (.xlsx)<input type="file" accept=".xlsx,.xls,.csv" hidden id="up"></label>` : ''}<div class="results">${ps.last ? resultsHTML(ps.last) : ''}</div>
   ${ps.tries.length ? `<h3>Your submissions</h3><ul>${ps.tries.map(t => `<li>${fmtDate(t.at)}: <b>${t.pct}%</b></li>`).join('')}</ul>` : ''}</section>`;
  $$('[data-file]').forEach(b => b.onclick = async () => { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); });
  const up = $('#up'); if (up) up.onchange = async e => { const f = e.target.files[0]; if (!f) return; if (!await confirmBox(`Submit “${esc(f.name)}” as submission ${ps.tries.length + 1} of 2?`, 'Submit')) return; $('.results').innerHTML = '<p>Checking…</p>'; const x = await readXlsx(f); const res = grade(x, P.checks); ps.tries.push({ at: new Date().toISOString(), pct: res.pct }); ps.last = res; ps.best = Math.max(ps.best ?? 0, res.pct); store.save(); if (res.pct >= 70) confetti(); go('#/project/' + c.id); };
});

// ---------- final exam ----------
register('exam', async (main, courseId) => {
  const c = await getCourse(courseId); if (!session.has(c.id)) { main.innerHTML = lockedPanel(); return; }
  const E = c.exam; const att = store.d.examAttempts[c.id] || (store.d.examAttempts[c.id] = []); const passed = att.some(a => a.passed);
  const all = lessonList(c); const lessonsDone = all.every(l => isDone(l.id)); const modsDone = c.modules.every(m => store.d.tests.some(t => t.kind === 'module' && t.module === c.id + '-m' + m.n && t.passed)); const proj = store.d.projects[c.id]; const projOk = proj && proj.best >= 70;
  const last = att.at(-1); const wait = last && !last.passed ? Math.max(0, Date.parse(last.at) + E.gapHours * 3600e3 - Date.now()) : 0;
  const ready = lessonsDone && modsDone && projOk && !passed && att.length < E.attempts && !wait;
  main.innerHTML = `<section class="card"><div class="row between"><div><p class="eyebrow">${esc(c.code)} final exam</p><h1>${esc(c.title)}: final exam</h1></div>${mascot('focus')}</div>
   <p class="lead">${E.minutes} minutes · ${E.questions} questions and ${E.tasks} practical tasks · pass mark ${E.pass}% · up to ${E.attempts} attempts, ${E.gapHours} hours apart, each with new questions and data.</p>
   <ul class="reqs"><li class="${lessonsDone ? 'ok' : ''}">${lessonsDone ? '✓' : '○'} Every lesson done (${all.filter(l => isDone(l.id)).length}/${all.length})</li><li class="${modsDone ? 'ok' : ''}">${modsDone ? '✓' : '○'} Every module test passed</li><li class="${projOk ? 'ok' : ''}">${projOk ? '✓' : '○'} Course project passed${proj && proj.best !== undefined ? ` (best ${proj.best}%)` : ''}</li></ul>
   ${att.length ? `<p>Attempts: ${att.map(a => `${fmtDate(a.at)} <b>${a.pct}%</b>`).join(' · ')}</p>` : ''}${wait ? `<p>Your next attempt opens in ${Math.ceil(wait / 3600e3)} hours.</p>` : ''}${passed ? '<p><b>You passed.</b> Your certificate is on the Certificates page.</p>' : ''}
   <p class="muted small">Before you start: use a laptop, close other tabs, and keep 90 minutes free. Sheetah’s exam follows an honour code: work alone, no help from others.</p>
   <div class="row"><button class="btn big" id="go" ${ready ? '' : 'disabled'}>Start the exam</button>${session.isAdmin ? '<button class="btn ghost" id="prev">Admin preview (doesn’t count)</button>' : ''}<a class="btn ghost" href="#/course/${c.id}">Back</a></div></section>`;
  const start = async preview => {
    if (!preview && !await confirmBox(`Start the exam now? The ${E.minutes}-minute timer can’t be paused.`, 'Start')) return;
    const bank = await getSecure(c.id, 'exam'); const qs = shuffle(bank.questions).slice(0, E.questions).map(shuffleOpts); const ts = shuffle(bank.tasks).slice(0, E.tasks);
    runTest(main, { title: `${c.code} final exam`, subtitle: preview ? 'Admin preview · not recorded' : `Attempt ${att.length + 1} of ${E.attempts}`, questions: qs, tasks: ts, minutes: E.minutes, pass: E.pass, kind: 'exam', onFinish: r => { if (preview) return; att.push(r); if (r.passed) award(c, r.pct); } });
  };
  $('#go').onclick = () => start(false); const pv = $('#prev'); if (pv) pv.onclick = () => start(true);
});
function award(c, examPct) {
  const proj = store.d.projects[c.id].best; const gradePct = Math.round(examPct * .6 + proj * .4); const id = 'SH-' + c.code + '-' + new Date().getFullYear() + '-' + Math.random().toString(36).slice(2, 7).toUpperCase();
  store.d.certs.push({ course: c.id, code: c.code, title: c.title, id, at: new Date().toISOString(), grade: gradePct, distinction: gradePct >= 90, name: (session.s && session.s.name) || store.d.name, skills: c.outcomes.slice(0, 4) });
}
register('certificates', async main => {
  const certs = store.d.certs;
  main.innerHTML = `<h1>Certificates</h1><p class="lead">Each course earns its own certificate once the project and final exam are passed. Finish all eight courses and the Master Certificate opens.</p>
  ${certs.length ? certs.map(ct => `<section class="card certrow"><div><h2>${esc(ct.code)} · ${esc(ct.title)}</h2><p>${fmtDate(ct.at)} · grade ${ct.grade}%${ct.distinction ? ' · with Distinction' : ''} · ID ${esc(ct.id)}</p></div><a class="btn" href="#/cert/${ct.id}">View and print</a></section>`).join('') : `<section class="card empty">${mascot('thinking')}<div><h2>No certificates yet</h2><p>C1 · Excel Foundations: finish the lessons and module tests, pass the project, then the final exam.</p><a class="btn" href="#/course/c1">Go to the course</a></div></section>`}`;
});
register('cert', async (main, id) => {
  const ct = store.d.certs.find(x => x.id === id); if (!ct) throw new Error('Certificate not found on this laptop');
  main.innerHTML = `<div class="row noprint"><button class="btn" onclick="print()">Print or save as PDF</button><a class="btn ghost" href="#/certificates">Back</a></div>
  <div class="certificate"><div class="cert-inner"><img src="img/mascot/sheetah-celebrate.svg" alt="" class="cert-m"><p class="cert-eyebrow">Certificate of completion</p><h1>${esc(ct.name)}</h1><p>has completed</p><h2>${esc(ct.code)} · ${esc(ct.title)}</h2><p class="grade">Grade ${ct.grade}%${ct.distinction ? ' · with Distinction' : ''}</p>
  <ul class="cert-skills">${(ct.skills || []).map(s => `<li>${esc(s)}</li>`).join('')}</ul><div class="cert-foot"><span>Issued ${fmtDate(ct.at)}</span><span>Certificate ID ${esc(ct.id)}</span></div>
  <p class="cert-legal">Issued by Sheetah. Not affiliated with or endorsed by Microsoft. Microsoft and Excel are trademarks of the Microsoft group of companies.</p></div></div>`;
});
