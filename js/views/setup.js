// First launch: one question per screen, each with a live effect, then the plan, the tour or the first lesson.
// Also the optional placement check that suggests where to start and which modules to test out of.
import { register, go, applyTheme, getCourse, refreshChrome } from '../app.js';
import { store } from '../store.js';
import { session } from '../content.js';
import { esc, md, $, $$, mascot, line, toast, bar } from '../ui.js';
import { lessonList, isDone, nextStep, stepHref, hasKey, moduleInfo, who, freeModulesText, override } from '../access.js';
import { runTest, questionsFrom, shuffleOpts } from './tests.js';
import { shuffle, dateField, readDate } from '../ui.js';
import { buildPlan, setBaseline, iso } from '../plan.js';
import { today } from '../store.js';
const CNAMES = { c1: 'Excel Foundations', c2: 'Formulas & Functions' };
function givenCourses() { const c = Object.keys(CNAMES).filter(k => session.has(k)).map(k => `<b>${CNAMES[k]}</b>`); return c.length ? c.join(' and ') : '<b>Excel Foundations</b>'; }

const STEPS = ['welcome', 'look', 'name', 'why', 'level', 'tone', 'rhythm', 'computer', 'plan'];
export const WHY = [['job', 'A job or interviews'], ['college', 'College or exams'], ['business', 'My own business'], ['work', 'Getting faster at work'], ['curious', 'Just curious']];
export const LEVEL = [['new', 'I’ve never really used Excel', 'Start at the very first lesson.'], ['some', 'I use it a little', 'Start at the beginning, or take a 10-minute placement check.'], ['comfortable', 'I’m comfortable with formulas', 'A placement check can skip the modules you already know.']];
export const TONES = [['warm', 'Warm and encouraging'], ['cheeky', 'Cheeky'], ['coach', 'Strict coach'], ['desi', 'Hinglish, yaar']];
const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

// when the plan says everything (lessons, tests, project, exam) will be done
export function finishDate(c) { const P = buildPlan(c); return P.last && !P.error ? new Date(P.last + 'T00:00:00') : null; }
register('setup', async (main, step = 'welcome') => {
  let i = Math.max(0, STEPS.indexOf(step)); const s = store.d.settings, p = store.d.plan, pr = store.d.profile; const student = session.signedIn && !session.isAdmin;
  const c = await getCourse().catch(() => null);
  const to = k => go('#/setup/' + STEPS[k]);
  const done = (href = '#/home') => { store.d.setupDone = true; store.save(); refreshChrome(); go(href); };
  const nav = (canNext = true, nextLabel = 'Next') => `<div class="row between setup-nav"><div class="row">${i ? `<button class="btn ghost" id="sb" type="button">← Back</button>` : ''}<button class="linkbtn" id="sskip" type="button">Skip setup</button></div><div class="row">${i && i < STEPS.length - 1 ? '<button class="btn ghost" id="snext0" type="button">Skip this</button>' : ''}<button class="btn" id="sn" type="button" ${canNext ? '' : 'disabled'}>${nextLabel} →</button></div></div>`;
  const wrap = (title, body, opts = {}) => `<section class="card setup"><div class="setup-head"><p class="eyebrow">${i >= STEPS.length - 1 ? 'Setup complete' : `Setup · ${i} of ${STEPS.length - 2}`}</p>${i ? bar(i / (STEPS.length - 1), 'thin') : ''}</div><div class="setup-body">${opts.pose === false ? '' : mascot(opts.pose || 'hello', 'setup-m')}<div class="grow"><h1>${title}</h1>${body}</div></div>${opts.nav || nav()}</section>`;
  const wire = (onNext) => { const b = $('#sb'); if (b) b.onclick = () => to(i - 1); $('#sskip').onclick = () => done(); const z = $('#snext0'); if (z) z.onclick = () => to(i + 1); $('#sn').onclick = () => { if (onNext && onNext() === false) return; store.save(); to(i + 1); }; };
  const st = STEPS[i];
  if (st === 'welcome') {
    main.innerHTML = `<section class="card setup welcome">${mascot('hello', 'big')}<div class="grow"><p class="eyebrow">${student ? 'Signed in' : 'Welcome to Sheetah'}</p><h1>${student ? `Welcome, ${esc((session.s.name || '').split(' ')[0] || 'there')}!` : 'Learn Excel for a purpose, one lesson at a time'}</h1>
      <p class="lead">${student ? `Your teacher has given you ${givenCourses()}. ${override() ? 'Every module is open for you.' : 'Modules open one at a time as you pass each module test.'}` : 'Short lessons with a lecture, a demo, hands-on practice on a real working sheet, a quiz and a recap. Modules 1 and 2 and every study tool are free.'}</p>
      <p>A few quick questions set Sheetah up for you. It takes about 90 seconds, and you can change everything later in Settings.</p>
      <div class="row"><button class="btn big" id="sgo" type="button">Start setup</button><button class="btn ghost" id="squick" type="button">Skip and start lesson 1</button>${student ? '' : '<a class="btn ghost" href="#/login">I have a student login</a>'}</div></div></section>`;
    $('#sgo').onclick = () => to(1); $('#squick').onclick = () => { if (!store.d.name) store.d.name = ''; done(c ? stepHref(c, nextStep(c)) : '#/home'); };
    return;
  }
  if (st === 'look') {
    main.innerHTML = wrap('How should Sheetah look?', `<p>Tap a choice: the page changes straight away.</p><fieldset><legend>Theme</legend><div class="choices three">${[['light', 'Light', '☀'], ['dark', 'Dark', '☾'], ['auto', 'Match my computer', '◐']].map(([v, tt, ic]) => `<button type="button" class="choice ${s.theme === v ? 'on' : ''}" data-theme-v="${v}"><span class="cic">${ic}</span>${tt}</button>`).join('')}</div></fieldset>
      <fieldset><legend>Text size</legend><div class="choices four">${[['normal', 'Normal'], ['large', 'Large'], ['xl', 'Extra large'], ['xxl', 'Largest']].map(([v, tt], k) => `<button type="button" class="choice ${s.textSize === v ? 'on' : ''}" data-size="${v}"><span class="aa" style="font-size:${14 + k * 3}px">Aa</span>${tt}</button>`).join('')}</div></fieldset>`, { pose: 'teaching' });
    $$('[data-theme-v]').forEach(b => b.onclick = () => { s.theme = b.dataset.themeV; applyTheme(); store.save(); $$('[data-theme-v]').forEach(x => x.classList.toggle('on', x === b)); });
    $$('[data-size]').forEach(b => b.onclick = () => { s.textSize = b.dataset.size; applyTheme(); store.save(); $$('[data-size]').forEach(x => x.classList.toggle('on', x === b)); });
    wire(); return;
  }
  if (st === 'name') {
    main.innerHTML = wrap('What should Sheetah call you?', `<label class="big-label">Your first name <input id="nm" maxlength="40" autocomplete="given-name" value="${esc(store.d.name || (session.s && (session.s.name || '').split(' ')[0]) || '')}" placeholder="e.g. Asha"></label>`);
    const nm = $('#nm'); nm.focus(); nm.onkeydown = e => { if (e.key === 'Enter') $('#sn').click(); };
    wire(() => { store.d.name = nm.value.trim(); }); return;
  }
  if (st === 'why') {
    main.innerHTML = wrap('Why are you learning Excel?', `<p>This shapes your suggestions and the message on your home page. It never locks anything.</p><div class="choices">${WHY.map(([v, tt]) => `<button type="button" class="choice wide ${pr.why === v ? 'on' : ''}" data-v="${v}">${tt}</button>`).join('')}</div>`, { pose: 'thinking' });
    $$('[data-v]').forEach(b => b.onclick = () => { pr.why = b.dataset.v; store.save(); $$('[data-v]').forEach(x => x.classList.toggle('on', x === b)); });
    wire(); return;
  }
  if (st === 'level') {
    main.innerHTML = wrap('Where are you starting from?', `<div class="choices">${LEVEL.map(([v, tt, sub]) => `<button type="button" class="choice wide ${pr.level === v ? 'on' : ''}" data-v="${v}"><b>${tt}</b><small>${sub}</small></button>`).join('')}</div>`, { pose: 'focus' });
    $$('[data-v]').forEach(b => b.onclick = () => { pr.level = b.dataset.v; store.save(); $$('[data-v]').forEach(x => x.classList.toggle('on', x === b)); });
    wire(); return;
  }
  if (st === 'tone') {
    const nm = store.d.name || 'there'; const cur = s.tone;
    main.innerHTML = wrap('How should Sheetah talk to you?', `<p>Each one has a sample line, so you can hear the difference.</p><div class="choices two">${TONES.map(([v, tt]) => { s.tone = v; const sample = line('hi', nm) + ' ' + line('wrong'); return `<button type="button" class="choice tone ${cur === v ? 'on' : ''}" data-v="${v}"><b>${tt}</b><small>“${esc(sample)}”</small></button>`; }).join('')}</div>`, { pose: 'encourage' });
    s.tone = cur;
    $$('[data-v]').forEach(b => b.onclick = () => { s.tone = b.dataset.v; store.save(); $$('[data-v]').forEach(x => x.classList.toggle('on', x === b)); });
    wire(); return;
  }
  if (st === 'rhythm') {
    if (!p.start) p.start = today();
    const quick = [[1, '1 month'], [2, '2 months'], [3, '3 months'], [6, '6 months']];
    const draw = () => { const P = c ? buildPlan(c) : null; $('#rplan').innerHTML = !p.end ? '<span class="muted">Pick a finish date to see your daily time.</span>' : P && P.error ? `<span class="pstat bad">⚠ ${esc(P.error)}</span>` : P ? `<div class="d-big"><div class="bn"><b>${P.dates.length}</b><span>study days</span></div><div class="bn"><b>≈ ${P.perDay}</b><span>minutes a day</span></div></div>${P.perDay > 120 ? '<p class="small warnbox">That’s a lot each day: try a later date.</p>' : ''}` : ''; };
    main.innerHTML = wrap('When do you want to finish?', `<p>Pick a finish date and your study days. Sheetah builds the day-by-day plan for you, and you can change it any time.</p>
      <fieldset><legend>Finish Excel Foundations by</legend><div class="choices four">${quick.map(([n, l]) => `<button type="button" class="choice" data-months="${n}"><b>${l}</b></button>`).join('')}</div><p class="small">or pick a date: ${dateField('end', p.end || '', { label: 'Finish date' })}</p></fieldset>
      <fieldset><legend>Study days</legend><div class="days">${[1, 2, 3, 4, 5, 6, 0].map(k => `<label class="pill"><input type="checkbox" value="${k}" ${p.days.includes(k) ? 'checked' : ''}> ${DAYS[k]}</label>`).join('')}</div></fieldset>
      <div id="rplan" class="plan-card"></div>
      <div class="grid2"><label class="big-label">Usual study time <input type="time" id="tm" value="${p.time}"></label><label class="ck"><input type="checkbox" id="rm" ${s.remind ? 'checked' : ''}> Remind me at my study time</label></div><p class="muted small" id="rmHint"></p>`, { pose: 'focus' });
    const mark = () => { $$('[data-months]').forEach(b => { const d = new Date(p.start + 'T00:00:00'); d.setMonth(d.getMonth() + +b.dataset.months); b.classList.toggle('on', iso(d) === p.end); }); };
    const sync = () => { if (c) setBaseline(c); else store.save(); draw(); mark(); };
    $$('[data-months]').forEach(b => b.onclick = () => { const d = new Date(p.start + 'T00:00:00'); d.setMonth(d.getMonth() + +b.dataset.months); p.end = iso(d); const box = $('[data-date="end"]'); const [y, m, dd] = p.end.split('-').map(Number); box.querySelector('[data-part="d"]').value = dd; box.querySelector('[data-part="m"]').value = m; box.querySelector('[data-part="y"]').value = y; sync(); });
    $$('[data-date="end"] select').forEach(x => x.onchange = () => { const v = readDate(main, 'end'); if (v) { p.end = v; sync(); } });
    $$('.days input').forEach(x => x.onchange = () => { p.days = $$('.days input').filter(y => y.checked).map(y => +y.value).sort(); sync(); });
    $('#tm').onchange = e => { p.time = e.target.value; store.save(); };
    $('#rm').onchange = async e => { s.remind = e.target.checked; store.save(); if (s.remind && 'Notification' in window && Notification.permission === 'default') { const r = await Notification.requestPermission(); $('#rmHint').textContent = r === 'granted' ? 'Done: you’ll get a reminder at your study time while Sheetah is open.' : 'No problem: Sheetah will remind you inside the app instead.'; } };
    draw(); mark(); wire(); return;
  }
  if (st === 'computer') {
    main.innerHTML = wrap('Which computer do you use Excel on?', `<p>Every shortcut Sheetah shows will match your keyboard.</p><div class="choices two">${[['win', 'Windows', 'Ctrl+C to copy'], ['mac', 'Mac', '⌘+C to copy']].map(([v, tt, ex]) => `<button type="button" class="choice ${s.os === v ? 'on' : ''}" data-v="${v}"><b>${tt}</b><small>${ex}</small></button>`).join('')}</div>`, { pose: 'teaching' });
    $$('[data-v]').forEach(b => b.onclick = () => { s.os = b.dataset.v; store.save(); $$('[data-v]').forEach(x => x.classList.toggle('on', x === b)); });
    wire(); return;
  }
  if (st === 'plan') {
    const fin = c && finishDate(c); const days = p.days.map(k => DAYS[k]).join(', ');
    const first = c ? stepHref(c, nextStep(c)) : '#/home'; const placement = pr.level === 'some' || pr.level === 'comfortable';
    main.innerHTML = wrap(`You’re all set${store.d.name ? ', ' + esc(store.d.name) : ''}!`, `<div class="plan-card"><p class="lead">${fin ? `About <b>${p.minutesGoal} minutes</b> on <b>${days || 'no days yet'}</b>, and you finish Excel Foundations, project and exam included, by <b>${fin.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' })}</b>.` : 'Pick a finish date in your Study plan to see your day-by-day plan.'}</p><p class="muted">Your day-by-day plan is ready under <b>Study plan</b>. Change the dates there any time.</p></div>
      ${placement ? `<div class="tip">${mascot('thinking', 'xs')}<p>You said you already use Excel. A <b>10-minute placement check</b> suggests where to start and which modules you could skip.</p></div>` : ''}`,
      { pose: 'celebrate', nav: `<div class="row between setup-nav"><button class="btn ghost" id="sb" type="button">← Back</button><div class="row">${placement ? '<button class="btn ghost" id="spl" type="button">Placement check</button>' : ''}<button class="btn ghost" id="stour" type="button">Take the tour</button><button class="btn big" id="sfirst" type="button">Start lesson 1 →</button></div></div>` });
    $('#sb').onclick = () => to(i - 1); $('#stour').onclick = () => done('#/tour'); $('#sfirst').onclick = () => done(first); const pl = $('#spl'); if (pl) pl.onclick = () => done('#/placement/c1');
    return;
  }
});

// ---------- placement check (U5) ----------
register('placement', async (main, courseId = 'c1') => {
  const c = await getCourse(courseId); const reach = c.modules.filter(m => m.free || hasKey(c));
  const last = store.d.placement && store.d.placement[c.id];
  main.innerHTML = `<section class="card"><div class="row between"><div><p class="eyebrow">${esc(c.title)} · optional</p><h1>Placement check</h1></div>${mascot('focus')}</div>
    <p class="lead">About 10 minutes: two or three questions from each module. At the end, Sheetah suggests where to start and which modules you could test out of. It doesn’t unlock anything by itself.</p>
    ${reach.length < c.modules.length ? `<p class="warn">Without a student login, the check covers ${freeModulesText(c)} only.</p>` : ''}
    ${last ? `<p class="muted">Last time (${new Date(last.at).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' })}) Sheetah suggested starting at Module ${last.start}.</p>` : ''}
    <div class="row"><button class="btn big" id="pgo" type="button">Start the check</button><a class="btn ghost" href="#/course/${c.id}">Back to the map</a></div></section>`;
  $('#pgo').onclick = async () => {
    const qs = []; for (const m of reach) { const pool = shuffle(await questionsFrom(c, m.lessons)).slice(0, reach.length > 4 ? 2 : 4).map(q => ({ ...q, mod: m.n })); qs.push(...pool); }
    runTest(main, { title: 'Placement check', subtitle: c.title, questions: shuffle(qs).map(shuffleOpts), minutes: 12, pass: 0, kind: 'placement', record: false,
      after: (res) => {
        const per = {}; res.review.forEach(r => { if (!r.q) return; const k = r.q.mod; per[k] = per[k] || { right: 0, n: 0 }; per[k].n++; if (r.ok) per[k].right++; });
        const startM = c.modules.find(m => per[m.n] && per[m.n].right / per[m.n].n < 0.7) || c.modules.find(m => !per[m.n]) || c.modules[0];
        store.d.placement = { ...(store.d.placement || {}), [c.id]: { at: new Date().toISOString(), per, start: startM.n } }; store.save();
        const strong = c.modules.filter(m => per[m.n] && per[m.n].right === per[m.n].n && !m.free && moduleInfo(c, m).status !== 'done');
        main.innerHTML = `<section class="card">${mascot('celebrate')}<div><p class="eyebrow">Placement check</p><h1>Start at Module ${startM.n}: ${esc(startM.title)}</h1><p class="lead">Here’s how you did on each module.</p>
          <table class="tbl"><thead><tr><th>Module</th><th>Right</th><th></th></tr></thead><tbody>${c.modules.filter(m => per[m.n]).map(m => `<tr><td>Module ${m.n}: ${esc(m.title)}</td><td>${per[m.n].right} of ${per[m.n].n}</td><td>${per[m.n].right / per[m.n].n >= 0.7 ? '✓ looks solid' : 'worth doing'}</td></tr>`).join('')}</tbody></table>
          ${strong.length ? `<h2>You could test out of these</h2><p>Pass a module’s full test (70% or more) to skip it and open the next one.</p><div class="row">${strong.map(m => `<a class="btn ghost" href="#/mtest/${c.id}/${m.n}">Module ${m.n} test</a>`).join('')}</div>` : ''}
          <div class="row"><a class="btn" href="#/module/${c.id}/${startM.n}">Go to Module ${startM.n}</a><a class="btn ghost" href="#/course/${c.id}">See the map</a></div></div></section>`;
      } });
  };
});
