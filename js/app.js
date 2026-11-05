// Sheetah app shell: routing, navigation, sign-in, and the shared helpers views use.
import { store, today } from './store.js';
import { session, catalog, course as loadCourse } from './content.js';
import { esc, $, $$, toast, mascot, line, modal } from './ui.js';
import { login, loadRoster } from './auth.js';
import { dayNum } from './srs.js';

const views = {};
export const register = (name, fn) => { views[name] = fn; };
export const app = { course: null, catalog: null, version: '1.0.0' };
export const go = h => { if (location.hash === h) route(); else location.hash = h; };

const NAV = [
  ['home', 'Home', '⌂'], ['courses', 'Courses', '▦'], ['review', 'Daily review', '↻'], ['tests', 'Practice tests', '✓'], ['errors', 'Error log', '!'],
  ['planner', 'Study plan', '▤'], ['journal', 'Journal', '✎'], ['insights', 'Insights', '▲'], ['library', 'Function library', 'ƒ'], ['shortcuts', 'Shortcuts', '⌘'], ['playground', 'Playground', '⊞'], ['certificates', 'Certificates', '★'], ['settings', 'Settings', '⚙']
];
function applyTheme() { const t = store.d.settings.theme; const dark = t === 'dark' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches); document.documentElement.dataset.theme = dark ? 'dark' : 'light'; document.documentElement.dataset.text = store.d.settings.textSize; document.documentElement.classList.toggle('reduce-motion', !!store.d.settings.reduceMotion); }
function shell() {
  document.body.innerHTML = `<a class="skip" href="#main">Skip to content</a>
  <aside class="side" aria-label="Main">
    <a class="brand" href="#/home">${mascot('avatar', 'brand-m')}<span>Sheetah<small>Excel, for a purpose</small></span></a>
    <nav>${NAV.map(([k, t, i]) => `<a href="#/${k}" data-nav="${k}"><span class="ni" aria-hidden="true">${i}</span>${t}</a>`).join('')}<a href="#/admin" data-nav="admin" class="admin-link" hidden><span class="ni">◆</span>Admin</a></nav>
    <div class="side-foot" id="sideFoot"></div>
  </aside>
  <div class="mainwrap"><header class="top"><button class="btn ghost menu-btn" aria-label="Menu">☰</button><div class="top-stats" id="topStats"></div><div class="top-user" id="topUser"></div></header><main id="main" tabindex="-1"></main></div>`;
  $('.menu-btn').onclick = () => document.body.classList.toggle('nav-open');
  $$('.side nav a').forEach(a => a.addEventListener('click', () => document.body.classList.remove('nav-open')));
}
export function refreshChrome() {
  const st = store.streak(); const t = store.d.activity[today()] || { minutes: 0 }; const goal = store.d.plan.minutesGoal || 30;
  const due = Object.values(store.d.cards).filter(c => c.due <= dayNum()).length + store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length;
  $('#topStats').innerHTML = `<a class="chip" href="#/insights" title="Study streak">🔥 <b>${st}</b> day${st === 1 ? '' : 's'}</a><a class="chip" href="#/planner" title="Today’s minutes against your goal">⏱ <b>${Math.round(t.minutes)}</b>/${goal} min</a><a class="chip ${due ? 'due' : ''}" href="#/review" title="Cards due">↻ <b>${due}</b> due</a>`;
  $('#topUser').innerHTML = session.signedIn ? `<span class="who">${esc(session.s.name || session.s.email)}${session.isAdmin ? ' · admin' : ''}</span><button class="btn ghost sm" id="signOut">Sign out</button>` : `<span class="who">${esc(store.d.name || 'Guest')} · free access</span><a class="btn sm" href="#/login">Student sign-in</a>`;
  const so = $('#signOut'); if (so) so.onclick = () => { session.clear(); store.load('guest'); applyTheme(); toast('Signed out. Free lessons are still open.'); go('#/home'); refreshChrome(); };
  $('.admin-link').hidden = !session.isAdmin;
  $('#sideFoot').innerHTML = `<small>${session.signedIn ? 'Signed in: full course unlocked' : 'Free access: Modules 1–2, tools and review'}</small>`;
}
async function route() {
  const hash = location.hash || '#/home'; const [, name = 'home', ...rest] = hash.split('/');
  $$('.side nav a').forEach(a => a.classList.toggle('on', a.dataset.nav === name || (name === 'course' || name === 'lesson' || name === 'exam' || name === 'project' || name === 'mtest') && a.dataset.nav === 'courses'));
  store.studying = name === 'lesson' || name === 'review' || name === 'exam' || name === 'mtest' || name === 'playground' || name === 'project'; store.currentLesson = name === 'lesson' ? rest[1] : null;
  const main = $('#main'); main.innerHTML = '<div class="loading">Loading…</div>';
  if (window.speechSynthesis) speechSynthesis.cancel();
  if (!store.d.name && !session.signedIn && name !== 'login' && name !== 'welcome') return views.welcome(main);
  const v = views[name] || views.home;
  try { await v(main, ...rest.map(decodeURIComponent)); } catch (e) { console.error(e); main.innerHTML = e.message === 'locked' ? lockedPanel() : `<div class="card"><h2>Something went wrong</h2><p>${esc(e.message)}</p><a class="btn" href="#/home">Go home</a></div>`; }
  refreshChrome(); main.focus({ preventScroll: true }); scrollTo(0, 0);
}
export const lockedPanel = () => `<div class="card locked-card">${mascot('thinking')}<div><h2>This part is for enrolled students</h2><p>Your teacher gives each student an email and password. Sign in to open Modules 3–8, the course project, the final exam and the certificate. Modules 1–2 and all the study tools stay free for everyone.</p><div class="row"><a class="btn" href="#/login">Student sign-in</a><a class="btn ghost" href="#/course/c1">See the free lessons</a></div></div></div>`;

// ----- welcome / onboarding (free for anyone) -----
register('welcome', async main => {
  main.innerHTML = `<section class="welcome card">${mascot('hello', 'big')}<div><p class="eyebrow">Welcome to Sheetah</p><h1>Learn Excel for a purpose, one lesson at a time</h1>
  <p class="lead">Real lessons with lectures, hands-on labs, quizzes and spaced review. Modules 1 and 2 of Excel Foundations, the function library, the shortcut trainer and every study tool are free. Students with a login get the full course and the certificate.</p>
  <form id="wf" class="form">
   <label>What should Sheetah call you? <input name="name" required maxlength="40" autocomplete="given-name" placeholder="Your first name"></label>
   <fieldset><legend>Which days will you study?</legend><div class="days">${['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].map((d, i) => `<label class="pill"><input type="checkbox" name="d" value="${i}" ${[1, 3, 5].includes(i) ? 'checked' : ''}> ${d}</label>`).join('')}</div></fieldset>
   <div class="grid2"><label>At what time? <input type="time" name="time" value="20:00"></label><label>Daily goal <select name="goal"><option value="15">15 minutes</option><option value="30" selected>30 minutes</option><option value="45">45 minutes</option><option value="60">60 minutes</option></select></label></div>
   <label>How should Sheetah talk to you? <select name="tone"><option value="warm">Warm and encouraging</option><option value="cheeky">Cheeky</option><option value="coach">Strict coach</option><option value="desi">Hinglish, yaar</option></select></label>
   <div class="row"><button class="btn big">Start learning</button><a class="btn ghost" href="#/login">I have a student login</a></div>
  </form></div></section>`;
  $('#wf').onsubmit = e => { e.preventDefault(); const f = new FormData(e.target); store.d.name = f.get('name').trim(); store.d.plan.days = f.getAll('d').map(Number); store.d.plan.time = f.get('time'); store.d.plan.minutesGoal = +f.get('goal'); store.d.settings.tone = f.get('tone'); store.save(); toast(line('hi', store.d.name)); go('#/home'); };
});
// ----- student sign-in -----
register('login', async main => {
  main.innerHTML = `<section class="card login">${mascot('hello')}<div><h1>Student sign-in</h1><p>Use the email and password your teacher gave you. There is nothing to buy: access is given by your teacher.</p>
  <form id="lf" class="form"><label>Email <input type="email" name="email" required autocomplete="username"></label><label>Password <input type="password" name="pw" required autocomplete="current-password"></label>
  <label class="ck"><input type="checkbox" name="rem" checked> Keep me signed in on this laptop</label><p class="err" id="lerr" role="alert"></p><div class="row"><button class="btn">Sign in</button><a class="btn ghost" href="#/home">Continue with free access</a></div></form>
  <p class="muted small">Your progress is saved on this laptop. Use Settings › Back up progress to move it to another laptop.</p></div></section>`;
  $('#lf').onsubmit = async e => {
    e.preventDefault(); const f = new FormData(e.target); const btn = e.target.querySelector('button'); btn.disabled = true; btn.textContent = 'Checking…'; $('#lerr').textContent = '';
    try { const p = await login(f.get('email'), f.get('pw')); const guest = store.d; session.set({ id: p.id, email: p.email, name: p.name, role: p.role, keys: p.keys, adminKey: p.adminKey || null, expires: p.expires || null }, !!f.get('rem'));
      store.load(session.profileId()); if (!store.d.name) { store.d.name = (p.name || '').split(' ')[0] || guest.name; // carry guest progress over on first sign-in
        if (Object.keys(guest.lessons).length && !Object.keys(store.d.lessons).length) { const keep = store.d.name; Object.assign(store.d, JSON.parse(JSON.stringify(guest)), { name: keep }); } store.save(); }
      applyTheme(); toast(`Welcome, ${esc(p.name || p.email)}. The full course is unlocked.`); go(p.role === 'admin' ? '#/admin' : '#/course/c1');
    } catch (err) { $('#lerr').textContent = err.message; btn.disabled = false; btn.textContent = 'Sign in'; }
  };
});

export async function getCourse(id = 'c1') { return loadCourse(id); }
export async function boot() {
  session.load(); store.load(session.profileId()); applyTheme(); matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
  shell(); app.catalog = await catalog().catch(() => ({ courses: [] }));
  store.on(() => refreshChrome()); addEventListener('hashchange', route); await route();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
  reminders();
}
export { applyTheme };
// gentle reminder at the learner's chosen time while the app is open (one a day at most)
function reminders() {
  setInterval(() => {
    const s = store.d.settings, p = store.d.plan; if (!s.remind) return; const now = new Date(); const hm = now.toTimeString().slice(0, 5); const k = 'sheetah:reminded:' + today();
    if (!p.days.includes(now.getDay()) || hm < p.time || localStorage.getItem(k) || store.studied(today())) return; localStorage.setItem(k, '1');
    const msg = line('idle'); if ('Notification' in window && Notification.permission === 'granted') new Notification('Sheetah', { body: msg, icon: 'img/icon-192.png' }); else toast(mascot('sleepy', 'xs') + ' ' + msg);
  }, 60000);
}
