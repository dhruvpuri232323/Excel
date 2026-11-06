// Sheetah app shell: side panel, top bar, phone bar, routing, search, keyboard keys, focus mode and sign-in.
import { store, today } from './store.js';
import { session, catalog, course as loadCourse, getJSON } from './content.js';
import { esc, $, $$, toast, mascot, modal, popover, closePopover, setKeyMap, md, confirmBox } from './ui.js';
import { login, aesKey, decryptBytes, b64 } from './auth.js';
import { dayNum } from './srs.js';
import { site, loadSite } from './site.js';
import { t } from './i18n.js';
import { HELP_TIPS, ARTICLES, askTeacher, setGlossary, showGloss } from './help.js';
import { startTour, clearTip } from './tour.js';
import { openSearch } from './search.js';

const views = {};
export const register = (name, fn) => { views[name] = fn; };
export const app = { course: null, catalog: null, version: '2.0.0' };
export const go = h => { if (location.hash === h) route(); else location.hash = h; };
export const keys = { handler: null }; // the current page's own keys (lessons use ← → N F)

const NAV = [
  ['learn', 'Learn', [['home', 'Home', '⌂'], ['courses', 'My courses', '▦'], ['review', 'Daily review', '↻'], ['tests', 'Tests', '✓']]],
  ['practise', 'Practise', [['playground', 'Playground', '⊞'], ['challenge', 'Weekly challenge', '⚑'], ['interview', 'Interview prep', '◎'], ['library', 'Function library', 'ƒ'], ['shortcuts', 'Shortcuts', '⌘']]],
  ['track', 'Track', [['planner', 'Study plan', '▤'], ['journal', 'Journal', '✎'], ['errors', 'Mistakes', '!'], ['insights', 'Insights', '▲'], ['certificates', 'Certificates', '★']]],
  ['mine', 'My stuff', [['notes', 'My notes', '✐'], ['saved', 'Saved', '☆'], ['downloads', 'Downloads', '⬇'], ['topics', 'Topic index', '≡']]],
];
const NAV_OF = { course: 'courses', module: 'courses', lesson: 'courses', mtest: 'courses', project: 'courses', exam: 'courses', placement: 'courses', practice: 'courses', cert: 'certificates', verify: 'certificates', how: 'help' };
const groupOf = k => (NAV.find(([, , items]) => items.some(i => i[0] === k)) || [])[0];
const STUDY = ['lesson', 'review', 'exam', 'mtest', 'playground', 'project', 'practice', 'challenge', 'placement'];

export function applyTheme() {
  const s = store.d.settings, h = document.documentElement; const dark = s.theme === 'dark' || (s.theme === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
  Object.assign(h.dataset, { theme: dark ? 'dark' : 'light', text: s.textSize, accent: s.accent, spacing: s.lineSpacing, font: s.font, density: s.density, contrast: s.contrast ? 'high' : 'normal' });
  h.classList.toggle('reduce-motion', !!s.reduceMotion); h.lang = s.lang === 'hi' ? 'hi' : 'en';
  document.body && document.body.classList.toggle('side-folded', !!s.sideFolded);
  if (s.font === 'easy' && !document.getElementById('easyFont')) { const l = document.createElement('link'); l.id = 'easyFont'; l.rel = 'stylesheet'; l.href = 'https://fonts.googleapis.com/css2?family=Atkinson+Hyperlegible:wght@400;700&display=swap'; document.head.appendChild(l); }
}
export function setFocus(on, full = false) {
  document.body.classList.toggle('focus', on); $('#focusExit').hidden = !on;
  if (on && full && document.documentElement.requestFullscreen && !document.fullscreenElement) document.documentElement.requestFullscreen().catch(() => {});
  if (!on && document.fullscreenElement) document.exitFullscreen().catch(() => {});
}
const navLink = ([k, label, icon]) => `<a href="#/${k}" data-nav="${k}" title="${esc(t(label))}"><span class="ni" aria-hidden="true">${icon}</span><span class="nl">${esc(t(label))}</span></a>`;
function shell() {
  document.body.innerHTML = `<a class="skip" href="#main">Skip to content</a>
  <aside class="side" aria-label="Main menu">
    <div class="side-top"><a class="brand" href="#/home">${mascot('avatar', 'brand-m')}<span class="nl">Sheetah<small>Excel, for a purpose</small></span></a><button class="foldbtn" id="foldBtn" type="button"></button></div>
    <button class="side-search" id="sideSearch" type="button"><span class="ni" aria-hidden="true">⌕</span><span class="nl">${esc(t('Search'))}</span><kbd class="nl">Ctrl K</kbd></button>
    <nav aria-label="Sections">${NAV.map(([g, title, items]) => `<div class="ngroup" data-g="${g}"><button class="ghead" type="button" aria-expanded="true"><span class="nl">${esc(t(title))}</span><span class="gch" aria-hidden="true">▾</span></button><div class="gitems">${items.map(navLink).join('')}</div></div>`).join('')}</nav>
    <div class="side-bottom">${navLink(['help', 'Help', '?'])}${navLink(['settings', 'Settings', '⚙'])}<a href="#/admin" data-nav="admin" class="admin-link" hidden title="Admin"><span class="ni">◆</span><span class="nl">${esc(t('Admin'))}</span></a></div>
    <div class="side-foot nl" id="sideFoot"></div>
  </aside>
  <div class="mainwrap"><header class="top"><button class="btn ghost menu-btn" type="button" aria-label="Menu">☰</button><div class="top-stats" id="topStats"></div>
    <div class="top-right"><button class="btn ghost sm top-search" id="searchBtn" type="button" aria-label="${esc(t('Search'))}"><span aria-hidden="true">⌕</span><span class="tl">${esc(t('Search'))}</span><kbd class="tl">Ctrl K</kbd></button><button class="btn ghost sm" id="helpBtn" type="button" aria-label="${esc(t('Help'))}">?<span class="tl">${esc(t('Help'))}</span></button><div class="top-user" id="topUser"></div></div></header>
    <div id="banner"></div><main id="main" tabindex="-1"></main></div>
  <nav class="bottombar" aria-label="Quick menu"><a href="#/home" data-tab="home"><span>⌂</span>${esc(t('Home'))}</a><a href="#/course/c1" data-tab="map"><span>▦</span>${esc(t('Map'))}</a><a href="#/review" data-tab="review"><span>↻</span>${esc(t('Review'))}</a><button data-tab="search" type="button"><span>⌕</span>${esc(t('Search'))}</button><button data-tab="more" type="button"><span>☰</span>${esc(t('More'))}</button></nav>
  <button class="focus-exit btn sm" id="focusExit" type="button" hidden>${esc(t('Exit focus'))} · Esc</button>`;
  const openDrawer = () => document.body.classList.toggle('nav-open');
  $('.menu-btn').onclick = openDrawer; $('[data-tab="more"]').onclick = openDrawer; $('[data-tab="search"]').onclick = () => openSearch();
  $$('.side a').forEach(a => a.addEventListener('click', () => document.body.classList.remove('nav-open')));
  $('#foldBtn').onclick = toggleFold; $('#sideSearch').onclick = () => openSearch(); $('#searchBtn').onclick = () => openSearch(); $('#helpBtn').onclick = () => go('#/help');
  $('#focusExit').onclick = () => setFocus(false);
  $$('.ghead').forEach(b => b.onclick = () => { const g = b.parentElement.dataset.g; store.d.navOpen[g] = !isOpen(g); store.save(); drawGroups(); });
  document.addEventListener('click', e => { if (e.target.closest('.side') || e.target.closest('.menu-btn') || e.target.closest('[data-tab="more"]')) return; document.body.classList.remove('nav-open'); });
  drawGroups(); drawFold();
}
const isOpen = g => store.d.navOpen[g] !== undefined ? !!store.d.navOpen[g] : g === 'learn';
function drawGroups(active) { $$('.ngroup').forEach(n => { const g = n.dataset.g; const open = isOpen(g) || g === groupOf(active); n.classList.toggle('closed', !open); $('.ghead', n).setAttribute('aria-expanded', String(open)); }); }
function drawFold() { const f = !!store.d.settings.sideFolded; document.body.classList.toggle('side-folded', f); const b = $('#foldBtn'); b.textContent = f ? '»' : '«'; b.title = (f ? t('Open the side panel') : t('Fold the side panel')) + ' ( [ )'; b.setAttribute('aria-label', b.title); }
export function toggleFold() { store.d.settings.sideFolded = !store.d.settings.sideFolded; store.save(); drawFold(); }

export function refreshChrome() {
  const st = store.streak(); const a = store.d.activity[today()] || { minutes: 0 }; const goal = store.d.plan.minutesGoal || 30;
  const due = Object.values(store.d.cards).filter(c => c.due <= dayNum() && !c.isNew).length + store.d.errors.filter(e => !e.resolved && e.due[e.step] <= dayNum()).length;
  $('#topStats').innerHTML = `<a class="chip" href="#/insights" title="${esc(HELP_TIPS.streak.replace(/\*\*/g, ''))}">🔥<span><b>${st}</b><span class="cl">${store.d.settings.lang === 'hi' ? ' ' + t('day streak') : '-day streak'}</span></span></a><a class="chip" href="#/planner" title="${esc(HELP_TIPS.minutes.replace(/\*\*/g, ''))}">⏱<span><b>${Math.round(a.minutes)}</b><span class="cl"> of ${goal} ${t('min today')}</span></span></a><a class="chip ${due ? 'due' : ''}" href="#/review" title="${esc(HELP_TIPS.due)}">↻<span><b>${due}</b><span class="cl"> ${store.d.settings.lang === 'hi' ? '' : due === 1 ? 'card ' : 'cards '}${t('to review')}</span></span></a>`;
  const nm = store.preview ? `Preview: ${{ visitor: 'visitor', new: 'new student', half: 'student halfway' }[store.preview.mode]}` : session.signedIn ? (store.d.name || (session.s.name || '').split(' ')[0] || session.s.email) : (store.d.name || 'Guest');
  $('#topUser').innerHTML = `<a class="who" href="#/settings/profile" title="Your profile">${mascot(store.d.profile.avatar || 'avatar', 'av keep')}<span class="tl">${esc(nm)}</span>${session.isAdmin && !store.preview ? '<span class="tag tl">admin</span>' : ''}</a>${session.signedIn ? `<button class="btn ghost sm" id="signOut" type="button">${esc(t('Sign out'))}</button>` : `<a class="btn sm" href="#/login">${store.d.settings.lang === 'hi' ? esc(t('Student sign-in')) : '<span class="tl">Student sign-in</span><span class="ph">Sign in</span>'}</a>`}`;
  const so = $('#signOut'); if (so) so.onclick = async () => { if (store.preview) store.endPreview(); session.clear(); store.load('guest'); applyTheme(); toast('Signed out. The free lessons are still open.'); go('#/home'); };
  $('.admin-link').hidden = !session.isAdmin;
  $('#sideFoot').innerHTML = `<small>${store.preview ? 'Preview mode: nothing is saved' : session.isAdmin ? 'Signed in as admin' : session.signedIn ? `${t('Signed in')}: ${esc(session.s.email)}` : `${t('Free access')}: the free modules of every course, and every tool`}</small>`;
  drawBanner();
}
function drawBanner() {
  const b = $('#banner'); const ann = site.data.announcement; let html = '';
  if (store.preview) html += `<div class="banner warnbox"><b>Previewing as a ${{ visitor: 'visitor', new: 'new student', half: 'student halfway through' }[store.preview.mode]}.</b> Nothing you do here is saved. <button class="btn sm" id="endPrev" type="button">Exit preview</button></div>`;
  if (ann && ann.text && store.d.annSeen !== ann.at) html += `<div class="banner ann"><span>📣</span><p>${md(ann.text)}</p><button class="linkbtn" id="annX" type="button">Dismiss</button></div>`;
  b.innerHTML = html;
  const ep = $('#endPrev'); if (ep) ep.onclick = () => { store.endPreview(); applyTheme(); go('#/admin'); };
  const ax = $('#annX'); if (ax) ax.onclick = () => { store.d.annSeen = ann.at; store.save(); drawBanner(); };
}

async function route() {
  closePopover(); clearTip(); keys.handler = null;
  const hash = location.hash || '#/home'; let [, name = 'home', ...rest] = hash.split('/'); name = name || 'home';
  if (!store.d.setupDone && !store.preview && !session.isAdmin && !['setup', 'login', 'help', 'how', 'verify'].includes(name)) { location.replace('#/setup'); return; }
  const navKey = NAV_OF[name] || name;
  $$('.side a').forEach(a => a.classList.toggle('on', a.dataset.nav === navKey)); $$('.bottombar [data-tab]').forEach(a => a.classList.toggle('on', a.dataset.tab === ({ home: 'home', courses: 'map', review: 'review' }[navKey] || '')));
  drawGroups(navKey);
  if (!['lesson', 'mtest', 'exam', 'placement'].includes(name) && document.body.classList.contains('focus')) setFocus(false);
  store.studying = STUDY.includes(name); store.currentLesson = name === 'lesson' ? rest[1] : null;
  const main = $('#main'); main.innerHTML = '<div class="loading">Loading…</div>';
  if (name === 'tour') { await views.home(main); refreshChrome(); history.replaceState(null, '', '#/home'); startTour(() => { const tc = document.querySelector('.tourcall'); if (tc) tc.remove(); }); return; }
  const v = views[name] || views.notfound;
  try { await v(main, ...rest.map(decodeURIComponent)); } catch (e) { console.error(e); main.innerHTML = e.message === 'locked' ? lockedPanel() : `<div class="card"><h2>Something went wrong</h2><p>${esc(e.message)}</p><div class="row"><a class="btn" href="#/home">${t('Go home')}</a><button class="btn ghost" id="errAsk" type="button">${t('Ask your teacher')}</button></div></div>`; const ea = $('#errAsk'); if (ea) ea.onclick = () => askTeacher({ where: location.hash, what: 'I saw this message: ' + e.message }); }
  refreshChrome(); if (!main.contains(document.activeElement)) main.focus({ preventScroll: true }); scrollTo(0, 0);
}
export const lockedPanel = (why = 'login', extra = '') => `<div class="card locked-card">${mascot('thinking')}<div>${why === 'login' ? `<h2>This part is for enrolled students</h2><p>Your teacher gives each student an email and password. With a login, the rest of each course you’re enrolled in opens one module at a time as you pass each module test, then the course project, the final exam and the certificate. The free modules and every study tool stay open to everyone.</p><div class="row"><a class="btn" href="#/login">Student sign-in</a><button class="btn ghost" type="button" data-ask="login">Ask your teacher for a login</button><a class="btn ghost" href="#/courses">See the free lessons</a></div>` : extra}</div></div>`;
document.addEventListener('click', e => {
  const q = e.target.closest('.qhelp'); if (q) { e.preventDefault(); popover(q, `<p>${md(HELP_TIPS[q.dataset.help] || '')}</p>`); return; }
  const g = e.target.closest('.gl'); if (g) { showGloss(g); return; }
  const a = e.target.closest('[data-ask]'); if (a) { const k = a.dataset.ask; askTeacher(k === 'login' ? { subject: 'Sheetah login', what: 'Could you give me a Sheetah student login, please?' } : k === 'password' ? { subject: 'Sheetah password', what: 'I’ve forgotten my Sheetah password. Could you set a new one for me?' } : { where: a.dataset.where || location.hash, what: a.dataset.what || '' }); }
});
document.addEventListener('mouseover', e => { const g = e.target.closest('.gl'); if (g && matchMedia('(hover: hover)').matches && !g._shown) { g._shown = true; setTimeout(() => { g._shown = false; }, 1200); showGloss(g); } });
document.addEventListener('keydown', e => { const g = e.target.closest && e.target.closest('.gl'); if (g && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); showGloss(g); } });
const typing = e => { const el = e.target; return el && (el.closest('input,textarea,select,[contenteditable="true"]') || el.closest('.xg')); };
addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); openSearch(); return; }
  if (e.key === 'Escape' && document.body.classList.contains('focus') && !document.querySelector('.modal-wrap')) { setFocus(false); return; }
  if (typing(e) || e.ctrlKey || e.metaKey || e.altKey || document.querySelector('.modal-wrap,.tour')) return;
  if (e.key === '/') { e.preventDefault(); openSearch(); return; }
  if (e.key === '[') { e.preventDefault(); toggleFold(); return; }
  if (e.key === '?') { e.preventDefault(); const k = ARTICLES.find(x => x.id === 'keys'); modal(`<h2>${esc(k.title)}</h2>${k.body.map(p => `<p>${md(p)}</p>`).join('')}`); return; }
  if (keys.handler) keys.handler(e);
});
document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && document.body.classList.contains('focus') && keys.fullscreenExit) keys.fullscreenExit(); });

register('notfound', async main => { main.innerHTML = `<section class="card empty">${mascot('oops')}<div><h1>${t('Page not found')}</h1><p class="lead">This link doesn’t lead anywhere in Sheetah. It may be old, or it may have been typed slightly wrong.</p><div class="row"><a class="btn" href="#/home">${t('Go home')}</a><button class="btn ghost" type="button" id="nfS">${t('Search')}</button><a class="btn ghost" href="#/help">${t('Help')}</a></div></div></section>`; $('#nfS').onclick = () => openSearch(); });
register('welcome', async () => { location.replace('#/setup'); });

// ----- student sign-in -----
register('login', async main => {
  main.innerHTML = `<section class="card login">${mascot('hello')}<div><h1>Student sign-in</h1><p>Use the email and password your teacher gave you. There is nothing to buy: your teacher gives you access.</p>
  <form id="lf" class="form"><label>Email <input type="email" name="email" required autocomplete="username"></label><label>Password <span class="pwrow"><input type="password" name="pw" required autocomplete="current-password"><button type="button" class="btn ghost sm" id="pwShow">Show</button></span></label>
  <label class="ck"><input type="checkbox" name="rem" checked> Keep me signed in on this laptop</label><p class="err" id="lerr" role="alert"></p><div class="row"><button class="btn">Sign in</button><a class="btn ghost" href="#/home">Continue with free access</a></div></form>
  <p class="small"><button class="linkbtn" type="button" data-ask="password">Forgot your password?</button> · <button class="linkbtn" type="button" data-ask="login">Don’t have a login yet?</button></p>
  <p class="muted small">Your progress is saved in this browser. If you studied before signing in, Sheetah offers to bring that progress with you.</p></div></section>`;
  $('#pwShow').onclick = () => { const i = $('#lf [name=pw]'); i.type = i.type === 'password' ? 'text' : 'password'; $('#pwShow').textContent = i.type === 'password' ? 'Show' : 'Hide'; };
  $('#lf').onsubmit = async e => {
    e.preventDefault(); const f = new FormData(e.target); const btn = e.target.querySelector('button.btn:not(.ghost)'); btn.disabled = true; btn.textContent = 'Checking…'; $('#lerr').textContent = '';
    try {
      const p = await login(f.get('email'), f.get('pw')); const guest = JSON.parse(JSON.stringify(store.d)); const guestHas = Object.keys(guest.lessons).length + Object.keys(guest.cards).length > 0;
      session.set({ id: p.id, email: p.email, name: p.name, role: p.role, keys: p.keys, adminKey: p.adminKey || null, expires: p.expires || null, open: !!p.open }, !!f.get('rem')); await topUpAdminKeys();
      store.load(session.profileId()); const fresh = !Object.keys(store.d.lessons).length && !Object.keys(store.d.cards).length;
      if (guestHas && fresh && await confirmBox(`<b>Bring your progress with you?</b><br>You did some lessons on this laptop before signing in (${Object.values(guest.lessons).filter(l => l.status === 'done').length} finished). Sheetah can copy that progress into your login.`, 'Bring it across', 'Start fresh')) { Object.assign(store.d, guest); }
      if (!store.d.name) store.d.name = (p.name || '').split(' ')[0] || guest.name; store.save(); applyTheme();
      toast(`Welcome, ${esc(p.name || p.email)}.`); go(!store.d.setupDone ? '#/setup' : p.role === 'admin' ? '#/admin' : '#/home');
    } catch (err) { $('#lerr').textContent = err.message; btn.disabled = false; btn.textContent = 'Sign in'; }
  };
});

// The course the learner is working on now (set when they open one of its lessons). Home, the study plan and notes follow it.
export const liveCourses = () => ((app.catalog && app.catalog.courses) || []).filter(x => x.status === 'live');
export const activeCourseId = () => { const a = store.d.active; const live = liveCourses().map(x => x.id); return a && (!live.length || live.includes(a)) ? a : 'c1'; };
export function setActiveCourse(id) { if (id && store.d.active !== id && !store.preview) { store.d.active = id; store.save(); } }
export async function getCourse(id) { return loadCourse(id || activeCourseId()); }
// Each course key is also published encrypted with the admin key, so the admin's login opens a new course
// by itself, without a new password or a new students.json.
export async function topUpAdminKeys() {
  if (!session.isAdmin || !session.s.adminKey) return;
  const cat = app.catalog || await catalog().catch(() => ({ courses: [] })); let changed = false;
  for (const x of (cat.courses || []).filter(x => x.status === 'live' && !session.has(x.id))) {
    try { const r = await fetch(`content/${x.id}/s/admin.bin`, { cache: 'no-cache' }); if (!r.ok) continue; const raw = await decryptBytes(await aesKey(session.s.adminKey), await r.arrayBuffer()); session.s.keys = { ...session.s.keys, [x.id]: b64(raw) }; changed = true; } catch (e) { /* not published yet */ }
  }
  if (changed) session.set(session.s, !!localStorage.getItem('sheetah:session'));
}
export async function boot() {
  session.load(); store.load(session.profileId()); applyTheme(); matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
  shell(); applyTheme();
  const lib = getJSON('content/library.json').then(l => { setKeyMap(l.shortcuts); setGlossary(l.glossary); }).catch(() => {});
  await Promise.all([Promise.race([loadSite(), new Promise(r => setTimeout(r, 1500))]), lib]);
  app.catalog = await catalog().catch(() => ({ courses: [] })); await topUpAdminKeys();
  store.on(() => refreshChrome()); addEventListener('hashchange', route); await route();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
  reminders();
}
// a gentle reminder at the learner's chosen time while the app is open (one a day at most)
function reminders() {
  setInterval(() => {
    const s = store.d.settings, p = store.d.plan; if (!s.remind || store.preview) return; const now = new Date(); const hm = now.toTimeString().slice(0, 5); const k = 'sheetah:reminded:' + today();
    if (!p.days.includes(now.getDay()) || hm < p.time || localStorage.getItem(k) || store.studied(today())) return; localStorage.setItem(k, '1');
    import('./ui.js').then(({ line }) => { const msg = line('idle'); if ('Notification' in window && Notification.permission === 'granted') new Notification('Sheetah', { body: msg, icon: 'img/icon-192.png' }); else toast(mascot('sleepy', 'xs') + ' ' + msg); });
  }, 60000);
}
