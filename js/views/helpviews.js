// The help centre, the "How Sheetah works" picture guide (with the one-minute video) and the A–Z topic index.
import { register, getCourse } from '../app.js';
import { store } from '../store.js';
import { esc, md, $, $$, mascot, crumbs } from '../ui.js';
import { ARTICLES, askTeacher } from '../help.js';
import { lessonList, canOpenLesson, lessonLabel, moduleInfo } from '../access.js';
import { t } from '../i18n.js';

const GROUPS = [['Getting started', ['start', 'unlock', 'saved', 'keys']], ['When you’re stuck', ['notquite', 'upload', 'noexcel', 'teacher']], ['Studying', ['review', 'mistakes', 'tests', 'certificate']], ['Your laptop', ['mac', 'install', 'move', 'privacy']]];
const find = id => ARTICLES.find(a => a.id === id);
register('help', async (main, id) => {
  if (id) {
    const a = find(id); if (!a) { main.innerHTML = `<section class="card empty">${mascot('oops')}<div><h1>No help page by that name</h1><a class="btn" href="#/help">All help topics</a></div></section>`; return; }
    const g = GROUPS.find(([, ids]) => ids.includes(id)); const sib = g ? g[1].filter(x => x !== id).map(find) : [];
    main.innerHTML = `${crumbs([[t('Help'), '#/help'], [a.title]])}<article class="card help-article"><h1>${esc(a.title)}</h1>${a.body.map(p => `<p>${md(p)}</p>`).join('')}${a.img ? `<figure class="shot guide"><img src="${a.img}" alt="${esc(a.title)}: a picture of this part of Sheetah" loading="lazy" onerror="this.closest('figure').remove()"></figure>` : ''}
      <div class="row"><span class="muted">Still not sure?</span><button class="btn ghost sm" type="button" id="hAsk">${t('Ask your teacher')}</button><a class="btn ghost sm" href="#/how">How Sheetah works</a></div></article>
      ${sib.length ? `<section class="card"><h2>Related</h2><ul class="plist">${sib.map(s => `<li><a href="#/help/${s.id}">${esc(s.title)}</a></li>`).join('')}</ul></section>` : ''}`;
    $('#hAsk').onclick = () => askTeacher({ where: 'Help: ' + a.title, what: 'I read the help page but I’m still not sure about this: ' });
    return;
  }
  main.innerHTML = `<h1>${t('Help')}</h1><p class="lead">Short answers to what students ask most. Type a few words, or pick a topic.</p>
    <div class="help-search"><input id="hq" type="search" placeholder="e.g. upload, locked, Mac, backup" aria-label="Search help"></div>
    <div class="grid3 help-quick"><a class="card quick" href="#/how">${mascot('teaching', 'xs')}<b>How Sheetah works</b><small>The one-page picture guide and a one-minute video</small></a><a class="card quick" href="#/tour"><span class="qi" aria-hidden="true">➜</span><b>${t('Take the tour')}</b><small>Seven stops around the app</small></a><button class="card quick" type="button" id="askT"><span class="qi" aria-hidden="true">✉</span><b>${t('Ask your teacher')}</b><small>A ready message on WhatsApp or email</small></button></div>
    <div id="hres" hidden></div>
    <div id="hgroups" class="grid2">${GROUPS.map(([title, ids]) => `<section class="card"><h2>${title}</h2><ul class="plist">${ids.map(find).filter(Boolean).map(a => `<li><a href="#/help/${a.id}">${esc(a.title)}</a></li>`).join('')}</ul></section>`).join('')}</div>
    <p class="muted small">Tip: anywhere in Sheetah, press <kbd>?</kbd> for the keyboard keys, or <kbd>Ctrl</kbd>+<kbd>K</kbd> to search lessons, functions and settings. Small <span class="qhelp static">?</span> buttons next to things explain them in two lines.</p>`;
  $('#askT').onclick = () => askTeacher({ where: 'Help centre' });
  const plain = s => s.replace(/\*\*|`/g, '').toLowerCase();
  $('#hq').oninput = e => { const q = e.target.value.trim().toLowerCase(); const box = $('#hres'); $('#hgroups').hidden = !!q; box.hidden = !q; if (!q) return;
    const words = q.split(/\s+/); const hits = ARTICLES.map(a => { const hay = plain(a.title + ' ' + a.tags + ' ' + a.body.join(' ')); const score = words.reduce((s, w) => s + (hay.includes(w) ? 1 : 0) + (plain(a.title).includes(w) ? 2 : 0), 0); return [a, score]; }).filter(([, s]) => s > 0).sort((a, b) => b[1] - a[1]);
    box.innerHTML = hits.length ? `<section class="card"><h2>${hits.length} answer${hits.length === 1 ? '' : 's'}</h2><ul class="help-hits">${hits.map(([a]) => { const p = a.body.find(b => words.some(w => plain(b).includes(w))) || a.body[0]; return `<li><a href="#/help/${a.id}"><b>${esc(a.title)}</b><small>${md(p).slice(0, 220)}</small></a></li>`; }).join('')}</ul></section>` : `<section class="card empty-in">${mascot('thinking', 'sm')}<p>No help page matches “${esc(q)}”. Try other words, <button class="linkbtn" type="button" id="hs2">search the lessons</button>, or ask your teacher.</p></section>`;
    const h2 = $('#hs2'); if (h2) h2.onclick = () => import('../search.js').then(m => m.openSearch(q)); };
});

// S5 + S6: the one-page picture guide
register('how', async main => {
  const sec = (img, title, body, link) => `<section class="card how-sec"><div class="how-text"><h2>${title}</h2>${body.map(p => `<p>${md(p)}</p>`).join('')}${link ? `<p><a href="${link[1]}">${link[0]} →</a></p>` : ''}</div>${img ? `<figure class="shot guide"><img src="img/guide/${img}.png" alt="${esc(title)}" loading="lazy" onerror="this.closest('figure').remove()"></figure>` : ''}</section>`;
  main.innerHTML = `${crumbs([[t('Help'), '#/help'], ['How Sheetah works']])}<h1>How Sheetah works</h1><p class="lead">Everything in one page: how a lesson runs, how modules open, the daily review, the mistakes log, and what’s saved where.</p>
    <section class="card how-video"><h2>The one-minute tour</h2><video controls preload="none" poster="img/guide/tour-poster.png" width="960" height="600"><source src="img/guide/tour.mp4" type="video/mp4">Your browser can’t play this video.</video><p class="muted small">A screen recording of a first lesson, with no sound. Prefer to click around? <a href="#/tour">Take the guided tour</a>.</p></section>
    ${sec('lesson', '1. Every lesson has five parts', ['**Lecture**: short slides with real Excel screenshots; click any screenshot to enlarge it. Each segment ends with one quick check.', '**Demo**: watch Sheetah do it on a live sheet, or step through it.', '**Lab**: your turn, on a practice sheet. Press **Check**, use hints if you need them, and see exactly what was checked.', '**Quiz** and **Recap**: a few questions, then the cards that go into your daily review.', 'Parts open in order. The ☰ Outline button lists every slide and task.'], ['Start lesson 1', '#/lesson/c1/1-1/lecture'])}
    ${sec('map', '2. Modules open as you pass tests', ['Your **learning map** shows every module as a stop. The diamonds between stops are module tests.', 'Modules 1 and 2 are free for everyone. With a student login, passing a module test (70% or more) opens the next module.', 'Already know a module? **Test out**: take its test straight away.'], ['Open the learning map', '#/course/c1'])}
    ${sec('lab', '3. When a task says “Not quite”', ['Sheetah checks what you actually did: that a cell holds a formula, that it gives the right answer, and that it still works when the numbers change.', 'Read the message, open a hint (there are three), and after three tries you can see the worked solution. Tasks done with help come back in review.'], ['Help: why a task says “Not quite”', '#/help/notquite'])}
    ${sec('review', '4. Ten minutes of review a day', ['Each lesson adds cards to your review deck. Cards come back just before you’d forget them.', 'Your mistakes come back too, after 1, 3 and 7 days, until they stick. The top bar shows how many are waiting.'], ['Go to Daily review', '#/review'])}
    ${sec('upload', '5. Real Excel tasks', ['Most lessons end with an optional task in real Excel. Download the workbook, do the steps, save as .xlsx and upload it. Sheetah checks it at once.', 'No Excel? The free Excel for the web works for almost everything.'], ['Help: uploading a file', '#/help/upload'])}
    ${sec('', '6. What’s saved where', ['Your progress, notes, journal and settings stay **in this browser on this laptop**. Nothing is sent anywhere.', 'To move laptops, or to keep a copy, download a backup from Settings › Your data.', 'Your login only unlocks the lessons. Your teacher sees your progress only if you send it.'], ['Settings › Your data', '#/settings/data'])}
    <div class="row center"><a class="btn big" href="#/home">${t('Go home')}</a><a class="btn ghost" href="#/help">More help</a></div>`;
});

// N5: every topic the course teaches, A–Z, each linking to its exact slide
register('topics', async (main, mod = 'all') => {
  const c = await getCourse('c1'); const all = lessonList(c);
  const rows = []; for (const l of all) for (const [si, k, h] of (l.topics || [])) rows.push({ h: h.replace(/\*\*|`/g, ''), l, href: `#/lesson/c1/${l.id}/lecture/${si + 1}-${k + 1}` });
  const shown = rows.filter(r => mod === 'all' || String(r.l.module.n) === mod).sort((a, b) => a.h.localeCompare(b.h, 'en', { sensitivity: 'base' }));
  const L0 = r => /[a-z]/i.test(r.h[0]) ? r.h[0].toUpperCase() : '#'; shown.sort((a, b) => (L0(a) === '#') - (L0(b) === '#') || 0); const letters = [...new Set(shown.map(L0))].sort((a, b) => a === '#' ? -1 : b === '#' ? 1 : a.localeCompare(b));
  main.innerHTML = `<h1>${t('Topic index')}</h1><p class="lead">${rows.length} topics from ${esc(c.title)}, A to Z. Each one opens the exact slide that teaches it.</p>
    <div class="row wrap"><div class="seg">${[['all', 'All modules'], ...c.modules.map(m => [String(m.n), 'Module ' + m.n])].map(([v, l]) => `<a class="segb ${v === mod ? 'on' : ''}" href="#/topics/${v}">${l}</a>`).join('')}</div><input id="tq" type="search" placeholder="Filter topics" aria-label="Filter topics" class="search"></div>
    <nav class="az" aria-label="Jump to letter">${letters.map(L => `<a href="#" data-l="${L}">${L}</a>`).join('')}</nav>
    <section class="card"><div class="topic-az">${letters.map(L => `<div class="az-group" id="az-${L}"><h2>${L === '#' ? '# Symbols and errors' : L}</h2><ul>${shown.filter(r => L0(r) === L).map(r => { const open = canOpenLesson(c, r.l); return `<li data-q="${esc(r.h.toLowerCase())}">${open ? `<a href="${r.href}">${esc(r.h)}</a>` : `<span class="muted">${esc(r.h)} 🔒</span>`}<small>${esc(lessonLabel(r.l))} · ${esc(r.l.title)}${open ? ` · <a href="#/practice/c1/${r.l.id}">practise</a>` : ''}</small></li>`; }).join('')}</ul></div>`).join('')}</div></section>`;
  $$('.az a').forEach(a => a.onclick = e => { e.preventDefault(); const el = document.getElementById('az-' + a.dataset.l); el && el.scrollIntoView({ block: 'start', behavior: 'smooth' }); });
  $('#tq').oninput = e => { const q = e.target.value.trim().toLowerCase(); $$('.topic-az li').forEach(li => li.hidden = !!q && !li.dataset.q.includes(q)); $$('.az-group').forEach(g => g.hidden = !$$('li', g).some(li => !li.hidden)); };
});
