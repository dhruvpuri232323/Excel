// Things students asked for beyond the lessons: saved items, all their notes (downloadable as Word),
// practice on any lesson with fresh numbers, a weekly real-world challenge, interview prep and every download in one place.
import { register, getCourse, go, lockedPanel } from '../app.js';
import { store, today } from '../store.js';
import { lesson as loadLesson, getJSON, assetURL } from '../content.js';
import { esc, md, $, $$, mascot, toast, modal, crumbs, shuffle, download, fmtDate, confirmBox, qhelp, confetti } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask, applySolution, describeCheck, randomize } from '../checker.js';
import { lessonList, canOpenLesson, lessonLabel, isDone, moduleInfo } from '../access.js';
import { hydrate } from './lesson.js';
import { runTest, shuffleOpts } from './tests.js';
import { t as tr } from '../i18n.js';

// ---------- saved items (X2) ----------
export const isSaved = key => (store.d.saved || []).some(x => x.key === key);
export function toggleSaved(item) { const s = store.d.saved || (store.d.saved = []); const i = s.findIndex(x => x.key === item.key); if (i >= 0) s.splice(i, 1); else s.unshift({ ...item, at: new Date().toISOString() }); store.save(); return i < 0; }
const savedHref = x => x.type === 'lesson' ? `#/lesson/${x.course}/${x.lesson}` : x.type === 'slide' ? `#/lesson/${x.course}/${x.lesson}/lecture/${x.pos}` : x.type === 'task' ? `#/lesson/${x.course}/${x.lesson}/lab` : x.type === 'function' ? `#/library/${encodeURIComponent(x.name)}` : x.type === 'challenge' ? `#/challenge/${x.id}` : '#/home';
register('saved', async main => {
  const c = await getCourse(); const all = lessonList(c); const lab = id => { const l = all.find(x => x.id === id); return l ? lessonLabel(l) + ' · ' + l.title : ''; };
  const groups = [['lesson', 'Lessons'], ['slide', 'Slides'], ['task', 'Lab tasks'], ['function', 'Functions'], ['challenge', 'Challenges']];
  const s = store.d.saved || [];
  main.innerHTML = `<h1>${tr('Saved')}</h1><p class="lead">Everything you starred, in one place. Look for <b>☆</b> on lessons, slides, lab tasks, functions and challenges.</p>
   ${s.length ? groups.filter(([k]) => s.some(x => x.type === k)).map(([k, label]) => `<section class="card"><h2>${label}</h2><ul class="saved-list">${s.filter(x => x.type === k).map(x => `<li><a href="${savedHref(x)}"><b>${md(x.title || x.name || '')}</b>${x.lesson && k !== 'lesson' ? `<small>${esc(lab(x.lesson))}${x.pos ? ' · slide ' + esc(x.pos) : ''}</small>` : ''}</a><span class="muted small">${fmtDate(x.at)}</span><button class="mini" type="button" data-un="${esc(x.key)}" aria-label="Remove from Saved">✕</button></li>`).join('')}</ul></section>`).join('')
    : `<section class="card empty">${mascot('thinking')}<div><h2>Nothing saved yet</h2><p>Tap <b>☆</b> on anything you want to come back to: a lesson, a single slide, a lab task, a function or a challenge. It appears here.</p><a class="btn" href="#/course/c1">Go to the course</a></div></section>`}`;
  $$('[data-un]').forEach(b => b.onclick = () => { const it = s.find(x => x.key === b.dataset.un); toggleSaved(it); b.closest('li').remove(); toast('Removed from Saved'); });
});

// ---------- my notes (X1) ----------
register('notes', async main => {
  const c = await getCourse(); const all = lessonList(c); const ids = all.filter(l => (store.d.notes[l.id] || '').trim()).map(l => l.id);
  const extra = Object.keys(store.d.notes).filter(k => k === 'general' && store.d.notes[k].trim());
  main.innerHTML = `<div class="row between"><h1>${tr('My notes')}</h1>${ids.length ? `<button class="btn" id="docx" type="button">⬇ Download as Word</button>` : ''}</div><p class="lead">Every note you write in a lesson (✐ Notes) gathers here. Edit them right here; they save as you type.</p>
   <section class="card"><h2>General notes</h2><textarea class="notes" data-n="general" rows="4" placeholder="Anything that isn’t about one lesson: questions for your teacher, things to try at work…">${esc(store.d.notes.general || '')}</textarea></section>
   ${ids.length ? c.modules.map(m => { const ls = m.lessons.filter(l => ids.includes(l.id)); return ls.length ? `<section class="card"><h2>Module ${m.n}: ${esc(m.title)}</h2>${ls.map(l => `<div class="note-block"><div class="row between"><h3>${esc(lessonLabel({ ...l, module: m }))} · ${esc(l.title)}</h3><a class="small" href="#/lesson/c1/${l.id}">Open lesson</a></div><textarea class="notes" data-n="${l.id}" rows="5">${esc(store.d.notes[l.id])}</textarea>${store.d.notesAt && store.d.notesAt[l.id] ? `<p class="muted small">Last changed ${fmtDate(store.d.notesAt[l.id])}</p>` : ''}</div>`).join('')}</section>` : ''; }).join('')
    : `<section class="card empty">${mascot('teaching', 'sm')}<div><h2>No lesson notes yet</h2><p>In any lesson, click <b>✐ Notes</b> to jot down what you want to remember, such as “F4 cycles $A$1 → A$1 → $A1”. Notes from every lesson collect here, and you can download them all as one Word file.</p></div></section>`}`;
  let tm; $$('textarea[data-n]').forEach(t => t.oninput = () => { clearTimeout(tm); tm = setTimeout(() => { store.d.notes[t.dataset.n] = t.value; store.d.notesAt = store.d.notesAt || {}; store.d.notesAt[t.dataset.n] = new Date().toISOString(); store.save(); }, 500); });
  const d = $('#docx'); if (d) d.onclick = async () => { const paras = [['Title', 'My Sheetah notes'], ['Normal', `${store.d.name || ''} · ${fmtDate(new Date().toISOString())}`]]; if ((store.d.notes.general || '').trim()) { paras.push(['Heading1', 'General notes']); store.d.notes.general.split('\n').forEach(p => paras.push(['Normal', p])); }
    for (const m of c.modules) { const ls = m.lessons.filter(l => (store.d.notes[l.id] || '').trim()); if (!ls.length) continue; paras.push(['Heading1', `Module ${m.n}: ${m.title}`]); for (const l of ls) { paras.push(['Heading2', `Lesson ${l.num.split('.')[1]}: ${l.title}`]); store.d.notes[l.id].split('\n').forEach(p => paras.push(['Normal', p])); } }
    try { download(`My Sheetah notes ${today()}.docx`, await docx(paras)); } catch (e) { toast('Couldn’t make the Word file: ' + esc(e.message)); } };
});
// a minimal Word file: a title, headings and paragraphs
async function docx(paras) {
  if (!window.JSZip) throw new Error('the zip library didn’t load');
  const x = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const z = new JSZip(); const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  const st = (id, name, sz, bold, color, after) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="${after}"/></w:pPr><w:rPr>${bold ? '<w:b/>' : ''}<w:color w:val="${color}"/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
  z.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${st('Title', 'Title', 44, true, '23263A', 120)}${st('Heading1', 'heading 1', 32, true, '3557D6', 120)}${st('Heading2', 'heading 2', 26, true, '23263A', 80)}</w:styles>`);
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${paras.map(([s, tx]) => `<w:p>${s !== 'Normal' ? `<w:pPr><w:pStyle w:val="${s}"/></w:pPr>` : ''}<w:r><w:t xml:space="preserve">${x(tx)}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// ---------- a board of tasks on a live sheet, shared by practice and challenges ----------
function taskBoard(host, { sheet, tasks, editable, inputs, allowSort, allowStructure, height = 360, state, courseId, lessonId, lessonNum, onChange }) {
  const grids = []; const gridTasks = tasks.filter(t => !['mc', 'hotspot', 'order'].includes(t.kind)); state.tasks = state.tasks || {};
  host.innerHTML = `<div class="lab-cols ${gridTasks.length ? '' : 'nogrid'}"><ol class="tasks">${tasks.map((t, i) => { const s = state.tasks[t.id] || {}; let extra = '';
      if (t.kind === 'mc') extra = `<div class="opts">${t.options.map((o, k) => `<button class="opt" data-i="${k}" type="button">${md(o)}</button>`).join('')}</div>`;
      if (t.kind === 'hotspot') extra = `<div class="hotspot"><img data-asset="${t.img}" alt="${esc(t.alt || 'Excel window screenshot')}"><span class="hotmark" hidden></span></div>`;
      if (t.kind === 'order') extra = `<ul class="order">${shuffle(t.items).map(v => `<li data-v="${esc(v)}"><span>⋮⋮</span> ${md(v)} <button class="mini" data-d="up" aria-label="Move up" type="button">↑</button><button class="mini" data-d="down" aria-label="Move down" type="button">↓</button></li>`).join('')}</ul>`;
      return `<li class="card task ${s.done ? 'done' : ''}" id="pt-${i}"><div class="tnum">${i + 1}</div><div class="tbody"><p>${md(t.prompt)}</p>${extra}<div class="row"><button class="btn sm check" type="button">${tr('Check')}</button><button class="btn ghost sm hint" type="button">${tr('Hint')} 1 of ${t.hints.length}</button><button class="btn ghost sm sol" type="button" hidden>${tr('Show solution')}</button></div><ul class="hints"></ul><div class="tfb" role="status">${s.done ? '<p class="ok">✓ Done</p>' : ''}</div></div></li>`; }).join('')}</ol>${gridTasks.length ? '<div class="lab-grid"><div class="pg"></div></div>' : ''}</div>`;
  let g = null; const keyLog = []; let usedMouse = false;
  if (gridTasks.length) { const wb = new Workbook(sheet); g = new Grid($('.pg', host), wb, { height, editable, highlight: editable, allowSort, allowStructure, noAddSheet: true, onKey: k => keyLog.push(k), onChange }); grids.push(g); $('.pg', host).addEventListener('mousedown', () => { usedMouse = true; }, true); }
  const answers = {};
  tasks.forEach((t, i) => {
    const card = $(`#pt-${i}`, host); const ts = () => state.tasks[t.id] || (state.tasks[t.id] = { hints: 0, attempts: 0 });
    if (t.kind === 'mc') $$('.opt', card).forEach(b => b.onclick = () => { $$('.opt', card).forEach(x => x.classList.remove('sel')); b.classList.add('sel'); answers[t.id] = +b.dataset.i; });
    if (t.kind === 'hotspot') { const img = $('img', card); img.onclick = e => { const r = img.getBoundingClientRect(); const x = (e.clientX - r.left) / r.width * 100, y = (e.clientY - r.top) / r.height * 100; const z = t.zones.find(z => x >= z.x && x <= z.x + z.w && y >= z.y && y <= z.y + z.h); answers[t.id] = z ? z.label : 'empty space'; const m = $('.hotmark', card); m.style.left = x + '%'; m.style.top = y + '%'; m.hidden = false; }; }
    if (t.kind === 'order') { const ul = $('.order', card); const sync = () => { answers[t.id] = $$('li', ul).map(li => li.dataset.v); }; sync(); $$('li', ul).forEach(li => $$('button', li).forEach(b => b.onclick = () => { const sib = b.dataset.d === 'up' ? li.previousElementSibling : li.nextElementSibling; if (sib) { b.dataset.d === 'up' ? ul.insertBefore(li, sib) : ul.insertBefore(sib, li); sync(); } })); }
    const showHints = s => { $('.hints', card).innerHTML = t.hints.slice(0, s.hints).map((h, k) => `<li><b>Hint ${k + 1}.</b> ${md(h)}</li>`).join(''); $('.hint', card).textContent = s.hints < t.hints.length ? `${tr('Hint')} ${s.hints + 1} of ${t.hints.length}` : 'No more hints'; if (s.hints >= t.hints.length) $('.sol', card).hidden = false; };
    $('.check', card).onclick = () => { const s = ts(); s.attempts++; const res = checkTask(t, { wb: g && g.wb, grid: g, inputs, answer: answers[t.id], keyLog, usedMouse }); const what = describeCheck(t, inputs); const fb = $('.tfb', card);
      const whatHTML = what.length ? `<details class="what"><summary>What Sheetah checked</summary><ul>${what.map(w => `<li>${esc(w)}</li>`).join('')}</ul></details>` : '';
      if (res.ok) { s.done = true; fb.innerHTML = `<p class="ok">✓ ${md(res.msg)}</p>${whatHTML}`; card.classList.add('done'); store.act('tasks'); }
      else { fb.innerHTML = `<p class="no">${mascot('oops', 'xs')} ${md(res.msg)}</p>${whatHTML}`; card.classList.add('shake'); setTimeout(() => card.classList.remove('shake'), 400); if (store.d.settings.autoHint && s.hints < t.hints.length) { s.hints++; showHints(s); } if (s.attempts >= 3) $('.sol', card).hidden = false; }
      store.save(); host.dispatchEvent(new CustomEvent('progress')); };
    $('.hint', card).onclick = () => { const s = ts(); if (s.hints < t.hints.length) s.hints++; showHints(s); store.save(); };
    $('.sol', card).onclick = () => { const m = modal(`<h2>Worked solution</h2>${t.solution ? `<p><code class="big">${esc(t.solution)}</code></p>` : ''}<p>${md(t.explain || '')}</p>${g && !['mc', 'hotspot', 'order'].includes(t.kind) ? '<div class="row end"><button class="btn ghost" id="put" type="button">Put it in the sheet for me</button><button class="btn" id="ok" type="button">I’ll do it myself</button></div>' : '<div class="row end"><button class="btn" id="ok" type="button">Got it</button></div>'}`); const p = $('#put', m.el); if (p) p.onclick = () => { applySolution(t, { wb: g.wb, grid: g }); g.render(); m.close(); }; $('#ok', m.el).onclick = m.close; };
    card.addEventListener('focusin', () => { if (g && t.focus) g.select(t.focus); });
  });
  hydrate(host, courseId);
  return { grid: g, destroy: () => grids.forEach(x => x.destroy()), done: () => tasks.filter(t => state.tasks[t.id] && state.tasks[t.id].done).length };
}

// ---------- practise any lesson (X3) ----------
let board = null; const kill = () => { if (board) { board.destroy(); board = null; } };
register('practice', async (main, courseId, id) => {
  kill(); const c = await getCourse(courseId); const all = lessonList(c);
  if (!id) { // the picker: weak spots first, then every lesson by module
    const errBy = {}; store.d.errors.filter(e => !e.resolved && e.lesson).forEach(e => { errBy[e.lesson] = (errBy[e.lesson] || 0) + 1; });
    const weak = Object.entries(errBy).sort((a, b) => b[1] - a[1]).map(([lid, n]) => [all.find(l => l.id === lid), n]).filter(([l]) => l && canOpenLesson(c, l)).slice(0, 4);
    main.innerHTML = `${crumbs([[c.title, '#/course/' + c.id], ['Practise']])}<h1>Practise any lesson</h1><p class="lead">Fresh copies of a lesson’s lab tasks, with new numbers wherever the task allows, plus a few of its quiz questions in a new order. Nothing here changes your lesson progress.</p>
      ${weak.length ? `<section class="card"><h2>Your weak spots</h2><p class="muted small">Lessons with the most open mistakes.</p><div class="chips">${weak.map(([l, n]) => `<a class="chip big" href="#/practice/${c.id}/${l.id}">${esc(lessonLabel(l))} · ${esc(l.title)} <small>${n} open</small></a>`).join('')}</div></section>` : ''}
      ${c.modules.map(m => { const info = moduleInfo(c, m); return `<section class="card"><h2>Module ${m.n}: ${esc(m.title)}</h2>${info.open ? `<ul class="plist">${m.lessons.map(l => { const p = store.d.practice[l.id]; return `<li><a href="#/practice/${c.id}/${l.id}">${esc(lessonLabel({ ...l, module: m }))} · ${esc(l.title)}</a>${p ? `<small class="muted">practised ${fmtDate(p.at)} · ${p.done} of ${p.total}</small>` : isDone(l.id) ? '' : '<small class="muted">lesson not finished yet</small>'}</li>`; }).join('')}</ul>` : `<p class="muted">🔒 ${info.reason === 'login' ? 'Needs a student login.' : `Opens when you pass the Module ${info.prev.n} test.`}</p>`}</section>`; }).join('')}`;
    return;
  }
  const meta = all.find(l => l.id === id); if (!meta) throw new Error('No such lesson');
  if (!canOpenLesson(c, meta)) { main.innerHTML = lockedPanel(moduleInfo(c, meta.module).reason === 'login' ? 'login' : 'test', `<h2>Module ${meta.module.n} isn’t open yet</h2><p>Practice opens with the module.</p><a class="btn" href="#/module/${c.id}/${meta.module.n}">See Module ${meta.module.n}</a>`); return; }
  const L = await loadLesson(c, id); const B = L.lab; const seed = Math.floor(Math.random() * 1e6) + 1;
  const wb = new Workbook(B.sheet); const fresh = !!(B.inputs && B.inputs.length); if (fresh) randomize(wb, 0, B.inputs, seed);
  const tasks = B.tasks.filter(t => t.kind !== 'keys'); const quiz = shuffle(L.quiz || []).slice(0, 4).map(shuffleOpts); const state = { tasks: {} };
  main.innerHTML = `${crumbs([[c.title, '#/course/' + c.id], ['Practise', '#/practice/' + c.id], [lessonLabel(meta)]])}<div class="row between"><div><p class="eyebrow">Practice · ${esc(lessonLabel(meta))}</p><h1>${esc(L.title)}</h1></div><div class="row"><button class="btn ghost" id="again" type="button">${fresh ? '↻ New numbers' : '↻ Start again'}</button><a class="btn ghost" href="#/lesson/${c.id}/${id}/lecture">Back to the lesson</a></div></div>
    <p class="lead">${fresh ? 'The numbers in this sheet are new, so the answers are too. ' : ''}${tasks.length} tasks and ${quiz.length} questions. Checks work exactly as in the lab ${qhelp('inputs')}.</p><div id="pb"></div>
    ${quiz.length ? `<section class="card"><h2>Quick questions</h2><ol class="pquiz">${quiz.map((q, i) => `<li data-q="${i}"><p>${md(q.q)}</p>${q.code ? `<p><code class="big">${esc(q.code)}</code></p>` : ''}${q.answers ? `<form class="typed row"><input class="ans" autocomplete="off" spellcheck="false" placeholder="Type your answer"><button class="btn sm">${tr('Check')}</button></form>` : `<div class="opts">${q.options.map((o, k) => `<button class="opt" data-o="${k}" type="button">${md(o)}</button>`).join('')}</div>`}<div class="feedback" role="status"></div></li>`).join('')}</ol></section>` : ''}
    <div class="row between"><span class="muted" id="pstat"></span><a class="btn" href="#/practice/${c.id}">Practise another lesson</a></div>`;
  const spec = wb.toSpec();
  board = taskBoard($('#pb'), { sheet: spec, tasks, editable: B.editable, inputs: B.inputs, allowSort: B.allowSort, allowStructure: B.allowStructure, height: B.height || 360, state, courseId: c.id, lessonId: id, lessonNum: L.num });
  let qRight = 0, qDone = 0;
  const stat = () => { const d = board.done(); $('#pstat').textContent = `${d} of ${tasks.length} tasks · ${qRight} of ${quiz.length} questions right`; store.d.practice[id] = { at: new Date().toISOString(), done: d + qRight, total: tasks.length + quiz.length }; store.save(); if (d === tasks.length && qDone === quiz.length) { toast('Practice round complete. Nicely done.'); confetti(); } };
  $('#pb').addEventListener('progress', stat);
  $$('.pquiz li').forEach(li => { const q = quiz[+li.dataset.q]; const fin = (ok, given) => { if (li.dataset.done) return; li.dataset.done = 1; qDone++; if (ok) qRight++; $('.feedback', li).innerHTML = `<p class="${ok ? 'ok' : 'no'}">${ok ? '✓ Right.' : `✗ The answer is ${md(q.answers ? q.answers[0] : q.options[q.answer])}.`} ${md(q.explain || '')}</p>`; $$('.opt', li).forEach(b => b.disabled = true); stat(); };
    $$('.opt', li).forEach(b => b.onclick = () => { const k = +b.dataset.o; b.classList.add(k === q.answer ? 'right' : 'wrong'); if (k !== q.answer) $$('.opt', li)[q.answer].classList.add('right'); fin(k === q.answer, q.options[k]); });
    const f = $('.typed', li); if (f) f.onsubmit = e => { e.preventDefault(); const v = $('.ans', li).value.trim(); if (!v) return; const n = s => s.replace(/\s+/g, '').toUpperCase(); fin(q.answers.some(a => n(a) === n(v)), v); }; });
  $('#again').onclick = () => { kill(); go(`#/practice/${c.id}/${id}`); };
  stat();
});

// ---------- weekly challenge (X4) ----------
const weekNo = (d = new Date()) => { const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day); const y = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - y) / 864e5 + 1) / 7) + t.getUTCFullYear() * 53; };
const weekRange = () => { const d = new Date(); const mon = new Date(d); mon.setDate(d.getDate() - ((d.getDay() + 6) % 7)); const sun = new Date(mon); sun.setDate(mon.getDate() + 6); const f = x => x.toLocaleDateString('en-IN', { day: 'numeric', month: 'short' }); return `${f(mon)} – ${f(sun)}`; };
register('challenge', async (main, id) => {
  kill(); const data = await getJSON('content/c1/challenges.json'); const list = data.challenges; const cur = list[weekNo() % list.length];
  const ch = id ? list.find(x => x.id === id) : cur; if (!ch) throw new Error('No such challenge');
  const st = store.d.challenges[ch.id] || (store.d.challenges[ch.id] = { tasks: {} }); const isCur = ch === cur;
  main.innerHTML = `${crumbs([['Weekly challenge', '#/challenge'], [ch.title]])}<div class="row between"><div><p class="eyebrow">${isCur ? `This week’s challenge · ${weekRange()}` : 'From the archive'}</p><h1>${esc(ch.title)}</h1></div><button class="btn ghost sm" id="savc" type="button">${isSaved('ch:' + ch.id) ? '★ Saved' : '☆ Save'}</button></div>
    <section class="card story"><div class="row">${mascot('focus', 'sm')}<div class="grow"><p>${md(ch.story)}</p><p class="small"><b>Skills:</b> ${ch.skills.map(esc).join(' · ')}</p></div></div><div class="row"><button class="btn ghost" data-file="${ch.file}" type="button">⬇ Workbook to try in Excel</button><button class="btn ghost" id="solB" type="button" ${st.done ? '' : 'hidden'}>⬇ Worked solution</button><button class="linkbtn" id="showSol" type="button" ${st.done ? 'hidden' : ''}>Stuck? Show the worked solution</button></div></section>
    <div id="cb"></div><section class="card" id="explain" ${st.done ? '' : 'hidden'}><h2>How it works</h2><p>${md(ch.explain)}</p></section>
    <section class="card"><h2>All challenges</h2><ul class="plist">${list.map(x => { const s = store.d.challenges[x.id]; return `<li><a href="#/challenge/${x.id}">${esc(x.title)}</a>${x === cur ? ' <span class="tag">this week</span>' : ''}<small class="muted">${x.skills.map(esc).join(' · ')}${s && s.done ? ' · ✓ done' : ''}</small></li>`; }).join('')}</ul><p class="muted small">A new challenge comes round every Monday.</p></section>`;
  board = taskBoard($('#cb'), { sheet: ch.sheet, tasks: ch.tasks, editable: ch.editable, inputs: ch.inputs, allowSort: ch.allowSort, allowStructure: ch.allowStructure, height: 380, state: st, courseId: 'c1', onChange: () => {} });
  const solved = () => { $('#explain').hidden = false; $('#solB').hidden = false; $('#showSol').hidden = true; };
  $('#cb').addEventListener('progress', () => { if (!st.done && board.done() === ch.tasks.length) { st.done = true; st.at = new Date().toISOString(); store.save(); confetti(); toast('Challenge complete!'); solved(); } });
  $('#showSol').onclick = solved;
  $$('[data-file]').forEach(b => b.onclick = async () => { const a = document.createElement('a'); a.href = await assetURL('c1', b.dataset.file); a.download = `Sheetah challenge - ${ch.title}.xlsx`; a.click(); });
  $('#solB').onclick = async () => { const a = document.createElement('a'); a.href = await assetURL('c1', ch.solutionFile); a.download = `Sheetah challenge - ${ch.title} - solution.xlsx`; a.click(); };
  $('#savc').onclick = () => { toggleSaved({ type: 'challenge', key: 'ch:' + ch.id, id: ch.id, title: ch.title }); $('#savc').textContent = isSaved('ch:' + ch.id) ? '★ Saved' : '☆ Save'; };
});

// ---------- interview prep (X5) ----------
register('interview', async (main, mode) => {
  const c = await getCourse('c1'); const all = lessonList(c); const Q = (await getJSON('content/c1/interview.json')).questions; const iv = store.d.interview; iv.seen = iv.seen || {}; iv.mocks = iv.mocks || [];
  if (mode === 'mock') {
    const qs = shuffle(Q).slice(0, 10).map(q => shuffleOpts({ q: q.q, options: q.options, answer: q.answer, explain: q.a, lesson: q.lesson, model: q.a, listen: q.listen }));
    runTest(main, { title: 'Interview mock round', subtitle: '10 questions · 10 minutes', questions: qs, minutes: 10, pass: 0, kind: 'interview', record: false, course: c.id,
      onFinish: r => { iv.mocks.push({ at: r.at, pct: r.pct }); store.save(); },
      after: r => { main.innerHTML = `<section class="card test-end">${mascot(r.pct >= 70 ? 'celebrate' : 'encourage')}<div><p class="eyebrow">Interview mock round</p><h1>${r.pct}%</h1><p class="lead">${r.qRight} of ${r.nq} right. In a real interview you’d answer out loud, so read each model answer and say it in your own words.</p><div class="row"><a class="btn" href="#/interview/mock">Another round</a><a class="btn ghost" href="#/interview">All questions</a></div></div></section>
        <section class="card"><ol class="review-list">${r.review.map(x => `<li class="${x.ok ? 'ok' : 'no'}"><p><b>${md(x.q.q)}</b></p><p class="small">${x.ok ? '✓ Right' : '✗ You chose: ' + esc(x.given)}</p><p>${md(x.q.model)}</p><p class="muted small"><b>Interviewers listen for:</b> ${md(x.q.listen)}</p></li>`).join('')}</ol></section>`; if (r.pct >= 70) confetti(); scrollTo(0, 0); } });
    return;
  }
  const best = iv.mocks.length ? Math.max(...iv.mocks.map(m => m.pct)) : null;
  main.innerHTML = `<h1>${tr('Interview prep')}</h1><p class="lead">${Q.length} questions Excel interviews really ask, with a model answer and what the interviewer is listening for. Try saying your answer out loud before you open the model one.</p>
    <section class="card row between"><div><h2>Timed mock round</h2><p class="muted">10 random questions in 10 minutes${best !== null ? ` · your best so far ${best}%` : ''}.</p></div><a class="btn" href="#/interview/mock">Start a mock round</a></section>
    <section class="card"><ol class="iv-list">${Q.map((q, i) => { const l = all.find(x => x.id === q.lesson); return `<li><details data-i="${i}" ${iv.seen[i] ? 'class="seen"' : ''}><summary><b>${md(q.q)}</b>${iv.seen[i] ? ' <span class="tag">seen</span>' : ''}</summary><p>${md(q.a)}</p><p class="muted small"><b>Interviewers listen for:</b> ${md(q.listen)}</p>${l ? `<p class="small"><a href="#/lesson/${c.id}/${l.id}/lecture">Taught in ${esc(lessonLabel(l))} · ${esc(l.title)}</a></p>` : ''}</details></li>`; }).join('')}</ol></section>`;
  $$('details[data-i]').forEach(d => d.addEventListener('toggle', () => { if (d.open && !iv.seen[d.dataset.i]) { iv.seen[d.dataset.i] = 1; store.save(); } }));
  if (mode && /^\d+$/.test(mode)) { const d = $(`details[data-i="${+mode - 1}"]`); if (d) { d.open = true; setTimeout(() => d.scrollIntoView({ block: 'center' }), 50); } }
});

// ---------- every download in one place (X8) ----------
register('downloads', async main => {
  const c = await getCourse('c1'); const ch = (await getJSON('content/c1/challenges.json').catch(() => ({ challenges: [] }))).challenges;
  const ext = f => f.split('.').pop(); const item = (d, ok) => `<button class="dl ${ok ? '' : 'locked'}" type="button" ${ok ? `data-file="${esc(d.file)}" data-name="${esc(d.name || d.label)}"` : 'data-locked="1"'}><span class="ext ${ext(d.file)}">${ext(d.file).toUpperCase()}</span><span><b>${esc(d.label)}</b>${d.desc ? `<small>${esc(d.desc)}</small>` : ''}</span>${ok ? '' : '<span class="lk" aria-label="Locked">🔒</span>'}</button>`;
  main.innerHTML = `<h1>${tr('Downloads')}</h1><p class="lead">Every workbook, set of notes and cheat sheet, grouped by module. Word files open in Word or Google Docs; Excel files in Excel or the free Excel for the web.</p>
    <section class="card"><h2>${esc(c.title)}: whole course</h2><div class="dl-list">${(c.downloads || []).map(d => item(d, !d.file.startsWith('s/') || canOpenLesson(c, lessonList(c)[0]))).join('')}</div></section>
    ${c.modules.map(m => { const open = moduleInfo(c, m).open; return `<section class="card"><h2>Module ${m.n}: ${esc(m.title)}</h2>${m.lessons.map(l => (l.downloads || []).length ? `<h3>${esc(lessonLabel({ ...l, module: m }))} · ${esc(l.title)}</h3><div class="dl-list">${l.downloads.map(d => item(d, open || d.file.startsWith('a/'))).join('')}</div>` : '').join('')}${open ? '' : `<p class="muted small">🔒 ${moduleInfo(c, m).reason === 'login' ? 'These open with a student login.' : `These open with the module: pass the Module ${m.n - 1} test.`}</p>`}</section>`; }).join('')}
    <section class="card"><h2>Weekly challenges</h2><div class="dl-list">${ch.map(x => item({ file: x.file, label: x.title, desc: x.skills.join(' · '), name: `Sheetah challenge - ${x.title}.xlsx` }, true)).join('')}</div></section>`;
  $$('[data-file]').forEach(b => b.onclick = async () => { try { const a = document.createElement('a'); a.href = await assetURL(c.id, b.dataset.file); a.download = b.dataset.name; a.click(); } catch (e) { toast('Sign in to download this file.'); } });
  $$('[data-locked]').forEach(b => b.onclick = () => toast('This file opens with its module.'));
});
