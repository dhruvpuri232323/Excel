// The lesson player: lecture → guided demo → lab → quiz → recap, always in that order.
import { register, getCourse, go, lockedPanel } from '../app.js';
import { store, today } from '../store.js';
import { session, lesson as loadLesson, lessonMeta, canOpen, assetURL } from '../content.js';
import { esc, md, $, $$, mascot, line, speak, stopSpeaking, toast, confetti, modal, bar, shuffle, download } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask, applySolution } from '../checker.js';
import { readXlsx, grade } from '../xlsxcheck.js';
import { dayNum } from '../srs.js';
import { lessonList } from './home.js';

const PARTS = ['Lecture', 'Demo', 'Lab', 'Quiz', 'Recap'];
let grids = [];
const killGrids = () => { grids.forEach(g => g.destroy()); grids = []; };
export async function visual(v, course) {
  if (!v) return '';
  if (v.type === 'shot') return `<figure class="shot"><img data-asset="${v.file}" alt="${esc(v.alt || v.caption || '')}"><figcaption>${md(v.caption || '')}</figcaption></figure>`;
  if (v.type === 'grid') return `<div class="mini-grid" data-grid='${esc(JSON.stringify(v))}'></div>${v.caption ? `<p class="cap">${md(v.caption)}</p>` : ''}`;
  if (v.type === 'formula') return `<div class="formula-anat"><div class="ftext">${v.parts.map((p, i) => `<span class="fp c${i % 6}">${esc(p[0])}</span>`).join('')}</div><ul>${v.parts.filter(p => p[1]).map((p, i) => `<li><span class="fp c${v.parts.indexOf(p) % 6}">${esc(p[0])}</span> ${md(p[1])}</li>`).join('')}</ul></div>`;
  if (v.type === 'keys') return `<table class="keys"><tbody>${v.items.map(([k, w]) => `<tr><td>${md('[[' + k + ']]')}</td><td>${md(w)}</td></tr>`).join('')}</tbody></table>`;
  if (v.type === 'table') return `<table class="tbl"><thead><tr>${v.head.map(h => `<th>${md(h)}</th>`).join('')}</tr></thead><tbody>${v.rows.map(r => `<tr>${r.map(c => `<td>${md(String(c))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  if (v.type === 'tip') return `<div class="tip">${mascot('thinking', 'xs')}<p>${md(v.text)}</p></div>`;
  if (v.type === 'compare') return `<div class="compare"><div class="bad"><b>✗ ${md(v.bad.title)}</b><p>${md(v.bad.text)}</p></div><div class="good"><b>✓ ${md(v.good.title)}</b><p>${md(v.good.text)}</p></div></div>`;
  if (v.type === 'steps') return `<ol class="steps">${v.items.map(s => `<li>${md(s)}</li>`).join('')}</ol>`;
  return '';
}
export async function hydrate(root, courseId) {
  for (const img of $$('img[data-asset]', root)) { try { img.src = await assetURL(courseId, img.dataset.asset); } catch (e) { img.alt = 'Screenshot unavailable'; } img.onclick = () => modal(`<img class="zoom" src="${img.src}" alt="${esc(img.alt)}">`, { wide: true }); }
  for (const d of $$('.mini-grid[data-grid]', root)) { const v = JSON.parse(d.dataset.grid); const g = new Grid(d, v.sheet, { readOnly: true, toolbar: false, tabs: (v.sheet.sheets || []).length > 1, status: false, height: v.height || 170, select: v.select }); if (v.showFormulas) { g.showFormulas = true; g.render(); } grids.push(g); }
}

register('lesson', async (main, courseId, id) => {
  killGrids(); stopSpeaking();
  const c = await getCourse(courseId); const meta = lessonMeta(c, id); if (!meta) throw new Error('No such lesson');
  if (!canOpen(c, meta)) { main.innerHTML = lockedPanel(); return; }
  const L = await loadLesson(c, id); const st = store.lesson(id); if (st.status === 'new') { st.status = 'started'; st.startedAt = new Date().toISOString(); store.save(); }
  const all = lessonList(c); const idx = all.findIndex(l => l.id === id); const next = all[idx + 1], prev = all[idx - 1];
  const render = async () => {
    killGrids(); stopSpeaking();
    main.innerHTML = `<div class="lesson">
    <header class="lesson-head"><a href="#/course/${c.id}" class="crumb">← ${esc(c.code)} · Module ${meta.module.n}: ${esc(meta.module.title)}</a><h1><span class="num">${esc(L.num)}</span> ${esc(L.title)}</h1>
      <nav class="parts" aria-label="Lesson parts">${PARTS.map((p, i) => `<button class="part ${i === st.part ? 'on' : ''} ${i < (st.maxPart || 0) || st.status === 'done' ? 'done' : ''}" data-p="${i}" ${i > (st.maxPart || 0) && st.status !== 'done' ? 'disabled' : ''}>${i + 1}. ${p}</button>`).join('')}</nav>
      <div class="lesson-tools"><button class="btn ghost sm" id="res">Downloads &amp; screenshots</button><button class="btn ghost sm" id="notesBtn">Notes</button><button class="btn ghost sm" id="bm" aria-pressed="${store.d.bookmarks.includes(id)}">${store.d.bookmarks.includes(id) ? '★ Bookmarked' : '☆ Bookmark'}</button><button class="btn ghost sm" id="narr" aria-pressed="${store.d.settings.narration}">${store.d.settings.narration ? '🔊 Voice on' : '🔈 Voice off'}</button></div></header>
    <div id="partBody"></div></div>`;
    $$('.part').forEach(b => b.onclick = () => { st.part = +b.dataset.p; store.save(); render(); });
    $('#bm').onclick = () => { const i = store.d.bookmarks.indexOf(id); if (i >= 0) store.d.bookmarks.splice(i, 1); else store.d.bookmarks.push(id); store.save(); render(); };
    $('#narr').onclick = () => { store.d.settings.narration = !store.d.settings.narration; store.save(); render(); };
    $('#notesBtn').onclick = () => { const m = modal(`<h2>Your notes · ${esc(L.num)}</h2><textarea id="nt" rows="12" class="notes" placeholder="Write what you want to remember. Notes stay on this laptop and go into your backup.">${esc(store.d.notes[id] || '')}</textarea><div class="row end"><button class="btn" id="ns">Save</button></div>`); $('#ns', m.el).onclick = () => { store.d.notes[id] = $('#nt', m.el).value; store.save(); m.close(); toast('Notes saved'); }; };
    $('#res').onclick = () => resources(c, L);
    const body = $('#partBody');
    if (st.part === 0) await lecture(body); else if (st.part === 1) await demo(body); else if (st.part === 2) await lab(body); else if (st.part === 3) await quiz(body); else await recap(body);
  };
  const advance = () => { st.part = Math.min(4, st.part + 1); st.maxPart = Math.max(st.maxPart || 0, st.part); store.save(); render(); };

  // ---------- 1. lecture ----------
  async function lecture(body) {
    const segs = L.lecture; st.seg = Math.min(st.seg || 0, segs.length - 1); st.slide = st.slide || 0;
    const seg = segs[st.seg]; const slide = seg.slides[Math.min(st.slide, seg.slides.length - 1)]; const atCheck = st.slide >= seg.slides.length;
    body.innerHTML = `<div class="card lecture"><div class="row between"><p class="eyebrow">Segment ${st.seg + 1} of ${segs.length}: ${esc(seg.title)}</p><div class="dots">${seg.slides.map((s, i) => `<i class="${i === st.slide ? 'on' : i < st.slide ? 'done' : ''}"></i>`).join('')}<i class="q ${atCheck ? 'on' : ''}">?</i></div></div>
      ${atCheck ? checkQuestion(seg.check) : `<div class="slide">${mascot(slide.pose || 'teaching', 'slide-m')}<div class="slide-body"><h2>${md(slide.h)}</h2>${(slide.p || []).map(p => `<p>${md(p)}</p>`).join('')}${await visual(slide.visual, c.id)}</div></div>`}
      <div class="row between nav-row"><button class="btn ghost" id="back" ${st.seg === 0 && st.slide === 0 ? 'disabled' : ''}>← Back</button>${atCheck ? '' : `<button class="btn" id="fwd">${st.slide === seg.slides.length - 1 ? 'Check yourself →' : 'Next →'}</button>`}</div>
      <details class="fulltext"><summary>Read the whole lecture as text</summary>${segs.map(s => `<h3>${esc(s.title)}</h3>${s.slides.map(sl => `<h4>${md(sl.h)}</h4>${(sl.p || []).map(p => `<p>${md(p)}</p>`).join('')}`).join('')}`).join('')}</details></div>`;
    await hydrate(body, c.id);
    if (!atCheck) speak(slide.h + '. ' + (slide.p || []).join(' '));
    const save = () => store.save();
    $('#back').onclick = () => { if (st.slide > 0) st.slide--; else if (st.seg > 0) { st.seg--; st.slide = segs[st.seg].slides.length; } save(); render(); };
    if (!atCheck) $('#fwd').onclick = () => { st.slide++; save(); render(); };
    else wireCheck(body, seg.check, () => { if (st.seg < segs.length - 1) { st.seg++; st.slide = 0; save(); render(); } else { st.seg = 0; st.slide = 0; advance(); } }, 'Next segment →', st.seg === segs.length - 1 ? 'On to the demo →' : null);
  }
  function checkQuestion(q) { return `<div class="checkq">${mascot('thinking', 'slide-m')}<div><p class="eyebrow">Quick check</p><h2>${md(q.q)}</h2><div class="opts">${q.options.map((o, i) => `<button class="opt" data-i="${i}">${md(o)}</button>`).join('')}</div><div class="feedback" role="status"></div></div></div>`; }
  function wireCheck(body, q, onNext, nextLabel, lastLabel) {
    $$('.opt', body).forEach(b => b.onclick = () => {
      const i = +b.dataset.i; const fb = $('.feedback', body); $$('.opt', body).forEach(x => x.classList.remove('right', 'wrong'));
      if (i === q.answer) { b.classList.add('right'); fb.innerHTML = `<p class="ok">✓ ${md(q.explain)}</p><button class="btn" id="cont">${lastLabel || nextLabel}</button>`; $('#cont', body).onclick = onNext; speak('Correct. ' + q.explain); }
      else { b.classList.add('wrong'); fb.innerHTML = `<p class="no">${mascot('oops', 'xs')} ${md((q.why && q.why[i]) || 'Not quite. Think about what the slides said, and try another answer.')}</p>`; }
    });
  }
  // ---------- 2. demo ----------
  async function demo(body) {
    const D = L.demo;
    if (D.frames) { // a walkthrough of screenshots, for skills that happen in Excel's menus
      st.frame = st.frame || 0; const f = D.frames[Math.min(st.frame, D.frames.length - 1)];
      body.innerHTML = `<div class="card demo"><p class="eyebrow">Guided demo · step ${st.frame + 1} of ${D.frames.length}</p><p class="lead">${md(D.intro)}</p><figure class="shot big"><img data-asset="${f.img}" alt="${esc(f.say)}"><figcaption><span class="stepn">${st.frame + 1}</span> ${md(f.say)}</figcaption></figure>
       <div class="row between"><button class="btn ghost" id="b" ${st.frame ? '' : 'disabled'}>← Back</button>${st.frame < D.frames.length - 1 ? '<button class="btn" id="n">Next step →</button>' : '<button class="btn" id="n">Now you try it →</button>'}</div></div>`;
      await hydrate(body, c.id); speak(f.say);
      $('#b').onclick = () => { st.frame--; render(); }; $('#n').onclick = () => { if (st.frame < D.frames.length - 1) { st.frame++; render(); } else { st.frame = 0; advance(); } };
      return;
    }
    body.innerHTML = `<div class="card demo"><p class="eyebrow">Guided demo</p><p class="lead">${md(D.intro)}</p><div class="say" id="cap">${mascot('teaching', 'sm')}<div class="bubble">Press <b>Play</b> and watch the grid. You’ll do it yourself in the lab next.</div></div>
      <div id="dg"></div><div class="row between"><div class="row"><button class="btn" id="play">▶ Play demo</button><button class="btn ghost" id="step">Step by step</button></div><button class="btn" id="n" ${st.demoSeen ? '' : 'disabled'}>Now you try it →</button></div></div>`;
    const g = new Grid($('#dg'), new Workbook(D.sheet), { readOnly: true, toolbar: false, height: D.height || 260 }); grids.push(g);
    let i = 0, busy = false;
    const cap = t => { $('#cap .bubble').innerHTML = md(t); speak(t); };
    const runStep = async s => { g.o.readOnly = false; try { await runStep0(s); } finally { g.o.readOnly = true; } };
    const runStep0 = async s => {
      if (s.say) cap(s.say);
      if (s.do === 'select') await g.demoSelect(s.at, 900);
      else if (s.do === 'type') await g.demoType(s.at, s.text);
      else if (s.do === 'fill') await g.demoFill(s.from, s.to);
      else if (s.do === 'style') await g.demoStyle(s.at, s.style);
      else if (s.do === 'formulas') { g.showFormulas = !g.showFormulas; g.render(); await wait(900); }
      else if (s.do === 'sort') { g.select(s.at); await wait(400); g.wb.sortRegion(0, g.wb.currentRegion(0, g.sel.ar, g.sel.ac), [{ col: g.sel.ac, desc: !!s.desc }], true); g.render(); await wait(700); }
      else if (s.do === 'width') { g.sheet.widths[s.col.charCodeAt(0) - 65] = s.px; g.render(); await wait(700); }
      else if (s.do === 'insertRow') { g.wb.insertRows(0, s.row - 1, 1); g.render(); await wait(700); }
      else if (s.do === 'deleteRow') { g.wb.deleteRows(0, s.row - 1, 1); g.render(); await wait(700); }
      else if (s.do === 'name') { g.wb.defineName(s.name, s.ref); g.select(s.ref.replace(/^.*!/, '').replace(/\$/g, '')); await wait(800); }
      else if (s.do === 'sheet') { g.setSheet(s.index); await wait(800); }
      else if (s.do === 'paste') { g.select(s.from); await wait(400); g.copy(); await wait(700); g.select(s.to); await wait(400); g.paste(s.mode || 'all'); await wait(800); }
      else if (s.do === 'clear') { g.select(s.at); await wait(400); g.clearSelection(); await wait(600); }
      else if (s.do === 'replace') { g.record(); for (const sh of g.wb.sheets) {} const u = g.wb.usedRange(g.sheet); for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) { const t = g.wb.input(g.si, r, c); if (t && (s.whole ? t.toLowerCase() === s.find.toLowerCase() : t.includes(s.find))) g.wb.setInput(g.si, r, c, s.whole ? s.with : t.split(s.find).join(s.with)); } g.render(); await wait(900); }
      else await wait(1400);
      if (s.say) await wait(Math.min(2600, 900 + s.say.length * 22));
    };
    $('#play').onclick = async () => { if (busy) return; busy = true; g.wb.load(D.sheet); g.setSel(0, 0); for (const s of D.steps) await runStep(s); busy = false; st.demoSeen = true; store.save(); $('#n').disabled = false; cap(D.outro || 'That’s the whole move. Your turn.'); };
    $('#step').onclick = async () => { if (busy) return; if (i === 0) { g.wb.load(D.sheet); g.setSel(0, 0); } busy = true; await runStep(D.steps[i]); busy = false; if (!$('#step')) return; i = (i + 1) % D.steps.length; $('#step').textContent = i ? `Next step (${i + 1}/${D.steps.length})` : 'Step by step'; if (!i) { st.demoSeen = true; store.save(); $('#n').disabled = false; } };
    $('#n').onclick = () => advance();
  }
  // ---------- 3. lab ----------
  async function lab(body) {
    const B = L.lab; st.tasks = st.tasks || {};
    const gridTasks = B.tasks.filter(t => !['mc', 'hotspot', 'order'].includes(t.kind));
    body.innerHTML = `<div class="lab"><div class="card lab-intro"><p class="eyebrow">Lab · ${B.tasks.length} tasks</p><p>${md(B.intro)}</p>${gridTasks.length ? `<p class="muted small">Work in the highlighted cells. Click a cell and type; press Enter to finish. Right-click for more options. Every check runs instantly, and nothing you do here can break anything: <button class="linkbtn" id="resetLab">reset the sheet</button>.</p>` : ''}</div>
     <div class="lab-cols ${gridTasks.length ? '' : 'nogrid'}"><ol class="tasks">${B.tasks.map((t, i) => taskCard(t, i)).join('')}</ol>${gridTasks.length ? '<div class="lab-grid"><div id="lg"></div></div>' : ''}</div>
     ${B.real ? realCard(B.real) : ''}
     <div class="row between"><span class="muted" id="labStat"></span><button class="btn" id="n">On to the quiz →</button></div></div>`;
    let g = null; const keyLog = []; let usedMouse = false;
    if (gridTasks.length) {
      const saved = st.labSheet; const wb = new Workbook(saved || B.sheet);
      g = new Grid($('#lg'), wb, { height: B.height || 380, editable: B.editable, highlight: B.highlight || B.editable, allowSort: B.allowSort, allowStructure: B.allowStructure, styleAnywhere: B.styleAnywhere, noAddSheet: !B.allowAddSheet, onChange: () => { st.labSheet = wb.toSpec(); store.save(); }, onKey: k => keyLog.push(k) }); grids.push(g);
      $('#lg').addEventListener('mousedown', () => { usedMouse = true; }, true);
      $('#resetLab').onclick = () => { delete st.labSheet; store.save(); render(); };
    }
    const answers = {};
    const refresh = () => { const done = B.tasks.filter(t => st.tasks[t.id] && st.tasks[t.id].done).length; $('#labStat').textContent = `${done} of ${B.tasks.length} tasks done`; $('#n').disabled = done < B.tasks.length; $('#n').title = done < B.tasks.length ? 'Finish every task to continue' : ''; };
    B.tasks.forEach((t, i) => {
      const card = $(`#task-${i}`); const ts = () => st.tasks[t.id] || (st.tasks[t.id] = { hints: 0, attempts: 0 });
      if (t.kind === 'mc') $$('.opt', card).forEach(b => b.onclick = () => { $$('.opt', card).forEach(x => x.classList.remove('sel')); b.classList.add('sel'); answers[t.id] = +b.dataset.i; });
      if (t.kind === 'hotspot') { const img = $('img', card); img.onclick = e => { const r = img.getBoundingClientRect(); const x = (e.clientX - r.left) / r.width * 100, y = (e.clientY - r.top) / r.height * 100; const z = t.zones.find(z => x >= z.x && x <= z.x + z.w && y >= z.y && y <= z.y + z.h); answers[t.id] = z ? z.label : 'empty space'; const m = $('.hotmark', card); m.style.left = x + '%'; m.style.top = y + '%'; m.hidden = false; }; }
      if (t.kind === 'order') { const ul = $('.order', card); let drag = null; const sync = () => { answers[t.id] = $$('li', ul).map(li => li.dataset.v); }; sync(); $$('li', ul).forEach(li => { li.draggable = true; li.ondragstart = () => { drag = li; li.classList.add('drag'); }; li.ondragend = () => { li.classList.remove('drag'); sync(); }; li.ondragover = e => { e.preventDefault(); if (drag && drag !== li) { const r = li.getBoundingClientRect(); ul.insertBefore(drag, e.clientY < r.top + r.height / 2 ? li : li.nextSibling); } }; $$('button', li).forEach(b => b.onclick = () => { const sib = b.dataset.d === 'up' ? li.previousElementSibling : li.nextElementSibling; if (sib) { b.dataset.d === 'up' ? ul.insertBefore(li, sib) : ul.insertBefore(sib, li); sync(); } }); }); }
      $('.check', card).onclick = () => {
        const s = ts(); s.attempts++; const res = checkTask(t, { wb: g && g.wb, grid: g, inputs: B.inputs, answer: answers[t.id], keyLog, usedMouse });
        const fb = $('.tfb', card);
        if (res.ok) { s.done = true; if (s.first === undefined) s.first = s.hints === 0 && s.attempts === 1; fb.innerHTML = `<p class="ok">✓ ${md(res.msg)}</p>`; card.classList.add('done'); store.act('tasks'); if (s.first) bumpSkill(id); }
        else { fb.innerHTML = `<p class="no">${mascot('oops', 'xs')} ${md(res.msg)}</p>`; card.classList.add('shake'); setTimeout(() => card.classList.remove('shake'), 400);
          if (!s.logged) { s.logged = true; const cell = t.cell || (t.range ? t.range.split(':')[0] : null); store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: t.id, kind: 'lab', prompt: t.prompt, yours: cell && g ? (g.wb.input(0, ...rc(cell)) || '(blank)') : String(answers[t.id] ?? '(no answer)'), right: t.solution || (t.kind === 'mc' ? t.options[t.answer] : t.kind === 'hotspot' ? t.answer : t.kind === 'order' ? t.items.join(' → ') : ''), why: res.msg }); }
          if (s.attempts >= 3) $('.sol', card).hidden = false; }
        store.save(); refresh();
      };
      $('.hint', card).onclick = () => { const s = ts(); if (s.hints < t.hints.length) s.hints++; $('.hints', card).innerHTML = t.hints.slice(0, s.hints).map((h, k) => `<li><b>Hint ${k + 1}.</b> ${md(h)}</li>`).join(''); $('.hint', card).textContent = s.hints < t.hints.length ? `Hint ${s.hints + 1} of ${t.hints.length}` : 'No more hints'; if (s.hints >= t.hints.length) $('.sol', card).hidden = false; store.save(); };
      $('.sol', card).onclick = () => { const s = ts(); s.sawSolution = true; const m = modal(`<h2>Worked solution</h2>${t.solution ? `<p><code class="big">${esc(t.solution)}</code></p>` : ''}<p>${md(t.explain || '')}</p>${g && !['mc', 'hotspot', 'order'].includes(t.kind) ? '<div class="row end"><button class="btn ghost" id="put">Put it in the grid for me</button><button class="btn" id="ok">I’ll type it myself</button></div>' : '<div class="row end"><button class="btn" id="ok">Got it</button></div>'}`); const p = $('#put', m.el); if (p) p.onclick = () => { applySolution(t, { wb: g.wb, grid: g }); st.labSheet = g.wb.toSpec(); store.save(); m.close(); }; $('#ok', m.el).onclick = m.close; };
      const s0 = st.tasks[t.id]; if (s0 && s0.hints) { $('.hints', card).innerHTML = t.hints.slice(0, s0.hints).map((h, k) => `<li><b>Hint ${k + 1}.</b> ${md(h)}</li>`).join(''); }
      if (s0 && (s0.attempts >= 3 || s0.hints >= t.hints.length)) $('.sol', card).hidden = false;
      card.addEventListener('focusin', () => { if (g && t.focus) g.select(t.focus); });
      $('.goto', card) && ($('.goto', card).onclick = () => { if (g && t.focus) { g.select(t.focus); g.focus(); } });
    });
    if (B.real) wireReal(body, B.real);
    await hydrate(body, c.id); refresh();
    $('#n').onclick = () => advance();
  }
  function taskCard(t, i) {
    const s = (st.tasks || {})[t.id] || {};
    let extra = '';
    if (t.kind === 'mc') extra = `<div class="opts">${t.options.map((o, k) => `<button class="opt" data-i="${k}">${md(o)}</button>`).join('')}</div>`;
    if (t.kind === 'hotspot') extra = `<div class="hotspot"><img data-asset="${t.img}" alt="${esc(t.alt || 'Excel window screenshot')}"><span class="hotmark" hidden></span></div>`;
    if (t.kind === 'order') extra = `<ul class="order">${shuffle(t.items).map(v => `<li data-v="${esc(v)}"><span>⋮⋮</span> ${md(v)} <button class="mini" data-d="up" aria-label="Move up">↑</button><button class="mini" data-d="down" aria-label="Move down">↓</button></li>`).join('')}</ul>`;
    return `<li class="card task ${s.done ? 'done' : ''}" id="task-${i}" tabindex="-1"><div class="tnum">${i + 1}</div><div class="tbody"><p>${md(t.prompt)}</p>${extra}
      <div class="row"><button class="btn sm check">Check</button><button class="btn ghost sm hint">${s.hints ? (s.hints < t.hints.length ? `Hint ${s.hints + 1} of ${t.hints.length}` : 'No more hints') : `Hint 1 of ${t.hints.length}`}</button>${t.focus ? '<button class="btn ghost sm goto">Go to cell</button>' : ''}<button class="btn ghost sm sol" hidden>Show solution</button></div>
      <ul class="hints"></ul><div class="tfb" role="status">${s.done ? '<p class="ok">✓ Done</p>' : ''}</div></div></li>`;
  }
  function realCard(R) {
    const r = (st.real || {}); return `<section class="card real"><div class="row between"><div><p class="eyebrow">Real Excel task${R.optional !== false ? ' · optional but recommended' : ''}</p><h2>${md(R.title)}</h2></div>${mascot('focus', 'sm')}</div>
     <ol class="steps"><li>Download the practice workbook and open it in Excel (desktop or the free Excel for the web).</li>${R.steps.map(s => `<li>${md(s)}</li>`).join('')}<li>Save it as an Excel Workbook (.xlsx) and upload it here. Sheetah checks it at once.</li></ol>
     <div class="row"><button class="btn ghost" data-file="${R.file}" data-name="${esc(R.fileName || 'practice.xlsx')}">⬇ Practice workbook</button><label class="btn">⬆ Upload my file<input type="file" accept=".xlsx,.xls,.csv" hidden class="up"></label>${R.solution ? `<button class="btn ghost" data-file="${R.solution}" data-name="${esc((R.fileName || 'practice.xlsx').replace('.xlsx', ' - solution.xlsx'))}" ${r.done || (r.tries || 0) >= 2 ? '' : 'hidden'} class="solbtn">⬇ Model answer</button>` : ''}</div>
     <div class="results">${r.last ? resultsHTML(r.last) : ''}</div></section>`;
  }
  function wireReal(body, R) {
    $$('[data-file]', body).forEach(b => b.onclick = async () => { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); });
    $('.up', body).onchange = async e => { const f = e.target.files[0]; if (!f) return; const box = $('.results', body); box.innerHTML = '<p>Checking your workbook…</p>'; const x = await readXlsx(f); const res = grade(x, R.checks); st.real = st.real || {}; st.real.tries = (st.real.tries || 0) + 1; st.real.last = res; if (res.pct === 100) { st.real.done = true; confetti(); bumpSkill(id); } store.save(); box.innerHTML = resultsHTML(res); const sb = $('.solbtn', body); if (sb && (st.real.done || st.real.tries >= 2)) sb.hidden = false; e.target.value = ''; };
  }
  // ---------- 4. quiz ----------
  async function quiz(body) {
    const Q = L.quiz; st.q = st.q || { i: 0, right: 0, answered: [] }; const q = Q[st.q.i];
    if (st.q.i >= Q.length) { body.innerHTML = `<div class="card quiz-end">${mascot(st.q.right / Q.length >= .8 ? 'celebrate' : 'encourage')}<div><h2>You got ${st.q.right} of ${Q.length}</h2><p>${st.q.right === Q.length ? 'Perfect score.' : 'The ones you missed are now in your error log and will come back in your daily review.'}</p><div class="row"><button class="btn ghost" id="again">Try the quiz again</button><button class="btn" id="n">To the recap →</button></div></div></div>`; $('#again').onclick = () => { st.q = { i: 0, right: 0, answered: [] }; store.save(); render(); }; $('#n').onclick = () => { st.quiz = { right: st.q.right, total: Q.length }; advance(); }; return; }
    body.innerHTML = `<div class="card quiz"><div class="row between"><p class="eyebrow">Question ${st.q.i + 1} of ${Q.length}${q.type ? ' · ' + ({ predict: 'Predict the result', fix: 'Fix the formula', which: 'Which tool?', mc: 'Choose one' }[q.type] || '') : ''}</p>${bar(st.q.i / Q.length, 'thin')}</div>
      <h2>${md(q.q)}</h2>${q.grid ? `<div class="mini-grid" data-grid='${esc(JSON.stringify(q.grid))}'></div>` : ''}${q.code ? `<p><code class="big">${esc(q.code)}</code></p>` : ''}
      ${q.answers ? `<form class="typed"><input class="ans" autocomplete="off" spellcheck="false" placeholder="Type your answer"><button class="btn">Check</button></form>` : `<div class="opts">${q.options.map((o, i) => `<button class="opt" data-i="${i}">${md(o)}</button>`).join('')}</div>`}<div class="feedback" role="status"></div></div>`;
    await hydrate(body, c.id);
    const finish = (ok, given) => { const fb = $('.feedback', body); if (ok) st.q.right++; else { store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: 'quiz-' + st.q.i, kind: 'quiz', prompt: q.q, yours: given, right: q.answers ? q.answers[0] : q.options[q.answer], why: q.explain }); }
      fb.innerHTML = `<p class="${ok ? 'ok' : 'no'}">${ok ? '✓ Right.' : '✗ Not this time.'} ${md(q.explain)}</p><button class="btn" id="nx">${st.q.i < Q.length - 1 ? 'Next question →' : 'See my score →'}</button>`; $$('.opt', body).forEach(b => b.disabled = true); st.q.i++; store.save(); $('#nx').onclick = () => render(); $('#nx').focus(); };
    if (q.answers) $('.typed', body).onsubmit = e => { e.preventDefault(); const v = $('.ans', body).value.trim(); if (!v) return; const norm = s => s.replace(/\s+/g, '').toUpperCase(); finish(q.answers.some(a => norm(a) === norm(v)), v); };
    else $$('.opt', body).forEach(b => b.onclick = () => { const i = +b.dataset.i; b.classList.add(i === q.answer ? 'right' : 'wrong'); if (i !== q.answer) $$('.opt', body)[q.answer].classList.add('right'); finish(i === q.answer, q.options[i]); });
  }
  // ---------- 5. recap ----------
  async function recap(body) {
    const R = L.recap; const fresh = R.cards.filter(cd => !store.d.cards[cd.id]).length;
    body.innerHTML = `<div class="card recap"><div class="row">${mascot('celebrate')}<div class="grow"><p class="eyebrow">Recap</p><h2>What you can do now</h2><ul class="ticks">${R.points.map(p => `<li>${md(p)}</li>`).join('')}</ul><div class="cheat"><b>Keep this line:</b> ${md(R.cheat)}</div></div></div>
      <h3>${R.cards.length} review cards ${st.status === 'done' ? 'in your deck' : 'go into your deck'}</h3><div class="cardgrid">${R.cards.map(cd => `<div class="mcard"><span class="ctype">${esc(cd.type)}</span><p>${md(cd.front)}</p><details><summary>Answer</summary><p>${md(cd.back)}</p></details></div>`).join('')}</div>
      <div class="row between">${prev ? `<a class="btn ghost" href="#/lesson/${c.id}/${prev.id}">← ${prev.num}</a>` : '<span></span>'}<div class="row"><button class="btn big" id="fin">${st.status === 'done' ? 'Lesson complete ✓' : 'Finish lesson'}</button>${next ? `<a class="btn ghost" href="#/lesson/${c.id}/${next.id}" id="nextL">Next: ${next.num} →</a>` : ''}</div></div></div>`;
    $('#fin').onclick = () => {
      const first = st.status !== 'done';
      for (const cd of R.cards) if (!store.d.cards[cd.id]) store.d.cards[cd.id] = { ...cd, lesson: id, course: c.id, isNew: true, added: dayNum(), due: dayNum() };
      if (first) { st.status = 'done'; st.doneAt = new Date().toISOString(); store.act('lessons'); confetti(); }
      store.save();
      const modEnd = meta.module.lessons.every(l => store.d.lessons[l.id] && store.d.lessons[l.id].status === 'done');
      const m = modal(`<div class="done-modal">${mascot('celebrate')}<h2>${esc(line('done'))}</h2><p>${fresh} new card${fresh === 1 ? '' : 's'} added to your review deck.${modEnd ? ` That’s every lesson in Module ${meta.module.n}. The module test is open.` : ''}</p><div class="row center">${modEnd ? `<a class="btn" href="#/mtest/${c.id}/${meta.module.n}">Take the Module ${meta.module.n} test</a>` : next ? `<a class="btn" href="${canOpen(c, next) ? `#/lesson/${c.id}/${next.id}` : '#/login'}">Next lesson: ${next.num}</a>` : ''}<a class="btn ghost" href="#/home">Back to home</a></div></div>`);
      $$('a', m.el).forEach(a => a.addEventListener('click', () => m.close()));
    };
  }
  await render();
});
function rc(a) { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2] - 1, c - 1]; }
export function resultsHTML(res) { return `<div class="score ${res.pct === 100 ? 'ok' : ''}"><b>${res.pct}%</b> · ${res.score} of ${res.total} checks passed</div><ul class="checks">${res.results.map(r => `<li class="${r.ok ? 'ok' : 'no'}"><span>${r.ok ? '✓' : '✗'}</span><div><b>${esc(r.label)}</b>${r.ok ? '' : `<p>${esc(r.msg)}</p>`}</div></li>`).join('')}</ul>`; }
function bumpSkill(id) { const s = store.lesson(id); s.firstTries = (s.firstTries || 0) + 1; }
async function resources(c, L) {
  const m = modal(`<h2>Downloads &amp; screenshots · ${esc(L.num)}</h2><div class="dl-list">${(L.downloads || []).map(d => `<button class="dl" data-file="${d.file}" data-name="${esc(d.name)}"><span class="ext ${d.file.split('.').pop()}">${d.file.split('.').pop().toUpperCase()}</span><span><b>${esc(d.label)}</b><small>${esc(d.desc || '')}</small></span></button>`).join('')}</div>
   <h3>How it should look</h3><div class="gallery">${(L.shots || []).map(s => `<figure class="shot"><img data-asset="${s.file}" alt="${esc(s.caption)}"><figcaption>${md(s.caption)}</figcaption></figure>`).join('') || '<p class="muted">No screenshots for this lesson.</p>'}</div>`, { wide: true });
  $$('[data-file]', m.el).forEach(b => b.onclick = async () => { try { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); } catch (e) { toast('Sign in to download this file.'); } });
  await hydrate(m.el, c.id);
}
const wait = ms => new Promise(r => setTimeout(r, ms));
