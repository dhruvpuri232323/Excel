// The lesson player: lecture → guided demo → lab → quiz → recap. Every part and slide has its own address,
// so Back, bookmarks and shared links work. One-time tips point at things the first time they appear.
import { register, getCourse, go, lockedPanel, keys, setFocus, setActiveCourse } from '../app.js';
import { store, today } from '../store.js';
import { session, lesson as loadLesson, lessonMeta, assetURL } from '../content.js';
import { esc, md, $, $$, mascot, line, toast, confetti, modal, bar, shuffle, download, crumbs, qhelp } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask, applySolution, describeCheck } from '../checker.js';
import { readXlsx, grade } from '../xlsxcheck.js';
import { dayNum } from '../srs.js';
import { lessonList, PARTS, PART_NAMES, partMinutes, moduleInfo, lessonLabel, lessonNo, fmtMin, canOpenLesson } from '../access.js';
import { glossify, askTeacher } from '../help.js';
import { tip, clearTip } from '../tour.js';
import { t as tr } from '../i18n.js';
import { toggleSaved, isSaved } from './extras.js';
import { diagram, DIAGRAMS } from '../visuals.js';

let grids = [];
// one listener for the whole app: a click outside an open menu closes it
document.addEventListener('click', e => { if (!e.target.closest('.menu')) document.querySelectorAll('.menu-pop').forEach(p => { p.hidden = true; const b = p.parentElement.querySelector('[aria-expanded]'); if (b) b.setAttribute('aria-expanded', 'false'); }); });
const killGrids = () => { grids.forEach(g => g.destroy()); grids = []; };
export async function visual(v) {
  if (!v) return '';
  if (v.type === 'shot') return `<figure class="shot"><button class="zoomable" type="button" aria-label="Enlarge screenshot"><img data-asset="${v.file}" alt="${esc(v.alt || v.caption || '')}" data-callouts='${esc(JSON.stringify(v.callouts || []))}' data-caption="${esc(v.caption || '')}"><span class="zoomhint" aria-hidden="true">⤢ Enlarge</span></button><figcaption>${md(v.caption || '')}</figcaption></figure>`;
  if (v.type === 'grid') return `<div class="mini-grid" data-grid='${esc(JSON.stringify(v))}'></div>${v.caption ? `<p class="cap">${md(v.caption)}</p>` : ''}`;
  if (v.type === 'formula') return `<div class="formula-anat"><div class="ftext">${v.parts.map((p, i) => `<span class="fp c${i % 6}">${esc(p[0])}</span>`).join('')}</div><ul>${v.parts.filter(p => p[1]).map(p => `<li><span class="fp c${v.parts.indexOf(p) % 6}">${esc(p[0])}</span> ${md(p[1])}</li>`).join('')}</ul></div>`;
  if (v.type === 'keys') return `<table class="keys vkeys"><tbody>${v.items.map(([k, w]) => `<tr><td>${md('[[' + k + ']]')}</td><td>${md(w)}</td></tr>`).join('')}</tbody></table>`;
  if (v.type === 'table') return `<table class="tbl"><thead><tr>${v.head.map(h => `<th>${md(h)}</th>`).join('')}</tr></thead><tbody>${v.rows.map(r => `<tr>${r.map(c => `<td>${md(String(c))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  if (v.type === 'tip') return `<div class="tip"><span class="tip-i" aria-hidden="true">💡</span><p>${md(v.text)}</p></div>`;
  if (v.type === 'compare') return `<div class="compare"><div class="bad"><b>✗ ${md(v.bad.title)}</b><p>${md(v.bad.text)}</p></div><div class="good"><b>✓ ${md(v.good.title)}</b><p>${md(v.good.text)}</p></div></div>`;
  if (v.type === 'steps') return `<ol class="steps vsteps">${v.items.map((s, i) => `<li><span class="stepn">${i + 1}</span><div>${md(s)}</div></li>`).join('')}</ol>`;
  if (DIAGRAMS.includes(v.type)) return diagram(v);
  return '';
}
// screenshots open full size, with zoom and the numbered callouts listed beside them (F2)
export function viewer(src, alt, callouts = [], caption = '') {
  let z = 1; const m = modal(`<div class="viewer"><div class="vtools row"><button class="btn ghost sm" data-z="-1" type="button">−</button><button class="btn ghost sm" data-z="0" type="button">Fit</button><button class="btn ghost sm" data-z="1" type="button">+</button><span class="muted small">${esc(caption)}</span></div>
    <div class="vbody"><div class="vpane"><img src="${src}" alt="${esc(alt)}" class="vimg"></div>${callouts.length ? `<ol class="vcall">${callouts.map(c => `<li><span class="stepn">${c.n}</span> ${esc(c.text)}</li>`).join('')}</ol>` : ''}</div></div>`, { wide: true });
  const img = $('.vimg', m.el), pane = $('.vpane', m.el); const set = () => { img.style.width = z === 1 ? '100%' : (z * 100) + '%'; img.style.maxWidth = 'none'; };
  $$('[data-z]', m.el).forEach(b => b.onclick = () => { const d = +b.dataset.z; z = d === 0 ? 1 : Math.min(3, Math.max(1, z + d * 0.5)); set(); });
  let drag = null; pane.onmousedown = e => { drag = [e.clientX, e.clientY, pane.scrollLeft, pane.scrollTop]; e.preventDefault(); }; addEventListener('mouseup', () => { drag = null; }); pane.onmousemove = e => { if (drag) { pane.scrollLeft = drag[2] - (e.clientX - drag[0]); pane.scrollTop = drag[3] - (e.clientY - drag[1]); } };
}
export async function hydrate(root, courseId) {
  for (const img of $$('img[data-asset]', root)) {
    try { img.src = await assetURL(courseId, img.dataset.asset); } catch (e) { img.alt = 'Screenshot unavailable'; }
    const btn = img.closest('.zoomable'); const open = () => viewer(img.src, img.alt, JSON.parse(img.dataset.callouts || '[]'), img.dataset.caption || ''); if (btn) btn.onclick = open; else if (!img.closest('.hotspot')) { img.onclick = open; img.style.cursor = 'zoom-in'; }
  }
  for (const d of $$('.mini-grid[data-grid]', root)) { const v = JSON.parse(d.dataset.grid); const g = new Grid(d, v.sheet, { readOnly: true, toolbar: false, tabs: (v.sheet.sheets || []).length > 1, status: false, height: v.height || 170, select: v.select }); if (v.showFormulas) { g.showFormulas = true; g.render(); } grids.push(g); }
}

register('lesson', async (main, courseId, id, partName, pos) => {
  killGrids(); clearTip();
  const c = await getCourse(courseId); const meta = lessonMeta(c, id); if (!meta) throw new Error('No such lesson');
  if (!canOpenLesson(c, meta)) { const i = moduleInfo(c, meta.module); main.innerHTML = i.reason === 'login' ? lockedPanel() : lockedPanel('test', `<h2>Module ${meta.module.n} isn’t open yet</h2><p>Pass the Module ${i.prev.n} test to open it, or test out of Module ${meta.module.n} if you already know it.</p><div class="row"><a class="btn" href="#/module/${c.id}/${i.prev.n}">Go to Module ${i.prev.n}</a><a class="btn ghost" href="#/mtest/${c.id}/${meta.module.n}">Test out of Module ${meta.module.n}</a></div>`); return; }
  setActiveCourse(c.id);
  const L = await loadLesson(c, id); const st = store.lesson(id); if (st.status === 'new') { st.status = 'started'; st.startedAt = new Date().toISOString(); store.save(); }
  const reach = () => st.status === 'done' ? 4 : Math.max(st.maxPart || 0, st.part || 0);
  // the address decides the part; a part not reached yet sends you to the furthest one you have reached
  let pi = PARTS.indexOf(partName); if (pi < 0) { location.replace(`#/lesson/${c.id}/${id}/${PARTS[st.part || 0]}`); return; }
  if (pi > reach()) { toast(`Finish the ${PART_NAMES[reach()].toLowerCase()} first: the ${PART_NAMES[pi].toLowerCase()} opens after it.`); location.replace(`#/lesson/${c.id}/${id}/${PARTS[reach()]}`); return; }
  st.part = pi; if (pi === 0 && pos && pos !== 'notes') { const [s, k] = pos.split('-').map(n => (+n || 1) - 1); if (L.lecture[s]) { st.seg = s; st.slide = Math.min(k, L.lecture[s].slides.length); } }
  store.save();
  const all = lessonList(c); const idx = all.findIndex(l => l.id === id); const next = all[idx + 1], prev = all[idx - 1];
  const partHref = i => `#/lesson/${c.id}/${id}/${PARTS[i]}`;
  const upNext = i => i < 4 ? `${tr(PART_NAMES[i + 1])} · about ${partMinutes(meta, i + 1)} min` : '';
  const advance = () => { const n = Math.min(4, st.part + 1); st.maxPart = Math.max(st.maxPart || 0, n); store.save(); go(partHref(n)); };
  const saved = isSaved('lesson:' + id);
  main.innerHTML = `<div class="lesson v3">
    <header class="lesson-head">${crumbs([[c.title, '#/course/' + c.id], [`Module ${meta.module.n}: ${meta.module.title}`, `#/module/${c.id}/${meta.module.n}`], [`Lesson ${lessonNo(meta)}`]])}
      <div class="lh-row"><div class="lh-title"><h1>${esc(L.title)}</h1><span class="lh-meta">${esc(lessonLabel(meta))} · about ${meta.minutes} min</span></div>
        <div class="lh-tools"><button class="iconbtn" id="outlineBtn" type="button" title="Every slide and task in this lesson">☰<span>${tr('Outline')}</span></button><button class="iconbtn" id="focusBtn" type="button" title="Hide everything else (F). Esc to leave.">⤢<span>${tr('Full screen')}</span></button>
          <div class="menu"><button class="iconbtn" id="moreBtn" type="button" aria-haspopup="true" aria-expanded="false">⋯<span>More</span></button><div class="menu-pop" id="morePop" hidden>
            <a href="#/lesson/${c.id}/${id}/lecture/notes">📖 Read the lecture as notes</a><button type="button" id="res">⬇ ${tr('Downloads & screenshots')}</button><button type="button" id="notesBtn">✐ ${tr('Notes')}${store.d.notes[id] ? ' ·' : ''}</button><button type="button" id="bm" aria-pressed="${saved}">${saved ? '★ Saved' : '☆ Save lesson'}</button><button type="button" data-ask="1" data-where="${esc(lessonLabel(meta) + ': ' + L.title + ', ' + PART_NAMES[pi])}">✉ ${tr('Ask your teacher')}</button></div></div></div></div>
      <nav class="parts stepper" aria-label="Lesson parts">${PART_NAMES.map((p, i) => { const can = i <= reach(); const done = i < reach() || st.status === 'done'; return `<a class="part ${i === pi ? 'on' : ''} ${done ? 'done' : ''} ${can ? '' : 'off'}" href="${can ? partHref(i) : '#'}" ${can ? '' : `aria-disabled="true" title="Finish the ${PART_NAMES[i - 1].toLowerCase()} to open the ${p.toLowerCase()}"`} ${i === pi ? 'aria-current="step"' : ''} data-p="${i}"><span class="pn">${done && i !== pi ? '✓' : i + 1}</span><span class="pl">${tr(p)}<small>${partMinutes(meta, i)} min</small></span></a>`; }).join('<i class="pline" aria-hidden="true"></i>')}${qhelp('parts')}</nav></header>
    <div id="partBody"></div></div>`;
  $$('.part.off').forEach(a => a.onclick = e => { e.preventDefault(); toast(a.title); });
  const pop = $('#morePop'), mb = $('#moreBtn'); const closeMenu = () => { pop.hidden = true; mb.setAttribute('aria-expanded', 'false'); };
  mb.onclick = e => { e.stopPropagation(); pop.hidden = !pop.hidden; mb.setAttribute('aria-expanded', String(!pop.hidden)); };  $$('a, button', pop).forEach(x => x.addEventListener('click', () => setTimeout(closeMenu, 0)));
  $('#bm').onclick = () => { toggleSaved({ type: 'lesson', key: 'lesson:' + id, course: c.id, lesson: id, title: L.title }); const on = isSaved('lesson:' + id); $('#bm').textContent = on ? '★ Saved' : '☆ Save lesson'; toast(on ? 'Saved. Find it in My stuff › Saved.' : 'Removed from Saved'); };
  $('#notesBtn').onclick = () => { const m = modal(`<h2>Your notes · ${esc(L.title)}</h2><textarea id="nt" rows="12" class="notes" placeholder="Write what you want to remember. Notes stay on this laptop, go into your backup, and all appear in My stuff › My notes.">${esc(store.d.notes[id] || '')}</textarea><div class="row end"><a class="btn ghost" href="#/notes">All my notes</a><button class="btn" id="ns" type="button">${tr('Save')}</button></div>`); $('#ns', m.el).onclick = () => { store.d.notes[id] = $('#nt', m.el).value; store.d.notesAt = store.d.notesAt || {}; store.d.notesAt[id] = new Date().toISOString(); store.save(); m.close(); toast('Notes saved'); }; };
  $('#res').onclick = () => resources(c, L);
  const focusToggle = () => { const on = !document.body.classList.contains('focus'); setFocus(on, on); };
  $('#focusBtn').onclick = focusToggle;
  $('#outlineBtn').onclick = () => outline(c, L, st, reach, id);
  keys.handler = e => { if (e.key === 'f' || e.key === 'F') { e.preventDefault(); focusToggle(); } else if ((e.key === 'n' || e.key === 'N') && st.part < reach()) { e.preventDefault(); go(partHref(st.part + 1)); } else if (keys.part) keys.part(e); };
  keys.part = null;
  const body = $('#partBody');
  if (pi === 0 && pos === 'notes') await notes(body); else if (pi === 0) await lecture(body); else if (pi === 1) await demo(body); else if (pi === 2) await lab(body); else if (pi === 3) await quiz(body); else await recap(body);
  setTimeout(() => tip('parts', $('.parts')), 400);

  // ---------- 1. lecture ----------
  async function lecture(body) {
    const segs = L.lecture; st.seg = Math.min(st.seg || 0, segs.length - 1); st.slide = st.slide || 0;
    const seg = segs[st.seg]; const atCheck = st.slide >= seg.slides.length; const slide = seg.slides[Math.min(st.slide, seg.slides.length - 1)];
    const skey = `slide:${id}:${st.seg}-${st.slide}`;
    history.replaceState(null, '', `#/lesson/${c.id}/${id}/lecture/${st.seg + 1}-${Math.min(st.slide, seg.slides.length) + 1}`);
    const total = segs.reduce((n, sg) => n + sg.slides.length + 1, 0); const at = segs.slice(0, st.seg).reduce((n, sg) => n + sg.slides.length + 1, 0) + Math.min(st.slide, seg.slides.length) + 1;
    const steps = segs.map((sg, si) => `<div class="lseg ${si === st.seg ? 'on' : si < st.seg ? 'done' : ''}"><span class="lseg-t">${si < st.seg ? '✓ ' : ''}${esc(sg.title)}</span><span class="lseg-dots">${sg.slides.map((x, k) => `<i class="${si < st.seg || (si === st.seg && k < st.slide) ? 'done' : si === st.seg && k === st.slide ? 'on' : ''}"></i>`).join('')}<i class="q ${si < st.seg ? 'done' : si === st.seg && atCheck ? 'on' : ''}" title="Quick check"></i></span></div>`).join('');
    body.innerHTML = `<div class="card lecture v3"><div class="lec-top"><div class="lec-steps" aria-label="Lecture progress">${steps}</div><div class="lec-right"><span class="lec-count">${atCheck ? 'Quick check' : `Slide ${at} of ${total}`}</span>${atCheck ? '' : `<button class="mini save-slide" type="button" aria-pressed="${isSaved(skey)}" title="Save this slide">${isSaved(skey) ? '★' : '☆'}</button>`}</div></div>
      ${atCheck ? checkQuestion(seg.check) : await slideHTML(slide)}
      <div class="lec-nav"><button class="btn ghost" id="back" type="button" ${st.seg === 0 && st.slide === 0 ? 'disabled' : ''}>← ${tr('Back')}</button><div class="lec-prog" aria-hidden="true"><i style="width:${Math.round(at / total * 100)}%"></i></div>${atCheck ? '<span class="muted small">Pick an answer</span>' : `<button class="btn" id="fwd" type="button">${st.slide === seg.slides.length - 1 ? 'Quick check →' : tr('Next') + ' →'}</button>`}</div></div>
      <p class="readrow"><a href="#/lesson/${c.id}/${id}/lecture/notes">📖 Prefer reading? Open the whole lecture as one page of notes</a></p>`;
    await hydrate(body, c.id); glossify($('.vs-side', body) || $('.slide-body', body) || $('.checkq', body));
    const save = () => store.save();
    const back = () => { if (st.slide > 0) st.slide--; else if (st.seg > 0) { st.seg--; st.slide = segs[st.seg].slides.length; } else return; save(); lecture(body); };
    const fwd = () => { if (atCheck) return; st.slide++; save(); lecture(body); };
    $('#back').onclick = back; if (!atCheck) { $('#fwd').onclick = fwd; setTimeout(() => tip('next', $('#fwd')), 900); }
    const ss = $('.save-slide', body); if (ss) ss.onclick = () => { toggleSaved({ type: 'slide', key: skey, course: c.id, lesson: id, pos: `${st.seg + 1}-${st.slide + 1}`, title: String(slide.h).replace(/\*\*|`/g, '') }); lecture(body); };
    keys.part = e => { if (e.key === 'ArrowRight') { e.preventDefault(); fwd(); } if (e.key === 'ArrowLeft') { e.preventDefault(); back(); } };
    if (atCheck) wireCheck(body, seg.check, () => { if (st.seg < segs.length - 1) { st.seg++; st.slide = 0; save(); lecture(body); scrollTo(0, 0); } else { st.seg = 0; st.slide = 0; advance(); } }, 'Next segment →', st.seg === segs.length - 1 ? `On to the demo · about ${partMinutes(meta, 1)} min →` : null);
  }
  // the lecture as one page of notes: what you'll learn, the plan, then each part in order with its pictures,
  // a check question with the answer tucked away, and the summary
  async function notes(body) {
    const segs = L.lecture; const plain = t => String(t).replace(/\*\*|`/g, '').replace(/[.:]$/, '');
    const low = t => { const x = plain(t); return /^[A-Z][a-z]/.test(x) ? x[0].toLowerCase() + x.slice(1) : x; };
    const list = a => a.length < 2 ? a.join('') : a.slice(0, -1).join('; ') + '; and ' + a.at(-1);
    const words = segs.reduce((n, sg) => n + sg.slides.reduce((m, sl) => m + (sl.p || []).join(' ').split(/\s+/).length, 0), 0) * (segs.some(sg => sg.slides.some(sl => sl.say)) ? 1.4 : 1);
    const num = n => ['', 'one idea', 'two ideas', 'three ideas', 'four ideas', 'five ideas'][n] || n + ' ideas'; const opener = (si, n) => si === 0 ? `First, ${num(n)}, in this order:` : si === segs.length - 1 ? `Finally, ${num(n)} to finish:` : `Then ${num(n)}:`;
    const parts = [];
    for (const [si, sg] of segs.entries()) {
      const steps = [];
      for (const [k, sl] of sg.slides.entries()) steps.push(`<div class="r-step" id="n${si + 1}-${k + 1}"><div class="r-head"><span class="r-num">${si + 1}.${k + 1}</span><h4>${md(sl.h)}</h4></div>${sl.say || (sl.points || []).length ? `${sl.say ? `<p class="r-say">${md(sl.say)}</p>` : ''}<ul class="vs-points">${(sl.points || []).map(p => `<li>${md(p)}</li>`).join('')}</ul>` : (sl.p || []).map(p => `<p>${md(p)}</p>`).join('')}${sl.visual ? `<div class="r-fig"><span class="r-figlab">Figure ${si + 1}.${k + 1}</span>${await visual(sl.visual)}</div>` : ''}${sl.real ? `<div class="vs-real"><span class="vs-real-i" aria-hidden="true">${esc(sl.real.icon || '🏢')}</span><div><span class="vs-tag">In real life</span><b>${md(sl.real.who)}</b><p>${md(sl.real.text)}</p></div></div>` : ''}${(sl.terms || []).length ? `<div class="vs-terms inline"><span class="vs-tag">📖 Words to know</span>${sl.terms.map(([t, d]) => `<div class="vs-term"><b>${md(t)}</b><span>${md(d)}</span></div>`).join('')}</div>` : ''}</div>`);
      const q = sg.check;
      parts.push(`<section class="r-part"><p class="eyebrow">Part ${si + 1} of ${segs.length}</p><h3>${esc(sg.title)}</h3><p class="r-lead">${opener(si, sg.slides.length)}</p><ol class="r-glance">${sg.slides.map((sl, k) => `<li><a href="#n${si + 1}-${k + 1}" data-jump="n${si + 1}-${k + 1}">${esc(plain(sl.h))}</a></li>`).join('')}</ol>${steps.join('')}
        <aside class="r-check"><b>Check yourself</b><p>${md(q.q)}</p><ol type="A">${q.options.map(o => `<li>${md(o)}</li>`).join('')}</ol><details><summary>Show the answer</summary><p><b>${String.fromCharCode(65 + q.answer)}. ${md(q.options[q.answer])}</b> ${md(q.explain || '')}</p></details></aside></section>`);
    }
    body.innerHTML = `<article class="card reader"><div class="row between noprint"><a class="btn ghost sm" href="#/lesson/${c.id}/${id}/lecture">← Back to the slides</a><button class="btn ghost sm" type="button" id="prn">🖨 Print or save as PDF</button></div>
      <header class="r-top"><p class="eyebrow">Lecture notes · ${esc(lessonLabel(meta))} · about ${Math.max(3, Math.round(words / 180))} minutes to read</p><h2>${esc(L.title)}</h2><p class="lead">${md(L.summary || '')}</p></header>
      <section class="r-intro"><div><h3>By the end, you can</h3><ul class="ticks">${(L.objectives || []).map(o => `<li>${md(o)}</li>`).join('')}</ul></div>
        <div><h3>How this lecture runs</h3><ol class="r-plan">${segs.map((sg, si) => `<li><a href="#n${si + 1}-1" data-jump="n${si + 1}-1"><b>${esc(sg.title)}</b></a><small>${esc(sg.slides.map(sl => plain(sl.h)).join(' · '))}</small></li>`).join('')}<li><b>Summary</b><small>What to remember</small></li></ol></div></section>
      ${parts.join('')}
      <section class="r-part r-sum"><p class="eyebrow">Summary</p><h3>What you can do now</h3><ul class="ticks">${L.recap.points.map(p => `<li>${md(p)}</li>`).join('')}</ul><div class="cheat"><b>Keep this line:</b> ${md(L.recap.cheat)}</div>
        <div class="row noprint"><a class="btn" href="#/lesson/${c.id}/${id}/lecture">Back to the slides</a>${reach() >= 1 ? `<a class="btn ghost" href="#/lesson/${c.id}/${id}/demo">On to the demo →</a>` : '<span class="muted small">Finish the slides and their checks to open the demo.</span>'}</div></section></article>`;
    await hydrate(body, c.id); glossify(body, 10);
    $('#prn').onclick = () => print();
    $$('[data-jump]', body).forEach(a => a.onclick = e => { e.preventDefault(); const el = document.getElementById(a.dataset.jump); el && el.scrollIntoView({ behavior: 'smooth', block: 'start' }); });
  }
  // a visual-first slide: the idea in one line, the diagram, two or three short points, a real-life example, words to know
  async function slideHTML(sl) {
    const vis = await visual(sl.visual); const wide = sl.visual && (sl.visual.wide || ['flow', 'beforeafter', 'icons', 'bignum', 'shot', 'table', 'compare', 'layers', 'analogy', 'anatomy'].includes(sl.visual.type));
    if (!sl.say && !(sl.points || []).length) return `<div class="vslide legacy"><h2 class="vs-h">${md(sl.h)}</h2><div class="vs-legacy slide-body">${(sl.p || []).map(p => `<p>${md(p)}</p>`).join('')}</div>${vis ? `<div class="vs-vis">${vis}</div>` : ''}</div>`;
    const pts = (sl.points || []).length ? `<ul class="vs-points">${sl.points.map(p => `<li>${md(p)}</li>`).join('')}</ul>` : '';
    const terms = (sl.terms || []).length ? `<div class="vs-terms"><span class="vs-tag">📖 Words to know</span>${sl.terms.map(([t, d]) => `<div class="vs-term"><b>${md(t)}</b><span>${md(d)}</span></div>`).join('')}</div>` : '';
    const real = sl.real ? `<div class="vs-real"><span class="vs-real-i" aria-hidden="true">${esc(sl.real.icon || '🏢')}</span><div><span class="vs-tag">In real life</span><b>${md(sl.real.who)}</b><p>${md(sl.real.text)}</p></div></div>` : '';
    // narrow diagrams sit beside the points; wide ones span the slide with the points underneath
    const bodyHTML = !vis ? `<div class="vs-body novis"><div class="vs-side">${pts}${real}${terms}</div></div>`
      : wide ? `<div class="vs-body wide"><div class="vs-vis">${vis}</div><div class="vs-side"><div>${pts}</div><div>${real}${terms}</div></div></div>`
      : `<div class="vs-body"><div class="vs-vis">${vis}${terms}</div><div class="vs-side">${pts}${real}</div></div>`;
    return `<div class="vslide"><h2 class="vs-h">${md(sl.h)}</h2>${sl.say ? `<p class="vs-say">${md(sl.say)}</p>` : ''}${bodyHTML}</div>`;
  }
  function checkQuestion(q) { return `<div class="checkq v3"><div><p class="eyebrow">✓ Quick check · one question to lock it in</p><h2>${md(q.q)}</h2><div class="opts">${q.options.map((o, i) => `<button class="opt" data-i="${i}" type="button">${md(o)}</button>`).join('')}</div><div class="feedback" role="status"></div></div></div>`; }
  function wireCheck(body, q, onNext, nextLabel, lastLabel) {
    let misses = 0; const fb = $('.feedback', body);
    const right = (b, shown) => { b.classList.add('right'); fb.innerHTML = `<p class="ok">✓ ${shown ? 'The answer is: ' + md(q.options[q.answer]) + '. ' : ''}${md(q.explain)}</p><button class="btn" id="cont" type="button">${lastLabel || nextLabel}</button>`; $('#cont', body).onclick = onNext; $('#cont', body).focus(); $$('.opt', body).forEach(x => x.disabled = true); };
    $$('.opt', body).forEach(b => b.onclick = () => {
      const i = +b.dataset.i; $$('.opt', body).forEach(x => x.classList.remove('right', 'wrong'));
      if (i === q.answer) return right(b);
      misses++; b.classList.add('wrong');
      fb.innerHTML = `<p class="no">${mascot('oops', 'xs')} ${md((q.why && q.why[i]) || 'Not quite. Think about what the slides said, and try another answer.')}</p>${misses >= 2 ? `<button class="btn ghost sm" id="showAns" type="button">${tr('Show me the answer')}</button>` : ''}`;
      const sa = $('#showAns', body); if (sa) sa.onclick = () => { store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: 'check-' + st.seg, kind: 'check', prompt: q.q, yours: q.options[i], right: q.options[q.answer], why: q.explain }); store.save(); right($$('.opt', body)[q.answer], true); };
    });
  }
  // ---------- 2. demo ----------
  async function demo(body) {
    const D = L.demo;
    if (D.frames) { // a walkthrough of screenshots, for skills that happen in Excel's menus
      st.frame = Math.min(st.frame || 0, D.frames.length - 1); const f = D.frames[st.frame];
      body.innerHTML = `<div class="card demo"><p class="eyebrow">Guided demo · step ${st.frame + 1} of ${D.frames.length}</p><p class="lead">${md(D.intro)}</p><figure class="shot big"><button class="zoomable" type="button" aria-label="Enlarge screenshot"><img data-asset="${f.img}" alt="${esc(f.say)}" data-callouts='${esc(JSON.stringify(f.callouts || []))}' data-caption="${esc(f.say)}"><span class="zoomhint" aria-hidden="true">⤢ Enlarge</span></button><figcaption><span class="stepn">${st.frame + 1}</span> ${md(f.say)}</figcaption></figure>
       <div class="row between"><button class="btn ghost" id="b" type="button" ${st.frame ? '' : 'disabled'}>← ${tr('Back')}</button>${st.frame < D.frames.length - 1 ? `<button class="btn" id="n" type="button">Next step →</button>` : `<button class="btn" id="n" type="button">Now you try it · ${upNext(1)} →</button>`}</div></div>`;
      await hydrate(body, c.id);
      const b = () => { if (st.frame > 0) { st.frame--; store.save(); demo(body); } }, n = () => { if (st.frame < D.frames.length - 1) { st.frame++; store.save(); demo(body); } else { st.frame = 0; advance(); } };
      $('#b').onclick = b; $('#n').onclick = n; keys.part = e => { if (e.key === 'ArrowRight') { e.preventDefault(); n(); } if (e.key === 'ArrowLeft') { e.preventDefault(); b(); } };
      return;
    }
    body.innerHTML = `<div class="card demo"><p class="eyebrow">Guided demo</p><p class="lead">${md(D.intro)}</p><div class="say" id="cap">${mascot('teaching', 'sm')}<div class="bubble">Press <b>Play</b> and watch the sheet, or go <b>Step by step</b>. You’ll do it yourself in the lab next.</div></div>
      <div id="dg"></div><div class="row between"><div class="row"><button class="btn" id="play" type="button">▶ Play demo</button><button class="btn ghost" id="step" type="button">Step by step</button></div><button class="btn" id="n" type="button" ${st.demoSeen ? '' : 'disabled'} title="${st.demoSeen ? '' : 'Watch the demo once to continue'}">Now you try it · ${upNext(1)} →</button></div>${st.demoSeen ? '' : '<p class="muted small right">Watch it once (Play or Step by step) to continue.</p>'}</div>`;
    const g = new Grid($('#dg'), new Workbook(D.sheet), { readOnly: true, toolbar: false, height: D.height || 260 }); grids.push(g);
    let i = 0, busy = false;
    const cap = t => { const b = $('#cap .bubble'); if (b) b.innerHTML = md(t); };
    const runStep = async s => { g.o.readOnly = false; try { await runStep0(s); } finally { g.o.readOnly = true; } };
    const runStep0 = async s => {
      if (s.say) cap(s.say);
      if (s.do === 'select') await g.demoSelect(s.at, 900);
      else if (s.do === 'type') await g.demoType(s.at, s.text);
      else if (s.do === 'fill') await g.demoFill(s.from, s.to);
      else if (s.do === 'style') await g.demoStyle(s.at, s.style);
      else if (s.do === 'formulas') { g.showFormulas = !g.showFormulas; g.render(); await wait(900); }
      else if (s.do === 'explain') { g.select(s.at); await wait(500); g.explain(); await wait(1800); }
      else if (s.do === 'sort') { g.select(s.at); await wait(400); g.wb.sortRegion(0, g.wb.currentRegion(0, g.sel.ar, g.sel.ac), [{ col: g.sel.ac, desc: !!s.desc }], true); g.render(); await wait(700); }
      else if (s.do === 'width') { g.sheet.widths[s.col.charCodeAt(0) - 65] = s.px; g.render(); await wait(700); }
      else if (s.do === 'insertRow') { g.wb.insertRows(0, s.row - 1, 1); g.render(); await wait(700); }
      else if (s.do === 'deleteRow') { g.wb.deleteRows(0, s.row - 1, 1); g.render(); await wait(700); }
      else if (s.do === 'name') { g.wb.defineName(s.name, s.ref); g.select(s.ref.replace(/^.*!/, '').replace(/\$/g, '')); await wait(800); }
      else if (s.do === 'sheet') { g.setSheet(s.index); await wait(800); }
      else if (s.do === 'paste') { g.select(s.from); await wait(400); g.copy(); await wait(700); g.select(s.to); await wait(400); g.paste(s.mode || 'all'); await wait(800); }
      else if (s.do === 'clear') { g.select(s.at); await wait(400); g.clearSelection(); await wait(600); }
      else if (s.do === 'replace') { g.record(); const u = g.wb.usedRange(g.sheet); for (let r = u.r1; r <= u.r2; r++) for (let cc = u.c1; cc <= u.c2; cc++) { const tx = g.wb.input(g.si, r, cc); if (tx && (s.whole ? tx.toLowerCase() === s.find.toLowerCase() : tx.includes(s.find))) g.wb.setInput(g.si, r, cc, s.whole ? s.with : tx.split(s.find).join(s.with)); } g.render(); await wait(900); }
      else await wait(1400);
      if (s.say) await wait(Math.min(2600, 900 + s.say.length * 22));
    };
    const seen = () => { st.demoSeen = true; store.save(); const n = $('#n'); if (n) { n.disabled = false; n.title = ''; } const hint = $('.demo .right'); if (hint) hint.remove(); };
    $('#play').onclick = async () => { if (busy) return; busy = true; g.wb.load(D.sheet); g.setSel(0, 0); for (const s of D.steps) { if (!document.body.contains(g.host)) return; await runStep(s); } busy = false; seen(); cap(D.outro || 'That’s the whole move. Your turn.'); };
    $('#step').onclick = async () => { if (busy) return; if (i === 0) { g.wb.load(D.sheet); g.setSel(0, 0); } busy = true; await runStep(D.steps[i]); busy = false; if (!$('#step')) return; i = (i + 1) % D.steps.length; $('#step').textContent = i ? `Next step (${i + 1}/${D.steps.length})` : 'Step by step'; if (!i) seen(); };
    $('#n').onclick = () => advance();
  }
  // ---------- 3. lab ----------
  async function lab(body) {
    const B = L.lab; st.tasks = st.tasks || {};
    const gridTasks = B.tasks.filter(t => !['mc', 'hotspot', 'order'].includes(t.kind));
    body.innerHTML = `<div class="lab"><div class="card lab-intro"><p class="eyebrow">Lab · ${B.tasks.length} tasks · about ${partMinutes(meta, 2)} min</p><p>${md(B.intro)}</p>${idx === 0 ? '<p class="small okbox note-line">No Excel on this laptop? That’s fine: every lab runs here and needs nothing installed. For the Real Excel tasks later, the free Excel for the web works too. <a href="#/help/noexcel">How to get it</a></p>' : ''}${gridTasks.length ? `<p class="muted small">Work in the highlighted cells. Click a cell and type; press Enter to finish. Right-click for more options. Nothing you do here can break anything: <button class="linkbtn" id="resetLab" type="button">reset the sheet</button>.${B.inputs ? ' ' + qhelp('inputs') : ''}</p>` : ''}</div>
     <div class="lab-cols ${gridTasks.length ? '' : 'nogrid'}"><ol class="tasks">${B.tasks.map((t, i) => taskCard(t, i)).join('')}</ol>${gridTasks.length ? '<div class="lab-grid"><div id="lg"></div></div>' : ''}</div>
     ${B.real ? realCard(B.real) : ''}
     <div class="row between"><span class="muted" id="labStat"></span><button class="btn" id="n" type="button">On to the quiz · ${upNext(2)} →</button></div></div>`;
    let g = null; const keyLog = []; let usedMouse = false;
    if (gridTasks.length) {
      const savedSheet = st.labSheet; const wb = new Workbook(savedSheet || B.sheet);
      g = new Grid($('#lg'), wb, { height: B.height || 380, editable: B.editable, highlight: B.highlight || B.editable, allowSort: B.allowSort, allowStructure: B.allowStructure, styleAnywhere: B.styleAnywhere, noAddSheet: !B.allowAddSheet, onChange: () => { st.labSheet = wb.toSpec(); store.save(); }, onKey: k => keyLog.push(k) }); grids.push(g);
      $('#lg').addEventListener('mousedown', () => { usedMouse = true; }, true);
      $('#resetLab').onclick = () => { delete st.labSheet; store.save(); lab(body); };
      setTimeout(() => tip('grid', $('#lg')), 2500);
    }
    const answers = {};
    const refresh = () => { const done = B.tasks.filter(t => st.tasks[t.id] && st.tasks[t.id].done).length; $('#labStat').textContent = `${done} of ${B.tasks.length} tasks done`; $('#n').disabled = done < B.tasks.length; $('#n').title = done < B.tasks.length ? 'Finish every task to continue' : ''; };
    B.tasks.forEach((t, i) => {
      const card = $(`#task-${i}`); const ts = () => st.tasks[t.id] || (st.tasks[t.id] = { hints: 0, attempts: 0 });
      if (t.kind === 'mc') $$('.opt', card).forEach(b => b.onclick = () => { $$('.opt', card).forEach(x => x.classList.remove('sel')); b.classList.add('sel'); answers[t.id] = +b.dataset.i; });
      if (t.kind === 'hotspot') { const img = $('img', card); img.onclick = e => { const r = img.getBoundingClientRect(); const x = (e.clientX - r.left) / r.width * 100, y = (e.clientY - r.top) / r.height * 100; const z = t.zones.find(z => x >= z.x && x <= z.x + z.w && y >= z.y && y <= z.y + z.h); answers[t.id] = z ? z.label : 'empty space'; const m = $('.hotmark', card); m.style.left = x + '%'; m.style.top = y + '%'; m.hidden = false; }; }
      if (t.kind === 'order') { const ul = $('.order', card); let drag = null; const sync = () => { answers[t.id] = $$('li', ul).map(li => li.dataset.v); }; sync(); $$('li', ul).forEach(li => { li.draggable = true; li.ondragstart = () => { drag = li; li.classList.add('drag'); }; li.ondragend = () => { li.classList.remove('drag'); sync(); }; li.ondragover = e => { e.preventDefault(); if (drag && drag !== li) { const r = li.getBoundingClientRect(); ul.insertBefore(drag, e.clientY < r.top + r.height / 2 ? li : li.nextSibling); } }; $$('button', li).forEach(b => b.onclick = () => { const sib = b.dataset.d === 'up' ? li.previousElementSibling : li.nextElementSibling; if (sib) { b.dataset.d === 'up' ? ul.insertBefore(li, sib) : ul.insertBefore(sib, li); sync(); } }); }); }
      const showHints = s => { $('.hints', card).innerHTML = t.hints.slice(0, s.hints).map((h, k) => `<li><b>Hint ${k + 1}.</b> ${md(h)}</li>`).join(''); $('.hint', card).textContent = s.hints < t.hints.length ? `${tr('Hint')} ${s.hints + 1} of ${t.hints.length}` : 'No more hints'; if (s.hints >= t.hints.length) $('.sol', card).hidden = false; };
      $('.check', card).onclick = () => {
        const s = ts(); s.attempts++; const res = checkTask(t, { wb: g && g.wb, grid: g, inputs: B.inputs, answer: answers[t.id], keyLog, usedMouse });
        const fb = $('.tfb', card); const what = describeCheck(t, B.inputs); const whatHTML = what.length ? `<details class="what"><summary>What Sheetah checked</summary><ul>${what.map(w => `<li>${esc(w)}</li>`).join('')}</ul></details>` : '';
        if (res.ok) { s.done = true; if (s.first === undefined) s.first = s.hints === 0 && s.attempts === 1; if (s.sawSolution) { s.help = true; if (!s.logged) { s.logged = true; store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: t.id, kind: 'lab', prompt: t.prompt, yours: '(done with the solution)', right: t.solution || '', why: t.explain || 'Done with help: try it again from memory in your review.' }); } }
          fb.innerHTML = `<p class="ok">✓ ${s.help ? 'Done with help. It comes back in your daily review. ' : ''}${md(res.msg)}</p>${whatHTML}`; card.classList.add('done'); store.act('tasks'); if (s.first) bumpSkill(id); }
        else { fb.innerHTML = `<p class="no">${mascot('oops', 'xs')} ${md(res.msg)}</p>${whatHTML}${s.attempts >= 3 ? `<p class="small">Still stuck? <button class="linkbtn" type="button" data-ask="1" data-where="${esc(lessonLabel(meta) + ', lab task ' + (i + 1))}" data-what="${esc('I’m stuck on this task: ' + String(t.prompt).replace(/\*\*/g, '') + ' Sheetah says: ' + res.msg)}">Ask your teacher</button></p>` : ''}`; card.classList.add('shake'); setTimeout(() => card.classList.remove('shake'), 400);
          if (!s.logged) { s.logged = true; const cell = t.cell || (t.range ? t.range.split(':')[0] : null); store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: t.id, kind: 'lab', prompt: t.prompt, yours: cell && g ? (g.wb.input(0, ...rc(cell)) || '(blank)') : String(answers[t.id] ?? '(no answer)'), right: t.solution || (t.kind === 'mc' ? t.options[t.answer] : t.kind === 'hotspot' ? t.answer : t.kind === 'order' ? t.items.join(' → ') : ''), why: res.msg }); }
          if (store.d.settings.autoHint && s.hints < t.hints.length) { s.hints++; showHints(s); }
          if (s.attempts >= 3) $('.sol', card).hidden = false; }
        store.save(); refresh();
      };
      $('.hint', card).onclick = () => { const s = ts(); if (s.hints < t.hints.length) s.hints++; showHints(s); store.save(); };
      $('.sol', card).onclick = () => { const s = ts(); s.sawSolution = true; store.save(); const m = modal(`<h2>Worked solution</h2>${t.solution ? `<p><code class="big">${esc(t.solution)}</code></p>` : ''}<p>${md(t.explain || '')}</p><p class="muted small">This task will count as “done with help” and come back in your daily review.</p>${g && !['mc', 'hotspot', 'order'].includes(t.kind) ? '<div class="row end"><button class="btn ghost" id="put" type="button">Put it in the sheet for me</button><button class="btn" id="ok" type="button">I’ll do it myself</button></div>' : '<div class="row end"><button class="btn" id="ok" type="button">Got it</button></div>'}`); const p = $('#put', m.el); if (p) p.onclick = () => { applySolution(t, { wb: g.wb, grid: g }); st.labSheet = g.wb.toSpec(); store.save(); m.close(); }; $('#ok', m.el).onclick = m.close; };
      const sv = $('.save-task', card); if (sv) sv.onclick = () => { toggleSaved({ type: 'task', key: `task:${id}:${t.id}`, course: c.id, lesson: id, title: String(t.prompt).replace(/\*\*|`/g, '').slice(0, 120) }); sv.textContent = isSaved(`task:${id}:${t.id}`) ? '★' : '☆'; };
      const s0 = st.tasks[t.id]; if (s0 && s0.hints) showHints(s0);
      if (s0 && (s0.attempts >= 3 || s0.hints >= t.hints.length)) $('.sol', card).hidden = false;
      card.addEventListener('focusin', () => { if (g && t.focus) g.select(t.focus); });
      $('.goto', card) && ($('.goto', card).onclick = () => { if (g && t.focus) { g.select(t.focus); g.focus(); } });
    });
    if (B.real) wireReal(body, B.real);
    await hydrate(body, c.id); refresh(); glossify($('.lab-intro', body), 3);
    $('#n').onclick = () => advance();
    const firstOpen = B.tasks.findIndex(t => !(st.tasks[t.id] && st.tasks[t.id].done)); if (firstOpen > 0) $(`#task-${firstOpen}`).scrollIntoView({ block: 'center' });
    setTimeout(() => { tip('check', $(`#task-${Math.max(0, firstOpen)} .check`)); }, 900);
    setTimeout(() => { tip('hint', $(`#task-${Math.max(0, firstOpen)} .hint`)); }, 9000);
  }
  function taskCard(t, i) {
    const s = (st.tasks || {})[t.id] || {};
    let extra = '';
    if (t.kind === 'mc') extra = `<div class="opts">${t.options.map((o, k) => `<button class="opt" data-i="${k}" type="button">${md(o)}</button>`).join('')}</div>`;
    if (t.kind === 'hotspot') extra = `<div class="hotspot"><img data-asset="${t.img}" alt="${esc(t.alt || 'Excel window screenshot')}"><span class="hotmark" hidden></span></div>`;
    if (t.kind === 'order') extra = `<ul class="order">${shuffle(t.items).map(v => `<li data-v="${esc(v)}"><span>⋮⋮</span> ${md(v)} <button class="mini" data-d="up" aria-label="Move up" type="button">↑</button><button class="mini" data-d="down" aria-label="Move down" type="button">↓</button></li>`).join('')}</ul>`;
    return `<li class="card task ${s.done ? 'done' : ''}" id="task-${i}" tabindex="-1"><div class="tnum">${i + 1}</div><div class="tbody"><div class="row between"><p>${md(t.prompt)}</p><button class="mini save-task" type="button" title="Save this task">${isSaved(`task:${id}:${t.id}`) ? '★' : '☆'}</button></div>${extra}
      <div class="row"><button class="btn sm check" type="button">${tr('Check')}</button><button class="btn ghost sm hint" type="button">${s.hints ? (s.hints < t.hints.length ? `${tr('Hint')} ${s.hints + 1} of ${t.hints.length}` : 'No more hints') : `${tr('Hint')} 1 of ${t.hints.length}`}</button>${t.focus ? '<button class="btn ghost sm goto" type="button">Go to cell</button>' : ''}<button class="btn ghost sm sol" type="button" hidden>${tr('Show solution')}</button></div>
      <ul class="hints"></ul><div class="tfb" role="status">${s.done ? `<p class="ok">✓ ${s.help ? 'Done with help' : 'Done'}</p>` : ''}</div></div></li>`;
  }
  function realCard(R) {
    const r = (st.real || {}); return `<section class="card real"><div class="row between"><div><p class="eyebrow">Real Excel task${R.optional !== false ? ' · optional but recommended' : ''} ${qhelp('real')}</p><h2>${md(R.title)}</h2></div>${mascot('focus', 'sm')}</div>
     <ol class="steps"><li>Download the practice workbook and open it in Excel (desktop, or the free Excel for the web: <a href="#/help/noexcel">no Excel?</a>).</li>${R.steps.map(s => `<li>${md(s)}</li>`).join('')}<li>Save it as an Excel Workbook (.xlsx) and upload it here. Sheetah checks it at once.</li></ol>
     <div class="row"><button class="btn ghost" data-file="${R.file}" data-name="${esc(R.fileName || 'practice.xlsx')}" type="button">⬇ Practice workbook</button><label class="btn">⬆ Upload my file<input type="file" accept=".xlsx,.xlsm,.xls,.csv,.ods,.numbers" hidden class="up"></label>${R.solution ? `<button class="btn ghost solbtn" data-file="${R.solution}" data-name="${esc((R.fileName || 'practice.xlsx').replace('.xlsx', ' - solution.xlsx'))}" type="button" ${r.done || (r.tries || 0) >= 2 ? '' : 'hidden'}>⬇ Model answer</button>` : ''}<a class="small" href="#/help/upload">Problems uploading?</a></div>
     <div class="results">${r.last ? resultsHTML(r.last) : ''}</div></section>`;
  }
  function wireReal(body, R) {
    $$('[data-file]', body).forEach(b => b.onclick = async () => { try { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); } catch (e) { toast('Sign in to download this file.'); } });
    $('.up', body).onchange = async e => { const f = e.target.files[0]; if (!f) return; const box = $('.results', body); box.innerHTML = '<p>Checking your workbook…</p>'; const x = await readXlsx(f); e.target.value = '';
      if (x.error) { box.innerHTML = `<div class="score">${mascot('oops', 'xs')} <b>That file didn’t open.</b> ${esc(x.error)} <a href="#/help/upload">More help</a></div>`; return; }
      const res = grade(x, R.checks); st.real = st.real || {}; st.real.tries = (st.real.tries || 0) + 1; st.real.last = res; if (res.pct === 100) { st.real.done = true; confetti(); bumpSkill(id); } store.save(); box.innerHTML = resultsHTML(res); const sb = $('.solbtn', body); if (sb && (st.real.done || st.real.tries >= 2)) sb.hidden = false; };
  }
  // ---------- 4. quiz ----------
  async function quiz(body) {
    const Q = L.quiz; st.q = st.q || { i: 0, right: 0, answered: [] }; const q = Q[st.q.i];
    if (st.q.i >= Q.length) { st.quizAt = st.quizAt && st.quizAt > (st.qStart || '') ? st.quizAt : new Date().toISOString(); store.save();
      body.innerHTML = `<div class="card quiz-end">${mascot(st.q.right / Q.length >= .8 ? 'celebrate' : 'encourage')}<div><h2>You got ${st.q.right} of ${Q.length}</h2><p>${st.q.right === Q.length ? 'Perfect score.' : 'The ones you missed are in Mistakes and come back in your daily review.'}</p><div class="row"><button class="btn ghost" id="again" type="button">Try the quiz again</button><button class="btn" id="n" type="button">To the recap · ${upNext(3)} →</button></div></div></div>`; $('#again').onclick = () => { st.q = { i: 0, right: 0, answered: [] }; st.qStart = new Date().toISOString(); store.save(); quiz(body); }; $('#n').onclick = () => { st.quiz = { right: st.q.right, total: Q.length }; advance(); }; return; }
    if (st.q.i === 0) st.qStart = new Date().toISOString();
    body.innerHTML = `<div class="card quiz"><div class="row between"><p class="eyebrow">Question ${st.q.i + 1} of ${Q.length}${q.type ? ' · ' + ({ predict: 'Predict the result', fix: 'Fix the formula', which: 'Which tool?', mc: 'Choose one' }[q.type] || '') : ''}</p>${bar(st.q.i / Q.length, 'thin')}</div>
      <h2>${md(q.q)}</h2>${q.grid ? `<div class="mini-grid" data-grid='${esc(JSON.stringify(q.grid))}'></div>` : ''}${q.code ? `<p><code class="big">${esc(q.code)}</code></p>` : ''}
      ${q.answers ? `<form class="typed"><input class="ans" autocomplete="off" spellcheck="false" placeholder="Type your answer"><button class="btn">${tr('Check')}</button></form>` : `<div class="opts">${q.options.map((o, i) => `<button class="opt" data-i="${i}" type="button">${md(o)}</button>`).join('')}</div>`}<div class="feedback" role="status"></div></div>`;
    await hydrate(body, c.id);
    const finish = (ok, given) => { const fb = $('.feedback', body); if (ok) st.q.right++; else { store.logError({ course: c.id, lesson: id, lessonNum: L.num, task: 'quiz-' + st.q.i, kind: 'quiz', prompt: q.q, yours: given, right: q.answers ? q.answers[0] : q.options[q.answer], why: q.explain }); }
      fb.innerHTML = `<p class="${ok ? 'ok' : 'no'}">${ok ? '✓ Right.' : '✗ Not this time.'} ${md(q.explain)}</p><button class="btn" id="nx" type="button">${st.q.i < Q.length - 1 ? 'Next question →' : 'See my score →'}</button>`; $$('.opt', body).forEach(b => b.disabled = true); st.q.i++; store.save(); $('#nx').onclick = () => quiz(body); $('#nx').focus(); };
    if (q.answers) $('.typed', body).onsubmit = e => { e.preventDefault(); const v = $('.ans', body).value.trim(); if (!v) return; const norm = s => s.replace(/\s+/g, '').toUpperCase(); finish(q.answers.some(a => norm(a) === norm(v)), v); };
    else $$('.opt', body).forEach(b => b.onclick = () => { const i = +b.dataset.i; b.classList.add(i === q.answer ? 'right' : 'wrong'); if (i !== q.answer) $$('.opt', body)[q.answer].classList.add('right'); finish(i === q.answer, q.options[i]); });
  }
  // ---------- 5. recap ----------
  async function recap(body) {
    const R = L.recap; const fresh = R.cards.filter(cd => !store.d.cards[cd.id]).length;
    body.innerHTML = `<div class="card recap"><div class="row">${mascot('celebrate')}<div class="grow"><p class="eyebrow">Recap</p><h2>What you can do now</h2><ul class="ticks">${R.points.map(p => `<li>${md(p)}</li>`).join('')}</ul><div class="cheat"><b>Keep this line:</b> ${md(R.cheat)}</div></div></div>
      <h3>${R.cards.length} review cards ${st.status === 'done' ? 'in your deck' : 'go into your deck'} ${qhelp('cards')}</h3><div class="cardgrid">${R.cards.map(cd => `<div class="mcard"><span class="ctype">${esc(cd.type)}</span><p>${md(cd.front)}</p><details><summary>Answer</summary><p>${md(cd.back)}</p></details></div>`).join('')}</div>
      <div class="row between">${prev ? `<a class="btn ghost" href="#/lesson/${c.id}/${prev.id}">← ${esc(prev.title)}</a>` : '<span></span>'}<div class="row"><button class="btn big" id="fin" type="button">${st.status === 'done' ? 'Lesson complete ✓' : 'Finish lesson'}</button></div></div></div>`;
    $('#fin').onclick = () => {
      const first = st.status !== 'done';
      for (const cd of R.cards) if (!store.d.cards[cd.id]) store.d.cards[cd.id] = { ...cd, lesson: id, course: c.id, isNew: true, added: dayNum(), due: dayNum() };
      if (first) { st.status = 'done'; st.doneAt = new Date().toISOString(); store.act('lessons'); confetti(); }
      store.save();
      const info = moduleInfo(c, meta.module); const modEnd = info.done === info.total; const nextOpen = next && canOpenLesson(c, next);
      const upnext = modEnd && !info.passed ? `<a class="btn" href="#/mtest/${c.id}/${meta.module.n}">Up next: the Module ${meta.module.n} test · about 25 min</a>` : nextOpen ? `<a class="btn" href="#/lesson/${c.id}/${next.id}">Up next: ${esc(lessonLabel(next))} · ${esc(next.title)} · ${next.minutes} min</a>` : next ? `<a class="btn" href="#/module/${c.id}/${next.module.n}">Up next: Module ${next.module.n} (see how it opens)</a>` : '';
      const m = modal(`<div class="done-modal">${mascot('celebrate')}<h2>${esc(line('done'))}</h2><p>${fresh} new card${fresh === 1 ? '' : 's'} added to your review deck.${modEnd && !info.passed ? ` That’s every lesson in Module ${meta.module.n}: the module test is ready when you are.` : ''}</p><div class="row center">${upnext}<a class="btn ghost" href="#/practice/${c.id}/${id}">Practise this lesson</a><a class="btn ghost" href="#/home">Back to home</a></div></div>`);
      $$('a', m.el).forEach(a => a.addEventListener('click', () => m.close()));
    };
  }
});
// the lesson outline (N7): every slide, task and question, with ticks; parts not reached yet are listed but not linked
function outline(c, L, st, reach, id) {
  const done = i => i < reach() || st.status === 'done'; const base = `#/lesson/${c.id}/${id}`;
  const m = modal(`<h2>${esc(L.title)} · outline ${''}</h2><p class="muted small">Ticks show what you’ve done. Parts open in order.</p><ol class="outline">
    <li><b>1. Lecture</b>${done(0) ? ' ✓' : ''}<ol>${L.lecture.map((s, si) => `<li><span class="muted">${esc(s.title)}</span><ol>${s.slides.map((sl, k) => `<li><a href="${base}/lecture/${si + 1}-${k + 1}">${md(sl.h)}</a></li>`).join('')}<li><a href="${base}/lecture/${si + 1}-${s.slides.length + 1}">Quick check</a></li></ol></li>`).join('')}</ol></li>
    <li><b>${reach() >= 1 ? `<a href="${base}/demo">2. Demo</a>` : '2. Demo'}</b>${done(1) ? ' ✓' : ''}</li>
    <li><b>${reach() >= 2 ? `<a href="${base}/lab">3. Lab</a>` : '3. Lab'}</b>${done(2) ? ' ✓' : ''}<ol>${L.lab.tasks.map(t => `<li class="${st.tasks && st.tasks[t.id] && st.tasks[t.id].done ? 'ok' : ''}">${st.tasks && st.tasks[t.id] && st.tasks[t.id].done ? '✓ ' : ''}${md(String(t.prompt).slice(0, 110))}</li>`).join('')}</ol></li>
    <li><b>${reach() >= 3 ? `<a href="${base}/quiz">4. Quiz</a>` : '4. Quiz'}</b>${done(3) ? ' ✓' : ''} <span class="muted small">${L.quiz.length} questions</span></li>
    <li><b>${reach() >= 4 ? `<a href="${base}/recap">5. Recap</a>` : '5. Recap'}</b>${st.status === 'done' ? ' ✓' : ''}</li></ol>`, { wide: false });
  $$('a', m.el).forEach(a => a.addEventListener('click', () => m.close()));
}
function rc(a) { const m = /^([A-Z]+)(\d+)$/.exec(a); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2] - 1, c - 1]; }
export function resultsHTML(res) { return `<div class="score ${res.pct === 100 ? 'ok' : ''}"><b>${res.pct}%</b> · ${res.score} of ${res.total} checks passed</div><ul class="checks">${res.results.map(r => `<li class="${r.ok ? 'ok' : 'no'}"><span>${r.ok ? '✓' : '✗'}</span><div><b>${esc(r.label)}</b>${r.ok ? '' : `<p>${esc(r.msg)}</p>`}</div></li>`).join('')}</ul>`; }
function bumpSkill(id) { const s = store.lesson(id); s.firstTries = (s.firstTries || 0) + 1; }
async function resources(c, L) {
  const m = modal(`<h2>Downloads &amp; screenshots · ${esc(L.title)}</h2><div class="dl-list">${(L.downloads || []).map(d => `<button class="dl" data-file="${d.file}" data-name="${esc(d.name)}" type="button"><span class="ext ${d.file.split('.').pop()}">${d.file.split('.').pop().toUpperCase()}</span><span><b>${esc(d.label)}</b><small>${esc(d.desc || '')}</small></span></button>`).join('')}</div>
   <h3>How it should look</h3><div class="gallery">${(L.shots || []).map(s => `<figure class="shot"><button class="zoomable" type="button" aria-label="Enlarge screenshot"><img data-asset="${s.file}" alt="${esc(s.caption)}" data-callouts='${esc(JSON.stringify(s.callouts || []))}' data-caption="${esc(s.caption)}"><span class="zoomhint" aria-hidden="true">⤢ Enlarge</span></button><figcaption>${md(s.caption)}</figcaption></figure>`).join('') || '<p class="muted">No screenshots for this lesson.</p>'}</div>`, { wide: true });
  $$('[data-file]', m.el).forEach(b => b.onclick = async () => { try { const url = await assetURL(c.id, b.dataset.file); const a = document.createElement('a'); a.href = url; a.download = b.dataset.name; a.click(); } catch (e) { toast('Sign in to download this file.'); } });
  await hydrate(m.el, c.id);
}
const wait = ms => new Promise(r => setTimeout(r, ms));
