// Daily review: spaced repetition (FSRS) for cards, plus the error log's 1-3-7 day returns.
import { register, getCourse } from '../app.js';
import { store, today } from '../store.js';
import { esc, md, $, $$, mascot, line, toast, confetti, bar, shuffle, modal } from '../ui.js';
import { review as fsrs, previewIntervals, dayNum, AGAIN, HARD, GOOD, EASY } from '../srs.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { checkTask } from '../checker.js';

export function buildQueue() {
  const t = dayNum(); const cards = Object.entries(store.d.cards).map(([id, c]) => ({ id, ...c }));
  const due = cards.filter(c => !c.isNew && c.due <= t).sort((a, b) => a.due - b.due);
  const newToday = cards.filter(c => c.firstSeen === t).length; const room = Math.max(0, (store.d.settings.newPerDay || 15) - newToday);
  const fresh = cards.filter(c => c.isNew).sort((a, b) => a.added - b.added).slice(0, room);
  const errs = store.d.errors.filter(e => !e.resolved && e.due[e.step] <= t).slice(0, 15).map(e => ({ err: true, ...e }));
  // mix topics: interleave due cards from different lessons (interleaving beats blocked practice)
  const byLesson = {}; for (const c of due) (byLesson[c.lesson] = byLesson[c.lesson] || []).push(c);
  const mixed = []; let more = true; while (more) { more = false; for (const k of Object.keys(byLesson)) { const x = byLesson[k].shift(); if (x) { mixed.push(x); more = true; } } }
  return [...mixed, ...errs, ...fresh];
}
register('review', async main => {
  let q = buildQueue(); let i = 0, done = 0; const start = Date.now();
  const total = q.length;
  if (!total) { main.innerHTML = `<div class="card empty">${`<img class="mascot" src="img/mascot/fil.svg" alt="Fil, the fill handle" width="110">`}<div><h1>Nothing due today</h1><p class="lead">Your review deck is clear. Cards come back just before you’d forget them, so check in tomorrow.</p><p>${Object.keys(store.d.cards).length ? `You have ${Object.keys(store.d.cards).length} cards in your deck. Next due: ${nextDue()}.` : 'Finish a lesson to add your first cards.'}</p><div class="row"><a class="btn" href="#/home">Back home</a><a class="btn ghost" href="#/tests">Take a practice test</a></div></div></div>${forecast()}`; return; }
  const show = () => {
    if (i >= q.length) { store.act('reviews', 0); const mins = Math.max(1, Math.round((Date.now() - start) / 60000)); main.innerHTML = `<div class="card empty">${mascot('celebrate')}<div><h1>Review done</h1><p class="lead">${done} items in about ${mins} minute${mins === 1 ? '' : 's'}. ${esc(line('done'))}</p><div class="row"><a class="btn" href="#/home">Back home</a><a class="btn ghost" href="#/errors">Open Mistakes</a></div></div></div>${forecast()}`; confetti(); return; }
    const it = q[i];
    main.innerHTML = `<div class="review"><div class="row between"><h1>Daily review</h1><span class="muted">${i + 1} of ${q.length}</span></div>${bar(i / q.length)}<div class="card rcard" id="rc"></div></div>`;
    const box = $('#rc');
    if (it.err) {
      box.innerHTML = `<p class="eyebrow">From your error log · lesson ${esc(it.lessonNum || '')} · return ${it.step + 1} of 3</p><h2>${md(it.prompt)}</h2><p class="muted">Last time you answered: <code>${esc(it.yours)}</code></p><p>Work out the right answer in your head (or in Excel), then reveal it.</p><button class="btn" id="rv">Show the right answer</button><div id="ans" hidden><div class="answer"><p><b>Right answer:</b> <code>${esc(it.right)}</code></p><p>${md(it.why || '')}</p></div><div class="row"><button class="btn ghost" data-g="0">Still shaky</button><button class="btn" data-g="1">I’ve got it now</button></div></div>`;
      $('#rv').onclick = () => { $('#ans').hidden = false; $('#rv').hidden = true; };
      $$('[data-g]', box).forEach(b => b.onclick = () => { const e = store.d.errors.find(x => x.id === it.id); if (b.dataset.g === '1') { e.step++; if (e.step >= 3) { e.resolved = true; e.resolvedAt = new Date().toISOString(); toast('Fixed for good. Moved to resolved.'); } } else { e.step = 0; const t0 = dayNum(); e.due = [t0 + 1, t0 + 3, t0 + 7]; } store.act('reviews'); store.save(); done++; i++; show(); });
      return;
    }
    const card = store.d.cards[it.id]; const prevI = previewIntervals(card.isNew ? {} : card);
    const fmtI = d => d === 0 ? 'again today' : d === 1 ? '1 day' : d < 30 ? d + ' days' : d < 365 ? Math.round(d / 30) + ' mo' : Math.round(d / 365) + ' yr';
    const rate = g => { const t = dayNum(); const base = card.isNew ? {} : card; const n = fsrs(base, g, t); Object.assign(card, { s: n.s, d: n.d, reps: n.reps, lapses: n.lapses || 0, last: n.last, due: n.due }); if (card.isNew) { card.isNew = false; card.firstSeen = t; } store.act('reviews'); store.save(); done++; if (g === AGAIN) q.splice(Math.min(q.length, i + 4), 0, it); i++; show(); };
    if (card.type === 'task' && card.task) {
      box.innerHTML = `<p class="eyebrow">Mini-task · ${esc(card.lessonNum || card.lesson)}</p><h2>${md(card.front)}</h2><div id="mg"></div><div class="row"><button class="btn" id="ck">Check</button><button class="btn ghost" id="gu">Show me</button></div><div class="tfb"></div>`;
      const g = new Grid($('#mg'), new Workbook(card.task.sheet), { height: 200, toolbar: false, editable: card.task.editable, highlight: card.task.editable, tabs: false, select: card.task.editable && card.task.editable[0].split(':')[0] });
      let tries = 0;
      $('#ck').onclick = () => { tries++; const res = checkTask(card.task.check, { wb: g.wb, grid: g }); const fb = $('.tfb', box); if (res.ok) { fb.innerHTML = `<p class="ok">✓ ${md(res.msg)}</p>`; setTimeout(() => rate(tries === 1 ? GOOD : HARD), 900); } else fb.innerHTML = `<p class="no">${md(res.msg)}</p>`; };
      $('#gu').onclick = () => { $('.tfb', box).innerHTML = `<p>${md(card.back)}</p><button class="btn sm" id="ag">Mark as missed</button>`; $('#ag').onclick = () => rate(AGAIN); };
      return;
    }
    box.innerHTML = `<p class="eyebrow">${esc({ concept: 'Concept', syntax: 'Syntax', predict: 'Predict the result', fix: 'Fix the formula', which: 'Which tool?', shortcut: 'Shortcut' }[card.type] || 'Card')}${card.isNew ? ' · new' : ''}</p><h2>${md(card.front)}</h2><button class="btn" id="rv">Show answer <kbd>Space</kbd></button>
      <div id="ans" hidden><div class="answer">${md(card.back)}</div><p class="muted small">How well did you remember?</p><div class="grades">${[['Again', AGAIN], ['Hard', HARD], ['Good', GOOD], ['Easy', EASY]].map(([n, g], k) => `<button class="btn grade g${g}" data-g="${g}">${n}<small>${fmtI(prevI[k])}</small></button>`).join('')}</div></div>`;
    const reveal = () => { $('#ans').hidden = false; $('#rv').hidden = true; $('[data-g="3"]', box).focus(); };
    $('#rv').onclick = reveal; $$('[data-g]', box).forEach(b => b.onclick = () => rate(+b.dataset.g));
    document.onkeydown = e => { if (!document.getElementById('rc')) { document.onkeydown = null; return; } if (e.key === ' ' && $('#ans').hidden) { e.preventDefault(); reveal(); } else if (!$('#ans').hidden && ['1', '2', '3', '4'].includes(e.key)) rate(+e.key); };
  };
  show();
});
function nextDue() { const ds = Object.values(store.d.cards).map(c => c.due).filter(Boolean).sort((a, b) => a - b); if (!ds.length) return '—'; const d = ds[0] - dayNum(); return d <= 0 ? 'today' : d === 1 ? 'tomorrow' : `in ${d} days`; }
export function forecast() {
  const t = dayNum(); const counts = Array.from({ length: 14 }, (_, k) => Object.values(store.d.cards).filter(c => !c.isNew && (k === 0 ? c.due <= t : c.due === t + k)).length);
  const max = Math.max(1, ...counts); return `<section class="card"><h2>Cards due · next 14 days</h2><div class="fc">${counts.map((n, k) => `<div title="${n} due"><i style="height:${n / max * 70 + 2}px"></i><small>${k === 0 ? 'Today' : new Date(Date.now() + k * 864e5).toLocaleDateString('en-IN', { weekday: 'narrow' })}</small></div>`).join('')}</div>
  <h3>Tricky cards</h3>${trickyList()}</section>`;
}
function trickyList() { const tricky = Object.values(store.d.cards).filter(c => (c.lapses || 0) >= 2).sort((a, b) => b.lapses - a.lapses).slice(0, 6); return tricky.length ? `<ul class="tricky">${tricky.map(c => `<li>${md(c.front)} <span class="muted">· forgotten ${c.lapses}×</span></li>`).join('')}</ul>` : '<p class="muted">None yet. Cards you forget twice or more show up here.</p>'; }
