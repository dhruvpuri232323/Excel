// The guided tour (7 stops) and one-time tips that point at things the first time they appear.
import { store } from './store.js';
import { esc, md } from './ui.js';

const isPhone = () => matchMedia('(max-width: 820px)').matches;
const STEPS = [
  { sel: '.side nav', phone: '.bottombar', title: 'Everything lives here', text: 'The side panel has four groups: **Learn**, **Practise**, **Track** and **My stuff**. Click a group name to fold it. The « button shrinks the whole panel to icons.' },
  { sel: '[data-nav="courses"]', phone: '[data-tab="map"]', title: 'Your learning map', text: '**My courses** opens the map of your course: where you are, what you’ve finished and what opens next.' },
  { sel: '#todayCard', title: 'Today’s plan', text: 'Your next step and a short checklist for the day. Each line is a button that takes you straight there.' },
  { sel: '#topStats', title: 'Your numbers', text: 'Your study streak, today’s minutes against your goal, and how many review cards are waiting. Click one to open its page.' },
  { sel: '#searchBtn', phone: '[data-tab="search"]', title: 'Search everything', text: 'Find any lesson, topic, function or setting. Press **Ctrl+K** or **/** from anywhere.' },
  { sel: '#helpBtn', title: 'Help is one click away', text: 'Short answers with pictures, and an **Ask your teacher** button. Look out for small **?** buttons too: they explain the thing next to them.' },
  { sel: '[data-nav="settings"]', phone: '[data-tab="more"]', title: 'Make it yours', text: 'Theme, text size, colours, tone, goals and reminders all live in Settings, and each change applies straight away. You can run setup or this tour again from there.' },
];
export function startTour(onDone) {
  let i = 0; const ov = document.createElement('div'); ov.className = 'tour'; document.body.appendChild(ov);
  const end = () => { ov.remove(); removeEventListener('resize', draw); removeEventListener('keydown', key); store.d.tourDone = true; store.save(); onDone && onDone(); };
  const key = e => { if (e.key === 'Escape') end(); if (e.key === 'ArrowRight') next(); if (e.key === 'ArrowLeft' && i > 0) { i--; draw(); } };
  const next = () => { i++; if (i >= STEPS.length) end(); else draw(); };
  function draw() {
    const s = STEPS[i]; let el = document.querySelector(isPhone() && s.phone ? s.phone : s.sel);
    if (!el || !el.getClientRects().length) { if (i < STEPS.length - 1) { i++; return draw(); } return end(); }
    el.scrollIntoView({ block: 'nearest' }); const r = el.getBoundingClientRect(); const pad = 6;
    ov.innerHTML = `<div class="tour-hole" style="left:${r.left - pad}px;top:${r.top - pad}px;width:${r.width + pad * 2}px;height:${r.height + pad * 2}px"></div>
     <div class="tour-tip" role="dialog" aria-live="polite"><p class="eyebrow">Tour · ${i + 1} of ${STEPS.length}</p><h3>${esc(s.title)}</h3><p>${md(s.text)}</p>
     <div class="row between"><button class="linkbtn" data-a="skip">Skip the tour</button><div class="row">${i ? '<button class="btn ghost sm" data-a="back">Back</button>' : ''}<button class="btn sm" data-a="next">${i === STEPS.length - 1 ? 'Finish' : 'Next'}</button></div></div></div>`;
    const tip = ov.querySelector('.tour-tip'); const tw = Math.min(340, innerWidth - 24); tip.style.width = tw + 'px';
    let left = Math.min(Math.max(12, r.left), innerWidth - tw - 12), top = r.bottom + 14; if (top + tip.offsetHeight > innerHeight - 12) top = Math.max(12, r.top - tip.offsetHeight - 14);
    if (r.width < 300 && r.right + tw + 20 < innerWidth && r.height > 200) { left = r.right + 18; top = Math.max(12, Math.min(r.top, innerHeight - tip.offsetHeight - 12)); }
    tip.style.left = left + 'px'; tip.style.top = top + 'px';
    ov.querySelector('[data-a="next"]').onclick = next; ov.querySelector('[data-a="skip"]').onclick = end; const b = ov.querySelector('[data-a="back"]'); if (b) b.onclick = () => { i--; draw(); };
    ov.querySelector('[data-a="next"]').focus();
  }
  addEventListener('resize', draw); addEventListener('keydown', key); draw();
}

// one-time tips: shown the first time their element appears (normally in the first lesson), gone once used or dismissed
const TIPS = {
  parts: 'These are the five parts of every lesson. Each opens when you finish the one before.',
  next: 'Use **Next** (or the → key) to move through the slides.',
  outline: 'The **Outline** lists every slide and task, so you can jump straight to one.',
  check: 'Do the task on the sheet, then press **Check**. Sheetah tells you exactly what’s right or not.',
  hint: 'Stuck? Hints come one at a time. After three, you can see the solution.',
  grid: 'This is a real working sheet. Click a cell and type; press **Enter** to finish. Highlighted cells are yours to fill.',
};
let tipEl = null;
export function tip(key, el) {
  if (!el || !TIPS[key] || store.d.tips[key] || tipEl || store.preview || document.querySelector('.tour')) return;
  const r = el.getBoundingClientRect(); if (!r.width) return;
  tipEl = document.createElement('div'); tipEl.className = 'coach'; tipEl.setAttribute('role', 'status');
  tipEl.innerHTML = `<p>${md(TIPS[key])}</p><button class="btn sm">Got it</button>`; document.body.appendChild(tipEl);
  const w = Math.min(280, innerWidth - 24); tipEl.style.width = w + 'px';
  tipEl.style.left = Math.min(Math.max(12, r.left), innerWidth - w - 12) + scrollX + 'px'; tipEl.style.top = r.bottom + 10 + scrollY + 'px';
  const done = () => { store.d.tips[key] = true; store.save(); if (tipEl) tipEl.remove(); tipEl = null; el.removeEventListener('click', done); };
  tipEl.querySelector('button').onclick = done; el.addEventListener('click', done);
  addEventListener('hashchange', () => { if (tipEl) { tipEl.remove(); tipEl = null; } }, { once: true });
}
export const clearTip = () => { if (tipEl) { tipEl.remove(); tipEl = null; } };
