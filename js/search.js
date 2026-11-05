// Search everything (Ctrl+K or /): lessons, slides, tasks, functions, shortcuts, glossary terms, help and settings.
import { getJSON, catalog, course as loadCourse } from './content.js';
import { PAGES, ARTICLES } from './help.js';
import { esc, md, $, $$, modal, keyText } from './ui.js';
import { lessonLabel, lessonNo } from './access.js';
import { go } from './app.js';

let INDEX = null;
const strip = s => String(s || '').replace(/\*\*|`|_/g, '').replace(/\[\[([^\]]+)\]\]/g, '$1');
async function build() {
  const items = [];
  PAGES.forEach(([t, h, k]) => items.push({ t, s: 'Page', h, k, g: 'Pages and settings' }));
  ARTICLES.forEach(a => items.push({ t: a.title, s: 'Help', h: '#/help/' + a.id, k: a.tags + ' ' + a.body.join(' '), g: 'Help' }));
  const cat = await catalog().catch(() => ({ courses: [] }));
  for (const cc of cat.courses.filter(x => x.status === 'live')) {
    const c = await loadCourse(cc.id).catch(() => null); if (!c) continue;
    for (const m of c.modules) {
      items.push({ t: `Module ${m.n}: ${m.title}`, s: c.title, h: `#/module/${c.id}/${m.n}`, k: m.summary, g: 'Lessons' });
      for (const l of m.lessons) {
        const lab = lessonLabel(l, m);
        items.push({ t: l.title, s: lab, h: `#/lesson/${c.id}/${l.id}`, k: l.summary + ' ' + (l.objectives || []).join(' '), g: 'Lessons' });
        (l.topics || []).forEach(([si, k, h]) => items.push({ t: h, s: `${lab} · ${l.title}`, h: `#/lesson/${c.id}/${l.id}/lecture/${si + 1}-${k + 1}`, k: '', g: 'Topics' }));
        (l.tasks || []).forEach((p, i) => items.push({ t: p, s: `${lab} · lab task ${i + 1}`, h: `#/lesson/${c.id}/${l.id}/lab`, k: '', g: 'Tasks' }));
      }
    }
    const ch = await getJSON(`content/${c.id}/challenges.json`).catch(() => null); if (ch) ch.challenges.forEach(x => items.push({ t: x.title, s: 'Weekly challenge · ' + x.skills.join(', '), h: '#/challenge/' + x.id, k: x.story, g: 'Practice' }));
    const iv = await getJSON(`content/${c.id}/interview.json`).catch(() => null); if (iv) iv.questions.forEach((q, i) => items.push({ t: q.q, s: 'Interview prep', h: '#/interview/' + (i + 1), k: q.a, g: 'Practice' }));
  }
  const lib = await getJSON('content/library.json').catch(() => null);
  if (lib) {
    lib.functions.forEach(f => items.push({ t: f.name, s: strip(f.short), h: '#/library/' + f.name, k: f.what + ' ' + f.cat, g: 'Functions' }));
    lib.shortcuts.forEach(s => items.push({ t: s.what, s: keyText(s.win), h: '#/shortcuts', k: s.win + ' ' + s.mac + ' ' + s.cat, g: 'Shortcuts' }));
    lib.glossary.forEach(([t, d, c, l]) => items.push({ t, s: strip(d), h: l ? `#/lesson/${c}/${l}` : '#/library', k: d, g: 'Excel words' }));
  }
  items.forEach(x => { x.lt = strip(x.t).toLowerCase(); x.lk = (x.lt + ' ' + strip(x.s) + ' ' + strip(x.k)).toLowerCase(); });
  return items;
}
export async function searchIndex() { if (!INDEX) INDEX = await build(); return INDEX; }
export async function query(q) {
  const idx = await searchIndex(); const words = q.toLowerCase().split(/\s+/).filter(Boolean); if (!words.length) return [];
  const out = [];
  for (const x of idx) {
    if (!words.every(w => x.lk.includes(w))) continue;
    let sc = 0; const ql = q.toLowerCase().trim();
    if (x.lt === ql) sc += 100; else if (x.lt.startsWith(ql)) sc += 60; else if (x.lt.includes(ql)) sc += 35;
    sc += words.filter(w => x.lt.includes(w)).length * 10; if (x.g === 'Lessons' || x.g === 'Functions') sc += 4; if (x.g === 'Topics') sc += 3;
    out.push([sc, x]);
  }
  return out.sort((a, b) => b[0] - a[0]).slice(0, 40).map(a => a[1]);
}
const ORDER = ['Lessons', 'Topics', 'Functions', 'Tasks', 'Shortcuts', 'Excel words', 'Practice', 'Help', 'Pages and settings'];
export function openSearch(initial = '') {
  const m = modal(`<div class="search-box"><input id="sq" type="search" placeholder="Search lessons, topics, functions, settings…" aria-label="Search" autocomplete="off" value="${esc(initial)}"><p class="muted small" id="shint">Try “freeze panes”, “VLOOKUP”, “lakh”, “F4” or “dark theme”.</p><div id="sres" class="sres" role="listbox"></div></div>`, { wide: true });
  const inp = $('#sq', m.el), box = $('#sres', m.el); let sel = 0, res = [];
  const draw = () => {
    if (!inp.value.trim()) { box.innerHTML = ''; $('#shint', m.el).hidden = false; return; } $('#shint', m.el).hidden = true;
    if (!res.length) { box.innerHTML = `<p class="muted">Nothing found for “${esc(inp.value)}”. Try fewer or different words, or <a href="#/topics">browse the topic index</a>.</p>`; return; }
    const groups = {}; res.forEach((x, i) => (groups[x.g] = groups[x.g] || []).push([x, i]));
    box.innerHTML = order.filter(g => groups[g]).map(g => `<h4>${esc(g)}</h4>${groups[g].map(([x, i]) => `<a class="sitem ${i === sel ? 'on' : ''}" href="${x.h}" data-i="${i}" role="option"><b>${md(x.t)}</b><small>${md(String(x.s).slice(0, 120))}</small></a>`).join('')}`).join('');
    $$('.sitem', box).forEach(a => a.onclick = () => m.close());
    const on = $('.sitem.on', box); on && on.scrollIntoView({ block: 'nearest' });
  };
  let tmr; inp.oninput = () => { clearTimeout(tmr); tmr = setTimeout(async () => { res = await query(inp.value); sel = 0; flat(); draw(); }, 80); };
  let order = ORDER; const flat = () => { const groups = {}; res.forEach(x => (groups[x.g] = groups[x.g] || []).push(x)); order = Object.keys(groups); res = order.flatMap(g => groups[g]); }; // the group with the best match comes first
  inp.onkeydown = e => { if (e.key === 'ArrowDown') { e.preventDefault(); sel = Math.min(res.length - 1, sel + 1); draw(); } if (e.key === 'ArrowUp') { e.preventDefault(); sel = Math.max(0, sel - 1); draw(); } if (e.key === 'Enter' && res[sel]) { e.preventDefault(); m.close(); go(res[sel].h); } };
  inp.focus(); if (initial) inp.oninput();
}
