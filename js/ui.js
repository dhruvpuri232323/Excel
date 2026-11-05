// Small UI helpers: escaping, toasts, dialogs, popovers, the mascot, charts drawn as SVG, confetti, dates.
import { store } from './store.js';
export const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
// light markup for lesson text: **bold**, `code`, [[Key]] keyboard keys, _italic_
// keyboard keys follow the learner's computer: Windows keys by default, Mac keys when chosen in Settings
const KEYMAP = {}; export const setKeyMap = list => { for (const k of list || []) if (k.win && k.mac) KEYMAP[k.win.toLowerCase()] = k.mac; };
export const keyText = k => { if (!store.d || store.d.settings.os !== 'mac') return k; const hit = KEYMAP[k.toLowerCase()]; if (hit) return hit; return k.replace(/\bCtrl\b/g, '⌘').replace(/\bAlt\b/g, 'Option'); };
export const md = s => esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\[\[([^\]]+)\]\]/g, (m, k) => keyText(k).split('+').map(x => `<kbd>${x}</kbd>`).join('+')).replace(/(^|\s)_(.+?)_(?=\s|[.,;:!?]|$)/g, '$1<i>$2</i>').replace(/(^|[\s(])\*([^*\s][^*]*?)\*(?=[\s.,;:!?)]|$)/g, '$1<i>$2</i>');
export const $ = (s, r = document) => r.querySelector(s); export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export function toast(msg, kind = '') { const t = document.createElement('div'); t.className = 'toast ' + kind; t.setAttribute('role', 'status'); t.innerHTML = msg; document.body.appendChild(t); requestAnimationFrame(() => t.classList.add('on')); setTimeout(() => { t.classList.remove('on'); setTimeout(() => t.remove(), 300); }, 3800); }
export function modal(html, opts = {}) {
  const wrap = document.createElement('div'); wrap.className = 'modal-wrap'; wrap.innerHTML = `<div class="modal ${opts.wide ? 'wide' : ''}" role="dialog" aria-modal="true">${opts.close === false ? '' : '<button class="modal-x" aria-label="Close">×</button>'}${html}</div>`;
  document.body.appendChild(wrap); const close = () => { wrap.remove(); opts.onClose && opts.onClose(); };
  wrap.addEventListener('mousedown', e => { if (e.target === wrap && opts.close !== false) close(); }); const x = wrap.querySelector('.modal-x'); if (x) x.onclick = close;
  addEventListener('keydown', function k(e) { if (e.key === 'Escape' && opts.close !== false) { close(); removeEventListener('keydown', k); } });
  const f = wrap.querySelector('input,button:not(.modal-x),select,textarea'); f && f.focus(); return { el: wrap.firstElementChild, close };
}
export function confirmBox(text, yes = 'Yes', no = 'Cancel') { return new Promise(res => { const m = modal(`<p class="lead">${text}</p><div class="row end"><button class="btn ghost" data-a="0">${no}</button><button class="btn" data-a="1">${yes}</button></div>`, { onClose: () => res(false) }); m.el.querySelectorAll('[data-a]').forEach(b => b.onclick = () => { res(b.dataset.a === '1'); m.close(); }); }); } // resolve before closing: closing resolves “no”
export const mascot = (pose = 'hello', cls = '') => store.d && store.d.settings.mascot === false && !/brand-m|keep/.test(cls) ? '' : `<img class="mascot ${cls}" src="img/mascot/sheetah-${pose}.svg" alt="Sheetah ${pose}" width="120" height="132">`;
export function say(el, pose, text) { el.innerHTML = `<div class="say">${mascot(pose, 'sm')}<div class="bubble">${text}</div></div>`; }

// Sheetah's lines in the learner's chosen tone (from the CAT app's humour tones)
const LINES = {
  warm: { hi: n => `Hello ${n}! Ready for today’s bit of Excel?`, done: 'Lesson done. That’s real progress.', wrong: 'Not quite. Mistakes are how this sticks; here’s a nudge.', streak: n => `${n} days in a row. Lovely rhythm.`, back: 'Welcome back. Let’s pick up exactly where you stopped.', review: n => `${n} cards are ready for review. Ten minutes keeps them fresh.`, idle: 'Even ten minutes today keeps the streak alive.' },
  cheeky: { hi: n => `Oh hey ${n}. The spreadsheets missed you. Mostly.`, done: 'Done! You made that look easy. Suspiciously easy.', wrong: 'Ooh, close. Excel is picky. Peek at a hint?', streak: n => `${n}-day streak. Don’t make me sleep on your keyboard.`, back: 'You’re back! I kept your cell warm.', review: n => `${n} cards want attention. They’re needy like that.`, idle: 'Your streak is giving me puppy eyes. Well, cheetah eyes.' },
  coach: { hi: n => `${n}, today’s target is clear. Let’s go.`, done: 'Lesson complete. Next rep.', wrong: 'Wrong answer. Read the hint, fix it, move on.', streak: n => `${n} days straight. Protect it.`, back: 'Back to work. Resume point loaded.', review: n => `${n} reviews due. Clear them first.`, idle: 'No session yet today. Ten minutes. Now.' },
  desi: { hi: n => `Arre ${n}! Chalo, aaj thoda Excel ho jaye?`, done: 'Ho gaya! Ekdum first class.', wrong: 'Thoda sa miss hua. Hint dekh lo, ho jayega.', streak: n => `${n} din lagataar! Kya baat hai.`, back: 'Welcome back! Wahi se shuru karte hain.', review: n => `${n} cards revision ke liye ready hain. Bas 10 minute.`, idle: 'Aaj ka 10 minute abhi baaki hai, yaar.' },
};
export const line = (key, arg) => { const t = LINES[store.d.settings.tone] || LINES.warm; const v = t[key] || LINES.warm[key]; return typeof v === 'function' ? v(arg) : v; };

export function confetti() {
  if (store.d.settings.reduceMotion || store.d.settings.celebrate === false || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const c = document.createElement('canvas'); c.className = 'confetti'; c.width = innerWidth; c.height = innerHeight; document.body.appendChild(c); const x = c.getContext('2d');
  const cols = ['#F2A541', '#3557D6', '#23263A', '#FFF4E1', '#1F8A5B']; const ps = Array.from({ length: 140 }, () => ({ x: innerWidth / 2 + (Math.random() - .5) * 200, y: innerHeight / 3, vx: (Math.random() - .5) * 12, vy: -Math.random() * 12 - 4, s: 5 + Math.random() * 6, c: cols[Math.floor(Math.random() * cols.length)], r: Math.random() * 6 }));
  let t = 0; (function f() { x.clearRect(0, 0, c.width, c.height); for (const p of ps) { p.x += p.vx; p.y += p.vy; p.vy += .35; p.r += .1; x.save(); x.translate(p.x, p.y); x.rotate(p.r); x.fillStyle = p.c; x.fillRect(-p.s / 2, -p.s / 2, p.s, p.s * .6); x.restore(); } if (++t < 110) requestAnimationFrame(f); else c.remove(); })();
}

// ----- charts (hand-drawn SVG, theme-aware via CSS variables) -----
export function heatmap(activity, weeks = 18, studied) {
  const cell = 13, gap = 3; const end = new Date(); end.setHours(0, 0, 0, 0); const start = new Date(end); start.setDate(end.getDate() - (weeks * 7 - 1) - end.getDay());
  let out = ''; const days = []; for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) days.push(new Date(d));
  days.forEach((d, i) => { const k = d.toISOString().slice(0, 10); const iso = new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); const a = activity[iso]; const m = a ? a.minutes || 0 : 0; const lvl = !a ? 0 : m >= 45 ? 4 : m >= 25 ? 3 : m >= 10 || a.lessons ? 2 : 1; const col = Math.floor(i / 7), row = d.getDay(); out += `<rect x="${col * (cell + gap) + 26}" y="${row * (cell + gap) + 16}" width="${cell}" height="${cell}" rx="3" class="hm${lvl}"><title>${iso}: ${Math.round(m)} min${a && a.lessons ? `, ${a.lessons} lesson(s)` : ''}</title></rect>`; });
  const months = []; let lastWk = -9; days.forEach((d, i) => { const wk = Math.floor(i / 7); if ((d.getDate() === 1 || i === 0) && wk - lastWk >= 3) { lastWk = wk; months.push(`<text x="${Math.floor(i / 7) * (cell + gap) + 26}" y="10" class="ax">${d.toLocaleString('en', { month: 'short' })}</text>`); } });
  const w = Math.ceil(days.length / 7) * (cell + gap) + 30;
  return `<svg class="chart heat" viewBox="0 0 ${w} ${7 * (cell + gap) + 20}" role="img" aria-label="Study calendar, last ${weeks} weeks">${months.join('')}<text x="0" y="40" class="ax">Mon</text><text x="0" y="88" class="ax">Wed</text><text x="0" y="136" class="ax">Fri</text>${out}</svg>`;
}
export function bars(items, { h = 160, fmt = v => v, label = '' } = {}) {
  const w = 640, pad = 28, max = Math.max(1, ...items.map(i => i.v)); const bw = (w - pad * 2) / items.length;
  const b = items.map((it, i) => { const bh = (it.v / max) * (h - 44); return `<g><rect x="${pad + i * bw + bw * .15}" y="${h - 22 - bh}" width="${bw * .7}" height="${bh}" rx="4" class="${it.hl ? 'bar hl' : 'bar'}"><title>${esc(it.k)}: ${fmt(it.v)}</title></rect>${items.length <= 16 ? `<text x="${pad + i * bw + bw / 2}" y="${h - 6}" text-anchor="middle" class="ax">${esc(it.k)}</text>` : ''}${it.v && items.length <= 12 ? `<text x="${pad + i * bw + bw / 2}" y="${h - 27 - bh}" text-anchor="middle" class="val">${fmt(it.v)}</text>` : ''}</g>`; }).join('');
  return `<svg class="chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="${esc(label)}"><line x1="${pad}" x2="${w - pad}" y1="${h - 22}" y2="${h - 22}" class="axis"/>${b}</svg>`;
}
export function lineChart(points, { h = 170, max = 100, target, label = '' } = {}) {
  const w = 640, pad = 32; if (!points.length) return '<p class="muted">No data yet.</p>';
  const X = i => pad + (points.length === 1 ? (w - 2 * pad) / 2 : i * (w - 2 * pad) / (points.length - 1)), Y = v => h - 24 - (v / max) * (h - 44);
  const path = points.map((p, i) => (i ? 'L' : 'M') + X(i) + ' ' + Y(p.v)).join(' ');
  return `<svg class="chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="${esc(label)}">${[0, 50, 100].map(t => `<line x1="${pad}" x2="${w - pad}" y1="${Y(t * max / 100)}" y2="${Y(t * max / 100)}" class="grid"/><text x="4" y="${Y(t * max / 100) + 4}" class="ax">${Math.round(t * max / 100)}</text>`).join('')}${target ? `<line x1="${pad}" x2="${w - pad}" y1="${Y(target)}" y2="${Y(target)}" class="target"/><text x="${w - pad}" y="${Y(target) - 5}" text-anchor="end" class="ax">target ${target}</text>` : ''}<path d="${path}" class="line"/>${points.map((p, i) => `<circle cx="${X(i)}" cy="${Y(p.v)}" r="4.5" class="dot"><title>${esc(p.k)}: ${p.v}</title></circle>`).join('')}</svg>`;
}
export function ring(pct, label, size = 92) { const r = size / 2 - 8, c = 2 * Math.PI * r; const p = Math.max(0, Math.min(1, pct)); return `<svg class="ring" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" role="img" aria-label="${esc(label)}"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" class="ring-bg"/><circle cx="${size / 2}" cy="${size / 2}" r="${r}" class="ring-fg" stroke-dasharray="${c}" stroke-dashoffset="${c * (1 - p)}" transform="rotate(-90 ${size / 2} ${size / 2})"/><text x="50%" y="52%" text-anchor="middle" dominant-baseline="middle" class="ring-t">${Math.round(p * 100)}%</text></svg>`; }
export const bar = (pct, cls = '') => `<div class="pbar ${cls}" role="progressbar" aria-valuenow="${Math.round(pct * 100)}" aria-valuemin="0" aria-valuemax="100"><i style="width:${Math.max(0, Math.min(100, pct * 100))}%"></i></div>`;
export const shuffle = a => { const b = a.slice(); for (let i = b.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [b[i], b[j]] = [b[j], b[i]]; } return b; };
export const fmtDate = iso => new Date(iso).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
export function download(name, data, type = 'application/octet-stream') { const blob = data instanceof Blob ? data : new Blob([data], { type }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }

// a small floating box next to an element (help tips, glossary words, map stops)
let openPop = null;
export function closePopover() { if (openPop) { openPop.remove(); openPop = null; } }
export function popover(anchor, html, { wide = false, cls = '' } = {}) {
  closePopover(); const p = document.createElement('div'); p.className = `popover ${wide ? 'wide' : ''} ${cls}`; p.setAttribute('role', 'dialog'); p.innerHTML = html; document.body.appendChild(p);
  const r = anchor.getBoundingClientRect(); const pw = p.offsetWidth, ph = p.offsetHeight;
  let left = Math.min(Math.max(8, r.left + r.width / 2 - pw / 2), innerWidth - pw - 8); let top = r.bottom + 8; if (top + ph > innerHeight - 8 && r.top - ph - 8 > 8) top = r.top - ph - 8;
  p.style.left = left + scrollX + 'px'; p.style.top = top + scrollY + 'px'; openPop = p;
  setTimeout(() => { const off = e => { if (!p.contains(e.target) && e.target !== anchor) { closePopover(); removeEventListener('mousedown', off, true); } }; addEventListener('mousedown', off, true); }, 0);
  return p;
}
addEventListener('keydown', e => { if (e.key === 'Escape') closePopover(); });
addEventListener('hashchange', closePopover);
export const qhelp = (key, label = 'What is this?') => `<button type="button" class="qhelp" data-help="${key}" aria-label="${esc(label)}">?</button>`;
export const crumbs = items => `<nav class="crumbs" aria-label="You are here">${items.map(([t, h], i) => i < items.length - 1 && h ? `<a href="${h}">${esc(t)}</a>` : `<span aria-current="${i === items.length - 1 ? 'page' : 'false'}">${esc(t)}</span>`).join('<span class="sep" aria-hidden="true">›</span>')}</nav>`;
export const waLink = (text, phone = '') => `https://wa.me/${String(phone).replace(/[^0-9]/g, '')}?text=${encodeURIComponent(text)}`;
export const mailLink = (to, subject, body) => `mailto:${encodeURIComponent(to || '')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
// dates as day / month / year boxes, so every browser shows them the Indian way (dd-mm-yyyy)
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export function dateField(name, iso = '', { from = new Date().getFullYear() - 1, to = new Date().getFullYear() + 3, label = '' } = {}) {
  const [y, m, d] = iso ? iso.split('-').map(Number) : [0, 0, 0];
  return `<span class="datef" data-date="${name}" role="group" aria-label="${esc(label || name)}"><select data-part="d" aria-label="Day"><option value="">Day</option>${Array.from({ length: 31 }, (_, i) => `<option ${d === i + 1 ? 'selected' : ''}>${i + 1}</option>`).join('')}</select><select data-part="m" aria-label="Month"><option value="">Month</option>${MONTHS.map((x, i) => `<option value="${i + 1}" ${m === i + 1 ? 'selected' : ''}>${x}</option>`).join('')}</select><select data-part="y" aria-label="Year"><option value="">Year</option>${Array.from({ length: to - from + 1 }, (_, i) => from + i).map(v => `<option ${y === v ? 'selected' : ''}>${v}</option>`).join('')}</select></span>`;
}
export function readDate(root, name) { const box = root.querySelector(`[data-date="${name}"]`); if (!box) return ''; const v = k => box.querySelector(`[data-part="${k}"]`).value; if (!v('d') || !v('m') || !v('y')) return ''; return `${v('y')}-${String(v('m')).padStart(2, '0')}-${String(v('d')).padStart(2, '0')}`; }
export const fmtDay = iso => new Date(iso + (iso.length === 10 ? 'T00:00:00' : '')).toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' });
