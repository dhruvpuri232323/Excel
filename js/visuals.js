// Diagrams for lesson slides. Each takes plain data from the lesson file and returns HTML, so a slide can show
// the idea instead of describing it: nested boxes, a cell address map, a static mini sheet with highlights,
// icon cards, a flow of steps, big numbers, an analogy, before/after sheets and a map of the Excel window.
import { esc, md } from './ui.js';

const colName = n => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const colNum = s => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const cell = a => { const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a.replace(/^.*!/, '')); return m ? [+m[2] - 1, colNum(m[1])] : null; };
const range = r => { const [a, b] = r.split(':'); const p = cell(a), q = cell(b || a); return { r1: Math.min(p[0], q[0]), c1: Math.min(p[1], q[1]), r2: Math.max(p[0], q[0]), c2: Math.max(p[1], q[1]) }; };
const inR = (R, r, c) => r >= R.r1 && r <= R.r2 && c >= R.c1 && c <= R.c2;
const CIRC = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨'];

// a static mini sheet: no engine, just what the picture needs
export function gridmap(v) {
  const cols = v.cols || 5, rows = v.rows || 6, cells = v.cells || {}, hl = (v.hl || []).map(h => ({ ...h, R: range(h.ref) })), marks = (v.marks || []).map((m, i) => ({ ...m, at: cell(m.ref), n: i }));
  const sp = v.spill ? range(v.spill) : null; // a spilled answer: dashed blue border round the block, as in Excel
  const warn = new Set(v.warn || []), sel = v.sel ? range(v.sel) : null, colHl = new Set((v.colHl || []).map(colNum)), rowHl = new Set((v.rowHl || []).map(n => n - 1));
  const fmt = x => typeof x === 'number' ? x.toLocaleString('en-IN') : x;
  const align = x => { if (typeof x === 'number') return 'num'; const t = String(x); if (/^(TRUE|FALSE)$/.test(t)) return 'mid'; return /^₹?[\d,]+(\.\d+)?%?$/.test(t) || /^\d{1,2}[-/][A-Za-z\d]{1,3}[-/]\d{2,4}$/.test(t) ? 'num' : ''; }; // numbers and dates sit on the right, text on the left, TRUE/FALSE in the middle, as in Excel
  let h = `<div class="gm ${v.compact ? 'compact' : ''}">${v.bar ? `<div class="gm-bar"><span class="gm-nb">${esc(v.bar.ref)}</span><span class="gm-fx">fx</span><span class="gm-f">${esc(v.bar.text)}</span></div>` : ''}<table><thead><tr><th class="gm-corner"></th>${Array.from({ length: cols }, (_, c) => `<th class="${colHl.has(c) || (sel && c >= sel.c1 && c <= sel.c2) ? 'on' : ''}">${colName(c)}</th>`).join('')}</tr></thead><tbody>`;
  for (let r = 0; r < rows; r++) {
    h += `<tr><th class="${rowHl.has(r) ? 'ron' : sel && r >= sel.r1 && r <= sel.r2 ? 'on' : ''}">${r + 1}</th>`;
    for (let c = 0; c < cols; c++) {
      const a = colName(c) + (r + 1); let raw = cells[a]; const isText = typeof raw === 'string' && raw.startsWith("'"); if (isText) raw = raw.slice(1);
      const tones = hl.filter(x => inR(x.R, r, c)).map(x => 'tone-' + (x.tone || 'blue')); const mk = marks.filter(m => m.at && m.at[0] === r && m.at[1] === c);
      const cls = [raw === undefined ? '' : isText ? '' : align(raw), ...tones, colHl.has(c) ? 'colhl' : '', rowHl.has(r) ? 'rowhl' : '', sel && inR(sel, r, c) ? 'insel' : '', v.active === a ? 'active' : '', (v.bold || []).includes(a) || (v.boldRow && r === 0) ? 'b' : '', v.headFill && r === 0 && raw !== undefined ? 'headfill' : '', ...(sp && inR(sp, r, c) ? ['sp', r === sp.r1 ? 'sp-t' : '', r === sp.r2 ? 'sp-b' : '', c === sp.c1 ? 'sp-l' : '', c === sp.c2 ? 'sp-r' : ''] : [])].filter(Boolean).join(' ');
      h += `<td class="${cls}">${warn.has(a) ? '<i class="gm-warn" title="Number stored as text"></i>' : ''}${raw === undefined ? '' : esc(fmt(raw))}${mk.map(m => `<span class="gm-mark">${CIRC[m.n]}</span>`).join('')}</td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  const legend = [...hl.filter(x => x.label).map(x => `<span class="gm-leg"><i class="tone-${x.tone || 'blue'}"></i>${md(x.label)}</span>`), ...marks.filter(m => m.text).map(m => `<span class="gm-leg"><b>${CIRC[m.n]}</b> ${md(m.text)}</span>`)];
  if (legend.length) h += `<div class="gm-legend">${legend.join('')}</div>`;
  return h + '</div>';
}

const R = {
  // nested boxes: workbook ⊃ worksheet ⊃ cell
  layers(v) { const n = v.items.length; let h = ''; v.items.slice().reverse().forEach((it, k) => { const i = n - 1 - k; h = `<div class="lay lay${i}"><div class="lay-l"><span class="lay-i">${esc(it[0])}</span><b>${esc(it[1])}</b><small>${md(it[2] || '')}</small></div>${h}</div>`; }); return `<div class="d-layers">${h}</div>`; },
  // a cell address: the column letter and the row number, and where they meet
  addr(v) { const p = cell(v.ref); const col = v.ref.replace(/\d+$/, ''), row = v.ref.replace(/^[A-Z]+/, ''); const cols = Math.max(v.cols || 6, p[1] + 2), rows = Math.max(v.rows || 6, p[0] + 2);
    return `<div class="d-addr"><div class="ad-big"><span class="ad-col">${esc(col)}</span><span class="ad-row">${esc(row)}</span><div class="ad-labs"><span class="ad-col-l">column ${esc(col)}</span><span class="ad-row-l">row ${esc(row)}</span></div></div>${gridmap({ cols, rows, colHl: [col], rowHl: [+row], active: v.ref, cells: { [v.ref]: v.value ?? '' }, compact: true })}</div>`; },
  gridmap,
  // icon cards: a few things at a glance
  icons(v) { return `<div class="d-icons cols${v.cols || Math.min(4, v.items.length)}">${v.items.map(([i, t, s]) => `<div class="ic"><span class="ic-i">${esc(i)}</span><b>${md(t)}</b>${s ? `<small>${md(s)}</small>` : ''}</div>`).join('')}</div>`; },
  // what happens, step by step, with arrows
  flow(v) { return `<div class="d-flow">${v.items.map(([i, t, s], k) => `${k ? '<span class="fl-arrow" aria-hidden="true"></span>' : ''}<div class="fl"><span class="fl-i">${esc(i)}</span><b>${md(t)}</b>${s ? `<small>${md(s)}</small>` : ''}</div>`).join('')}</div>`; },
  // big numbers that make a point
  bignum(v) { return `<div class="d-big">${v.items.map(([n, l, s]) => `<div class="bn"><b>${esc(n)}</b><span>${md(l)}</span>${s ? `<small>${md(s)}</small>` : ''}</div>`).join('')}</div>`; },
  // something familiar on the left, the Excel idea on the right
  analogy(v) { return `<div class="d-ana">${v.title ? `<p class="ana-t">${md(v.title)}</p>` : ''}<div class="ana-g"><span class="ana-h">${esc(v.left || 'In everyday life')}</span><span></span><span class="ana-h">${esc(v.right || 'In Excel')}</span>${v.pairs.map(([a, b]) => `<div class="ana-a">${md(a)}</div><span class="ana-eq" aria-hidden="true">→</span><div class="ana-b">${md(b)}</div>`).join('')}</div></div>`; },
  // the same sheet before and after one action
  beforeafter(v) { return `<div class="d-ba"><div><span class="ba-l">${esc(v.before.label || 'Before')}</span>${gridmap({ compact: true, ...v.before })}</div><div class="ba-mid"><span class="ba-act">${md(v.action || '')}</span><span class="ba-arrow" aria-hidden="true"></span></div><div><span class="ba-l">${esc(v.after.label || 'After')}</span>${gridmap({ compact: true, ...v.after })}</div></div>`; },
  // a plain map of the Excel window, with the parts this slide is about lit up and numbered
  // a formula taken apart: each piece with its own colour and a label under it
  // the formula stays in one piece; labelled parts get a number above them and a line in the key below
  anatomy(v) { let n = 0; const P = v.parts.map(([t, l, tone]) => ({ t, l, tone: tone || 'plain', n: l ? ++n : 0 }));
    const len = P.reduce((a, p) => a + p.t.length, 0);
    return `<div class="d-anat"><div class="an-f ${len > 46 ? 'an-xlong' : len > 30 ? 'an-long' : ''}" role="img" aria-label="${esc(P.map(p => p.t).join(''))}">${P.map(p => `<span class="an-p an-${p.tone}">${p.n ? `<i class="an-n">${CIRC[p.n - 1] || p.n}</i>` : ''}${esc(p.t)}</span>`).join('')}</div>
      <ul class="an-key">${P.filter(p => p.n).map(p => `<li><span class="an-n2 an-${p.tone}">${CIRC[p.n - 1] || p.n}</span><code class="an-${p.tone}">${esc(p.t.trim())}</code><span>${md(p.l)}</span></li>`).join('')}</ul>
      ${v.result !== undefined ? `<p class="an-res">→ <b>${esc(v.result)}</b>${v.resultNote ? ` <span>${md(v.resultNote)}</span>` : ''}</p>` : ''}</div>`; },
  // how Excel works a formula out, one step at a time (like Formulas › Evaluate Formula); the part being worked out is marked
  evalsteps(v) { const mark = (f, hi) => { const t = esc(f); if (!hi) return t; const h = esc(hi); const k = t.indexOf(h); return k < 0 ? t : t.slice(0, k) + '<mark>' + h + '</mark>' + t.slice(k + h.length); };
    return `<ol class="d-eval">${v.steps.map(([f, note, hi], k) => `<li class="${k === v.steps.length - 1 ? 'last' : ''}"><span class="ev-n">${k === v.steps.length - 1 ? '=' : k + 1}</span><code>${mark(f, hi)}</code>${note ? `<span class="ev-note">${md(note)}</span>` : ''}</li>`).join('')}</ol>`; },
  windowmap(v) { const on = v.hl || []; const z = (k, label, extra = '') => { const i = on.indexOf(k); return `<div class="wz wz-${k} ${i >= 0 ? 'on' : ''}" ${extra}>${i >= 0 ? `<span class="wz-n">${i + 1}</span>` : ''}<span class="wz-l">${label}</span></div>`; };
    return `<div class="d-win" role="img" aria-label="A map of the Excel window">${z('qat', 'Quick Access Toolbar · file name · AutoSave')}${z('tabs', 'File · Home · Insert · Page Layout · Formulas · Data · Review · View')}${z('ribbon', 'Ribbon: groups of buttons')}<div class="wz-row">${z('namebox', 'Name Box')}${z('fbar', 'fx  Formula bar')}</div>${z('grid', 'The grid: columns A, B, C… rows 1, 2, 3…')}${z('sheets', 'Sheet tabs  ⊕')}${z('status', 'Status bar: Average · Count · Sum · zoom')}</div>${v.legend ? `<ol class="wz-legend">${on.map((k, i) => `<li><span class="wz-n">${i + 1}</span>${md(v.legend[i] || k)}</li>`).join('')}</ol>` : ''}`; },
};
export const DIAGRAMS = Object.keys(R);
export function diagram(v) { const f = R[v.type]; if (!f) return ''; return `<div class="diag d-${v.type}-wrap">${f(v)}${v.caption ? `<p class="cap">${md(v.caption)}</p>` : ''}</div>`; }
