// The chart builder: pick the data, the chart type and the design choices, and see the chart as you go. Used by
// lab tasks of kind 'chart', and by the practical tasks of module tests and final exams.
import { esc, $, $$ } from './ui.js';
import { chartSVG, BUILDER_DEFAULTS } from './charts.js';
import { chartOf } from './checker.js';
import { parseRange, rangeStr, addr } from './engine.js';

export const CHART_TYPES = { column: ['▮▮', 'Column'], bar: ['☰', 'Bar'], line: ['⟋', 'Line'], area: ['◢', 'Area'], pie: ['◔', 'Pie'], doughnut: ['◎', 'Doughnut'], scatter: ['⁘', 'Scatter'], combo: ['▮⟋', 'Combo'],
  histogram: ['▃▇▅', 'Histogram'], pareto: ['▇▃⟋', 'Pareto'], box: ['⊟', 'Box & Whisker'], waterfall: ['▀▄▀', 'Waterfall'], funnel: ['▼', 'Funnel'], treemap: ['▦', 'Treemap'], sunburst: ['✺', 'Sunburst'], stock: ['┼', 'Stock'], map: ['▩', 'Filled map'] };
const opt = (v, label) => `<option value="${v}">${label}</option>`;

// controls: range type title labels endlabels grid switch reverse legend axis highlight stacked secondary bins totals nofill invert
export function builderHTML(t) {
  const C = t.controls || ['range', 'type', 'title', 'labels', 'legend', 'grid', 'axis']; const has = k => C.includes(k);
  return `<div class="cb" role="group" aria-label="Chart builder">
    <div class="cb-row"><label class="cb-lab">Data <input data-k="range" class="cb-range" placeholder="like A1:B13" spellcheck="false" autocomplete="off"></label><button class="btn ghost sm cb-sel" type="button">Use selection</button></div>
    ${has('type') ? `<div class="cb-types">${(t.allow || ['column', 'bar', 'line', 'pie', 'doughnut', 'scatter', 'combo']).map(k => `<button type="button" class="cb-type" data-type="${k}" title="${CHART_TYPES[k][1]}"><span>${CHART_TYPES[k][0]}</span>${CHART_TYPES[k][1]}</button>`).join('')}</div>` : ''}
    ${has('title') ? `<label class="cb-lab cb-full">Title <input data-k="title" placeholder="${has('titlecell') ? 'Type a title, or = and a cell, like =B20' : 'Say the finding, not the topic'}" autocomplete="off"></label>` : ''}
    <div class="cb-opts">${has('labels') ? '<label><input type="checkbox" data-k="labels"> Data labels</label>' : ''}${has('endlabels') ? '<label><input type="checkbox" data-k="endLabels"> Label lines at the end</label>' : ''}${has('grid') ? '<label><input type="checkbox" data-k="grid"> Gridlines</label>' : ''}${has('switch') ? '<label><input type="checkbox" data-k="byRow"> Switch row/column</label>' : ''}${has('reverse') ? '<label class="cb-bar"><input type="checkbox" data-k="reverse"> Categories in reverse order</label>' : ''}${has('nofill') ? '<label><input type="checkbox" data-k="hideFirst"> First series: no fill</label>' : ''}${has('invert') ? '<label><input type="checkbox" data-k="invert"> Invert if negative</label>' : ''}
      ${has('legend') ? `<label>Legend <select data-k="legend">${opt('auto', 'Automatic')}${opt('none', 'None')}${opt('bottom', 'Bottom')}${opt('right', 'Right')}${opt('top', 'Top')}</select></label>` : ''}${has('axis') ? `<label>Axis starts at <select data-k="yMinMode">${opt('auto', 'Automatic')}${opt('0', '0')}${opt('custom', 'Custom…')}</select><input data-k="yMinNum" class="cb-num" type="number" step="any" aria-label="Axis minimum" hidden></label>` : ''}
      ${has('highlight') ? '<label>Highlight <select data-k="highlight"><option value="">None</option></select></label>' : ''}${has('stacked') ? `<label>Series <select data-k="stacked">${opt('', 'Side by side')}${opt('stacked', 'Stacked')}${opt('percent', '100% stacked')}</select></label>` : ''}${has('secondary') ? '<label class="cb-combo">Line on 2nd axis <select data-k="secondary"><option value="">None</option></select></label>' : ''}
      ${has('bins') ? `<label class="cb-hist">Bins <select data-k="bins">${opt('auto', 'Automatic')}${opt('width', 'Bin width')}${opt('count', 'Number of bins')}</select><input data-k="binVal" class="cb-num" type="number" step="any" min="0" aria-label="Bin width or number of bins"></label><label class="cb-hist">Underflow bin ≤ <input data-k="underflow" class="cb-num" type="number" step="any" aria-label="Underflow bin"></label><label class="cb-hist">Overflow bin > <input data-k="overflow" class="cb-num" type="number" step="any" aria-label="Overflow bin"></label>` : ''}</div>
    ${has('totals') ? '<div class="cb-totals cb-wf" role="group" aria-label="Set as total"><span class="cb-tl">Set as total:</span><span class="cb-tlist"></span></div>' : ''}
    <div class="cb-prev"></div></div>`;
}

// box: the .cb element; cfg: the settings (changed in place); getWb/getSel: the sheet and its selected range;
// onSave(cfg) runs after every change. Returns draw(), to call when the sheet changes. box._apply(cfg) sets a whole answer.
export function wireBuilder(box, t, cfg, getWb, getSel, onSave = () => {}) {
  const q = k => $(`[data-k="${k}"]`, box);
  const isNum = v => v !== '' && v !== null && v !== undefined && v !== 'auto' && isFinite(+v);
  const axisMode = () => cfg.yMinCustom ? 'custom' : !isNum(cfg.yMin) ? 'auto' : +cfg.yMin === 0 ? '0' : 'custom';
  const syncInputs = () => {
    for (const el of $$('[data-k]', box)) { const k = el.dataset.k;
      if (k === 'yMinMode') { el.value = axisMode(); continue; } if (k === 'yMinNum') { el.hidden = axisMode() !== 'custom'; el.value = axisMode() === 'custom' ? cfg.yMin : ''; continue; }
      if (el.type === 'checkbox') el.checked = !!cfg[k]; else el.value = cfg[k] === 0 ? '0' : (cfg[k] ?? ''); }
    $$('.cb-type', box).forEach(b => b.classList.toggle('on', b.dataset.type === cfg.type)); box.dataset.type = cfg.type; };
  const draw = () => { const wb = getWb(); if (!wb) return; const r = chartOf(wb, 0, cfg); const pv = $('.cb-prev', box); const pw = Math.round(Math.max(300, Math.min(600, pv.clientWidth || 520)));
    pv.innerHTML = chartSVG({ ...r.spec, width: pw, height: Math.round(Math.max(230, Math.min(330, pw * 0.6))) });
    const hs = q('highlight'); if (hs) { const opts = r.data.series.length > 1 ? r.data.series.map(x => x.name) : r.data.cats; const cur = cfg.highlight; hs.innerHTML = '<option value="">None</option>' + opts.map(o => `<option${o === cur ? ' selected' : ''}>${esc(o)}</option>`).join(''); }
    const sc = q('secondary'); if (sc) { const cur = cfg.secondary; sc.innerHTML = '<option value="">None</option>' + r.data.series.map(x => `<option${x.name === cur ? ' selected' : ''}>${esc(x.name)}</option>`).join(''); }
    const tl = $('.cb-tlist', box); if (tl) { const on = new Set(cfg.totals || []); tl.innerHTML = r.data.cats.map(c => `<label><input type="checkbox" data-total="${esc(c)}"${on.has(c) ? ' checked' : ''}> ${esc(c)}</label>`).join('') || '<span class="muted">Choose the data first</span>';
      $$('[data-total]', tl).forEach(el => el.onchange = () => { const set = new Set(cfg.totals || []); el.checked ? set.add(el.dataset.total) : set.delete(el.dataset.total); cfg.totals = r.data.cats.filter(c => set.has(c)); onSave(cfg); draw(); }); } };
  const save = () => { onSave(cfg); draw(); };
  $$('[data-k]', box).forEach(el => el.addEventListener(el.tagName === 'INPUT' && el.type !== 'checkbox' ? 'input' : 'change', () => { const k = el.dataset.k;
    if (k === 'yMinMode') { if (el.value === 'custom') { cfg.yMinCustom = true; if (!isNum(cfg.yMin) || +cfg.yMin === 0) cfg.yMin = ''; } else { cfg.yMinCustom = false; cfg.yMin = el.value === '0' ? 0 : 'auto'; } const n = q('yMinNum'); if (n) { n.hidden = el.value !== 'custom'; n.value = isNum(cfg.yMin) ? cfg.yMin : ''; if (el.value === 'custom') n.focus(); } save(); return; }
    if (k === 'yMinNum') { cfg.yMin = el.value === '' ? '' : +el.value; save(); return; }
    cfg[k] = el.type === 'checkbox' ? el.checked : el.value; save(); }));
  $$('.cb-type', box).forEach(b => b.onclick = () => { cfg.type = b.dataset.type; syncInputs(); save(); });
  $('.cb-sel', box).onclick = () => { let sel = getSel(); if (!sel) return; const wb = getWb(); // a whole Table (or the spill of one formula) is picked by its name, so the chart grows with it, as in Excel
    const tb = wb && (wb.sheets[0].tables || []).find(x => x.ref === sel.replace(/\$/g, '')); if (tb) sel = tb.name;
    else if (wb) { const g = parseRange(sel.replace(/\$/g, '')); const b = g && wb.spillBlock(0, g.r1, g.c1); if (b && !b.blocked && rangeStr(b.g) === rangeStr(g)) sel = addr(b.g.r1, b.g.c1) + '#'; }
    cfg.range = sel; syncInputs(); save(); };
  box._apply = sol => { Object.assign(cfg, BUILDER_DEFAULTS, { totals: [] }, sol || {}); syncInputs(); save(); };
  syncInputs(); return draw;
}
