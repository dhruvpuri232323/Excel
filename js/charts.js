// Charts as SVG, from plain data: used by lesson slides, the lab chart builder and the screenshot tool.
// spec = { type, title, subtitle, cats, series: [{ name, values, kind, axis, color }], labels, legend, grid, yMin, yMax,
//          highlight, palette, fmt, dp, reverse, stacked, markers, junk, xTitle, yTitle, points, trend, holeSize, width, height }
// type: column | bar | line | area | pie | doughnut | scatter | combo   (stacked: 'stacked' | 'percent' for column, bar, area)
export const PALETTES = {
  office: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'],
  cb: ['#0072B2', '#E69F00', '#009E73', '#CC79A7', '#56B4E9', '#D55E00', '#F0E442', '#5B5B5B'],
  mono: ['#2F6EBA', '#7FA6D9', '#1E4C85', '#B9CDEB', '#4F86C6', '#A8C1E4'],
  rainbow: ['#E6194B', '#3CB44B', '#FFE119', '#4363D8', '#F58231', '#911EB4', '#46F0F0', '#F032E6', '#BCF60C', '#FABEBE', '#008080', '#9A6324'],
};
export const ACCENT = '#2F6EBA';
const GREY = 'var(--ch-grey, #BFBFBF)';
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const tw = (s, px = 11) => String(s).length * px * 0.56; // rough text width
const indian = n => { const neg = n < 0; let s = String(Math.round(Math.abs(n))); if (s.length > 3) { const t = s.slice(-3); let h = s.slice(0, -3); const o = []; while (h.length > 2) { o.unshift(h.slice(-2)); h = h.slice(0, -2); } if (h) o.unshift(h); s = o.join(',') + ',' + t; } return (neg ? '−' : '') + s; };
export function fmtNum(v, fmt = 'num', dp) {
  if (v === null || v === undefined || v === '' || Number.isNaN(v)) return '';
  if (typeof v !== 'number') return String(v);
  switch (fmt) {
    case 'inr': return (v < 0 ? '−₹' : '₹') + indian(Math.abs(v)).replace('−', '');
    case 'lakh': { const l = v / 1e5; return '₹' + l.toFixed(dp ?? (Math.abs(l - Math.round(l)) < 1e-9 || Math.abs(l) >= 100 ? 0 : 1)) + ' L'; }
    case 'cr': return '₹' + (v / 1e7).toFixed(dp ?? 1) + ' Cr';
    case 'k': return (v / 1000).toFixed(dp ?? 0) + 'k';
    case 'pct': return (v * 100).toFixed(dp ?? 0) + '%';
    case 'date': { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5); return `${d.getUTCDate()}-${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()]}`; }
    default: return Math.abs(v) >= 1000 ? indian(v) : (+v.toFixed(dp ?? 2)).toLocaleString('en-IN');
  }
}
// Excel's automatic axis minimum: 0, unless every value is above zero and the smallest is more than 5/6 of the largest;
// then the axis starts at the smallest value minus half the spread (rounded down by niceScale), which exaggerates small differences
export function autoMin(lo, hi) { if (lo < 0 || hi <= 0) return Math.min(0, lo); if (lo > hi * 5 / 6) return lo - (hi - lo) / 2; return 0; }
export function axisStart(spec) { const t = spec.type || 'column'; if (['pie', 'doughnut', 'scatter'].includes(t)) return 0; if (spec.yMin !== undefined && spec.yMin !== null) return spec.yMin;
  const ser = (spec.series || []).filter(s => (s.axis || 1) === 1); if (spec.stacked) return 0; let lo = Infinity, hi = -Infinity; ser.forEach(s => s.values.forEach(v => { if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); } })); if (!isFinite(lo)) return 0; return niceScale(autoMin(lo, hi), hi).lo; }
function niceScale(lo, hi, n = 5) {
  if (lo === hi) { hi = lo + 1; lo = lo > 0 ? 0 : lo - 1; }
  const raw = (hi - lo) / n, m = Math.pow(10, Math.floor(Math.log10(raw))); const step = (m === 1 ? [1, 2, 5, 10] : [1, 2, 2.5, 5, 10]).map(k => k * m).find(k => k >= raw * 0.999);
  const a = Math.floor(lo / step + 1e-9) * step, b = Math.ceil(hi / step - 1e-9) * step; const ticks = []; for (let t = a; t <= b + step / 2; t += step) ticks.push(+t.toFixed(10)); return { lo: a, hi: b, ticks };
}
// colours for each series (and each point when one series is highlighted)
function colours(spec) {
  const pal = PALETTES[spec.palette] || PALETTES.office; const H = spec.highlight;
  return { series: (s, i) => s.color || (H && H.series ? (H.series.includes(s.name) ? (spec.accent || ACCENT) : GREY) : (spec.palette === 'grey' ? GREY : pal[i % pal.length])),
    point: (s, i, j) => { if (H && H.cats) return H.cats.includes(spec.cats[j]) ? (spec.accent || ACCENT) : GREY; if (s.points && s.points[j]) return s.points[j]; return null; } };
}
export function chartSVG(spec) {
  const t = spec.type || 'column';
  if (t === 'pie' || t === 'doughnut') return pieSVG(spec);
  if (t === 'scatter') return scatterSVG(spec);
  if (t === 'histogram') return axisSVG(histSpec(spec));
  if (t === 'pareto') return axisSVG(paretoSpec(spec));
  if (t === 'waterfall') return waterfallSVG(spec);
  if (t === 'box') return boxSVG(spec);
  if (t === 'funnel') return funnelSVG(spec);
  if (t === 'treemap') return treemapSVG(spec);
  if (t === 'sunburst') return sunburstSVG(spec);
  if (t === 'stock') return stockSVG(spec);
  if (t === 'map') return tilemapSVG(spec);
  if (t === 'bullet') return bulletSVG(spec);
  return axisSVG(spec);
}
// a long title shrinks a little, then wraps onto two lines, so it never runs off the chart
function titleLayout(spec, W) {
  const t = String(spec.title || ''); if (!t) return { lines: [], size: 14, h: 12 }; const room = W - 20, w = (str, px) => str.length * px * 0.6;
  if (w(t, 14) <= room) return { lines: [t], size: 14, h: 34 };
  if (w(t, 12) <= room) return { lines: [t], size: Math.floor(room / (t.length * 0.6) * 10) / 10, h: 34 };
  const words = t.split(' '); let best = [t, ''], bw = Infinity;
  for (let i = 1; i < words.length; i++) { const a = words.slice(0, i).join(' '), b = words.slice(i).join(' '); const m = Math.max(a.length, b.length); if (m < bw) { bw = m; best = [a, b]; } }
  const size = Math.max(10, Math.min(14, room / (bw * 0.6))); return { lines: best.filter(Boolean), size: Math.floor(size * 10) / 10, h: 34 + Math.round(size * 1.2) };
}
function frame(spec, W, H, inner, legendItems) {
  const TL = titleLayout(spec, W); const head = TL.lines.map((ln, i) => `<text x="${W / 2}" y="${22 + i * Math.round(TL.size * 1.2)}" text-anchor="middle" class="ch-title"${TL.size < 14 ? ` style="font-size:${TL.size}px"` : ''}>${esc(ln)}</text>`).join('') + (spec.subtitle ? `<text x="${W / 2}" y="${spec.title ? TL.h + 6 : 20}" text-anchor="middle" class="ch-sub">${esc(spec.subtitle)}</text>` : '');
  return `<svg class="chart${spec.junk ? ' junk' : ''}" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(spec.alt || spec.title || (t => t)(spec.type + ' chart'))}" xmlns="http://www.w3.org/2000/svg"><rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" class="ch-bg${spec.junk || spec.border ? ' ch-border' : ''}"/>${head}${inner}${legendItems || ''}</svg>`;
}
function legendRows(names, W) { const widths = names.map(n => 26 + tw(n, 11)); const rows = [[]]; let w = 0; widths.forEach((x, i) => { if (w + x > W - 16 && rows[rows.length - 1].length) { rows.push([]); w = 0; } rows[rows.length - 1].push(i); w += x; }); return { rows, widths }; }
function legend(spec, items, W, H, pos) {
  if (!pos || pos === 'none' || !items.length) return '';
  if (pos === 'right') { const x = W - 118; return items.map(([n, c, line], i) => `<g transform="translate(${x},${70 + i * 20})">${line ? `<line x1="0" y1="5" x2="14" y2="5" style="stroke:${c}" stroke-width="3"/>` : `<rect width="11" height="11" style="fill:${c}"/>`}<text x="18" y="10" class="ch-leg">${esc(String(n).slice(0, 16))}</text></g>`).join(''); }
  const L = legendRows(items.map(i => i[0]), W); const y0 = pos === 'top' ? (spec.title ? titleLayout(spec, W).h + 10 : 10) + (spec.subtitle ? 18 : 0) : H - 18 - (L.rows.length - 1) * 18;
  return L.rows.map((row, r) => { const total = row.reduce((a, i) => a + L.widths[i], 0); let x = Math.max(8, (W - total) / 2); const y = y0 + r * 18;
    return row.map(i => { const [n, c, line] = items[i]; const g = `<g transform="translate(${x},${y})">${line ? `<line x1="0" y1="5" x2="14" y2="5" style="stroke:${c}" stroke-width="3"/>` : `<rect width="11" height="11" style="fill:${c}"/>`}<text x="18" y="10" class="ch-leg">${esc(n)}</text></g>`; x += L.widths[i]; return g; }).join(''); }).join('');
}
function axisSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const t = spec.type || 'column'; const horiz = t === 'bar';
  let cats = (spec.cats || []).map(String); let series = (spec.series || []).map(s => ({ ...s, values: s.values.map(v => typeof v === 'number' ? v : (v === '' || v === null || v === undefined ? null : +v || null)) }));
  if (!series.length || !cats.length) return frame(spec, W, H, `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" class="ch-empty">${esc(spec.empty || 'Pick some data to chart')}</text>`);
  const kindOf = s => t === 'combo' ? (s.kind || 'column') : (t === 'bar' ? 'bar' : t === 'line' ? 'line' : t === 'area' ? 'area' : 'column');
  const C = colours(spec); const stacked = spec.stacked; const pct = stacked === 'percent';
  const legPos = spec.legend ?? (series.length > 1 ? 'bottom' : 'none');
  const prim = series.filter(s => (s.axis || 1) === 1), sec = series.filter(s => s.axis === 2);
  const range = list => { let lo = Infinity, hi = -Infinity; if (stacked && list.length) { cats.forEach((_, j) => { let p = 0, n = 0; list.forEach(s => { const v = s.values[j] || 0; if (v >= 0) p += v; else n += v; }); if (pct) { p = p ? 1 : 0; n = n ? -1 : 0; } lo = Math.min(lo, n); hi = Math.max(hi, p); }); } else list.forEach(s => s.values.forEach(v => { if (v !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); } })); if (!isFinite(lo)) { lo = 0; hi = 1; } return [lo, hi]; };
  const scaleFor = (list, min, max) => { let [lo, hi] = range(list); if (min === undefined || min === null) lo = autoMin(lo, hi); else lo = min; if (max !== undefined && max !== null) hi = max; const s = niceScale(lo, hi); if (min !== undefined && min !== null) { const st = s.ticks.length > 1 ? s.ticks[1] - s.ticks[0] : 1; s.lo = min; s.ticks = s.ticks.filter(x => x >= min - 1e-9); if (s.ticks[0] !== min) { if (s.ticks.length && s.ticks[0] - min < st * 0.6) s.ticks.shift(); s.ticks.unshift(min); } } if (max !== undefined && max !== null) { s.hi = max; s.ticks = s.ticks.filter(x => x <= max + 1e-9); } return s; };
  const S1 = scaleFor(prim.length ? prim : series, spec.yMin, spec.yMax), S2 = sec.length ? scaleFor(sec, spec.y2Min, spec.y2Max) : null;
  const fmt1 = v => fmtNum(v, pct ? 'pct' : (spec.fmt || 'num'), spec.dp), fmt2 = v => fmtNum(v, spec.fmt2 || spec.fmt || 'num', spec.dp2);
  const legX = (legPos === 'top' || legPos === 'bottom') && spec.labels !== 'end' ? (legendRows(series.map(x => x.name), W).rows.length - 1) * 18 : 0; spec = { ...spec, _legX: legX };
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + (legPos === 'top' ? 24 + legX : 0) + 8;
  const yLab = Math.max(...S1.ticks.map(v => tw(fmt1(v), 10.5))) + 12 + (spec.yTitle ? 16 : 0);
  const rightPad = (legPos === 'right' ? 124 : 14) + (S2 ? Math.max(...S2.ticks.map(v => tw(fmt2(v), 10.5))) + 10 : 0) + (spec.labels === 'end' ? Math.max(...series.map(s => tw(s.name, 11))) + 6 : 0);
  if (horiz) return barH(spec, W, H, cats, series, C, S1, fmt1, top, legPos);
  const maxCat = Math.max(...cats.map(c => tw(c, 10.5))); const left = yLab; const plotW = W - left - rightPad; const band = plotW / cats.length; const rot = maxCat > band - 4;
  const catH = rot ? Math.min(70, maxCat * 0.7 + 10) : 20; const bottom = (legPos === 'bottom' ? 30 + (spec._legX || 0) : 8) + catH + (spec.xTitle ? 16 : 0);
  const plotH = H - top - bottom; const Y = (v, S) => top + plotH - (v - S.lo) / ((S.hi - S.lo) || 1) * plotH;
  let g = '';
  if (spec.grid !== false) g += S1.ticks.map(v => `<line x1="${left}" y1="${Y(v, S1).toFixed(1)}" x2="${left + plotW}" y2="${Y(v, S1).toFixed(1)}" class="ch-grid"/>`).join('');
  g += S1.ticks.map(v => `<text x="${left - 6}" y="${(Y(v, S1) + 3.5).toFixed(1)}" text-anchor="end" class="ch-tick">${esc(fmt1(v))}</text>`).join('');
  if (S2) g += S2.ticks.map(v => `<text x="${left + plotW + 6}" y="${(Y(v, S2) + 3.5).toFixed(1)}" class="ch-tick">${esc(fmt2(v))}</text>`).join('');
  const base = Y(Math.max(S1.lo, Math.min(0, S1.hi)), S1);
  g += `<line x1="${left}" y1="${base.toFixed(1)}" x2="${left + plotW}" y2="${base.toFixed(1)}" class="ch-axis"/>`;
  cats.forEach((c, j) => { const x = left + band * (j + .5); g += rot ? `<text transform="translate(${x.toFixed(1)},${top + plotH + 12}) rotate(-40)" text-anchor="end" class="ch-cat">${esc(c)}</text>` : `<text x="${x.toFixed(1)}" y="${top + plotH + 15}" text-anchor="middle" class="ch-cat">${esc(c)}</text>`; });
  if (spec.yTitle) g += `<text transform="translate(12,${top + plotH / 2}) rotate(-90)" text-anchor="middle" class="ch-axt">${esc(spec.yTitle)}</text>`;
  if (spec.xTitle) g += `<text x="${left + plotW / 2}" y="${H - (legPos === 'bottom' ? 30 + (spec._legX || 0) : 6)}" text-anchor="middle" class="ch-axt">${esc(spec.xTitle)}</text>`;
  // columns
  const cols = series.filter(s => kindOf(s) === 'column'); const gw = band * (spec.gap === undefined ? 0.66 : 1 - spec.gap);
  const stackPos = cats.map(() => 0), stackNeg = cats.map(() => 0); const tot = cats.map((_, j) => cols.reduce((a, s) => a + Math.abs(s.values[j] || 0), 0));
  cols.forEach((s, i) => { const S = s.axis === 2 ? S2 : S1;
    s.values.forEach((v0, j) => { if (v0 === null) return; let v = pct ? (tot[j] ? v0 / tot[j] : 0) : v0; let y0, y1, x, w;
      if (stacked) { w = gw; x = left + band * j + (band - gw) / 2; if (v >= 0) { y0 = Y(stackPos[j], S); stackPos[j] += v; y1 = Y(stackPos[j], S); } else { y0 = Y(stackNeg[j], S); stackNeg[j] += v; y1 = Y(stackNeg[j], S); } }
      else { w = gw / cols.length; x = left + band * j + (band - gw) / 2 + w * i; y0 = Y(Math.max(S.lo, 0), S); y1 = Y(v, S); }
      if (s.color === 'none') return; const fill = (spec.invert && v0 < 0 && !C.point(s, i, j)) ? (spec.invertColor || '#ED7D31') : (C.point(s, i, j) || C.series(s, i)); g += `<rect x="${x.toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${Math.max(w - (stacked ? 0 : 1), 1).toFixed(1)}" height="${Math.abs(y1 - y0).toFixed(1)}" style="fill:${fill}"/>`;
      if (spec.labels && spec.labels !== 'end') g += stacked ? (Math.abs(y1 - y0) > 13 ? `<text x="${(x + w / 2).toFixed(1)}" y="${((y0 + y1) / 2 + 4).toFixed(1)}" text-anchor="middle" class="ch-lab in">${esc(pct ? fmtNum(v, 'pct') : (s.axis === 2 ? fmt2 : fmt1)(v0))}</text>` : '') : `<text x="${(x + w / 2).toFixed(1)}" y="${(v >= 0 ? y1 - 5 : y1 + 13).toFixed(1)}" text-anchor="middle" class="ch-lab">${esc((s.axis === 2 ? fmt2 : fmt1)(v0))}</text>`; }); });
  // areas
  const areas = series.filter(s => kindOf(s) === 'area'); const acc = cats.map(() => 0);
  areas.forEach((s, i) => { const S = s.axis === 2 ? S2 : S1; const lo = acc.slice(); const hi = s.values.map((v, j) => { const vv = pct ? (v || 0) / (series.reduce((a, q) => a + Math.abs(q.values[j] || 0), 0) || 1) : (v || 0); return stacked ? lo[j] + vv : vv; }); if (stacked) hi.forEach((v, j) => acc[j] = v);
    const pts = hi.map((v, j) => `${(left + band * (j + .5)).toFixed(1)},${Y(v, S).toFixed(1)}`); const bl = (stacked ? lo : lo.map(() => Math.max(S.lo, 0))).map((v, j) => `${(left + band * (j + .5)).toFixed(1)},${Y(v, S).toFixed(1)}`).reverse();
    const col = C.series(s, i); g += `<polygon points="${[...pts, ...bl].join(' ')}" style="fill:${col}" opacity="${stacked ? 0.9 : 0.55}"/><polyline points="${pts.join(' ')}" fill="none" style="stroke:${col}" stroke-width="1.5"/>`; });
  // lines
  const lines = series.filter(s => kindOf(s) === 'line');
  lines.forEach((s, i0) => { const i = series.indexOf(s); const S = s.axis === 2 ? S2 : S1; const col = C.series(s, i); let d = ''; let pen = false; const pts = [];
    s.values.forEach((v, j) => { if (v === null) { pen = false; return; } const x = left + band * (j + .5), y = Y(v, S); d += `${pen ? 'L' : 'M'}${x.toFixed(1)},${y.toFixed(1)} `; pen = true; pts.push([x, y, v, j]); });
    const hl = spec.highlight && spec.highlight.series && spec.highlight.series.includes(s.name);
    g += `<path d="${d}" fill="none" style="stroke:${col}" stroke-width="${hl ? 3.2 : 2.4}" stroke-linejoin="round"/>`;
    if (spec.markers !== false && (spec.markers || pts.length <= 14)) g += pts.map(([x, y, v, j]) => `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3.2" style="fill:${C.point(s, i, j) || col}"/>`).join('');
    if (spec.labels === true) g += pts.map(([x, y, v]) => `<text x="${x.toFixed(1)}" y="${(y - 8).toFixed(1)}" text-anchor="middle" class="ch-lab">${esc((s.axis === 2 ? fmt2 : fmt1)(v))}</text>`).join('');
    if (spec.labels === 'end' && pts.length) { const [x, y] = pts[pts.length - 1]; g += `<text x="${(x + 6).toFixed(1)}" y="${(y + 4).toFixed(1)}" class="ch-lab end" style="fill:${col === GREY ? 'var(--ch-ink2,#6b6b6b)' : col}">${esc(s.name)}</text>`; } });
  if (spec.annotate) for (const a of [].concat(spec.annotate)) { const j = cats.indexOf(String(a.cat)); if (j < 0) continue; const s = series[a.series || 0]; const v = s.values[j]; const x = left + band * (j + .5), y = Y(v, s.axis === 2 ? S2 : S1); const up = y - 34 > top; g += up ? `<line x1="${x}" y1="${y - 6}" x2="${x}" y2="${y - 26}" class="ch-note-l"/><text x="${x}" y="${y - 30}" text-anchor="middle" class="ch-note">${esc(a.text)}</text>` : `<line x1="${x}" y1="${y + 6}" x2="${x}" y2="${y + 22}" class="ch-note-l"/><text x="${x}" y="${y + 34}" text-anchor="middle" class="ch-note">${esc(a.text)}</text>`; }
  const items = spec.labels === 'end' ? [] : series.map((s, i) => [s.name, C.series(s, i), kindOf(s) === 'line']).filter(it => it[1] !== 'none');
  return frame(spec, W, H, g, legend(spec, items, W, H, legPos));
}
function barH(spec, W, H, cats, series, C, S, fmt, top, legPos) {
  const order = cats.map((_, j) => j); if (!spec.reverse) order.reverse(); // like Excel: the first category sits at the bottom unless the axis is reversed
  const left = Math.min(170, Math.max(...cats.map(c => tw(c, 10.5))) + 14); const rightPad = (legPos === 'right' ? 124 : 18) + (spec.labels ? 48 : 0);
  const bottom = (legPos === 'bottom' ? 30 + (spec._legX || 0) : 8) + 20 + (spec.xTitle ? 16 : 0); const plotW = W - left - rightPad, plotH = H - top - bottom; const band = plotH / cats.length;
  const X = v => left + (v - S.lo) / ((S.hi - S.lo) || 1) * plotW; const pct = spec.stacked === 'percent'; const gw = band * (spec.gap === undefined ? 0.66 : 1 - spec.gap);
  let g = '';
  if (spec.grid !== false) g += S.ticks.map(v => `<line x1="${X(v).toFixed(1)}" y1="${top}" x2="${X(v).toFixed(1)}" y2="${top + plotH}" class="ch-grid"/>`).join('');
  g += S.ticks.map(v => `<text x="${X(v).toFixed(1)}" y="${top + plotH + 15}" text-anchor="middle" class="ch-tick">${esc(fmt(v))}</text>`).join('');
  g += `<line x1="${X(Math.max(S.lo, 0)).toFixed(1)}" y1="${top}" x2="${X(Math.max(S.lo, 0)).toFixed(1)}" y2="${top + plotH}" class="ch-axis"/>`;
  const tot = cats.map((_, j) => series.reduce((a, s) => a + Math.abs(s.values[j] || 0), 0)); const pos = cats.map(() => 0);
  order.forEach((j, k) => { const y = top + band * k; g += `<text x="${left - 6}" y="${(y + band / 2 + 4).toFixed(1)}" text-anchor="end" class="ch-cat">${esc(cats[j])}</text>`;
    series.forEach((s, i) => { const v0 = s.values[j]; if (v0 === null) return; const v = pct ? (tot[j] ? v0 / tot[j] : 0) : v0; let x0, x1, yy, h;
      if (spec.stacked) { h = gw; yy = y + (band - gw) / 2; x0 = X(pos[j]); pos[j] += v; x1 = X(pos[j]); } else { h = gw / series.length; yy = y + (band - gw) / 2 + h * i; x0 = X(Math.max(S.lo, 0)); x1 = X(v); }
      if (s.color === 'none') return; const fill = (spec.invert && v0 < 0 && !C.point(s, i, j)) ? (spec.invertColor || '#ED7D31') : (C.point(s, i, j) || C.series(s, i)); g += `<rect x="${Math.min(x0, x1).toFixed(1)}" y="${yy.toFixed(1)}" width="${Math.abs(x1 - x0).toFixed(1)}" height="${Math.max(h - (spec.stacked ? 0 : 1), 1).toFixed(1)}" style="fill:${fill}"/>`;
      if (spec.labels) g += spec.stacked ? (Math.abs(x1 - x0) > 34 ? `<text x="${((x0 + x1) / 2).toFixed(1)}" y="${(yy + h / 2 + 4).toFixed(1)}" text-anchor="middle" class="ch-lab in">${esc(pct ? fmtNum(v, 'pct') : fmt(v0))}</text>` : '') : `<text x="${(x1 + 5).toFixed(1)}" y="${(yy + h / 2 + 4).toFixed(1)}" class="ch-lab">${esc(fmt(v0))}</text>`; }); });
  const items = series.map((s, i) => [s.name, C.series(s, i), false]).filter(it => it[1] !== 'none');
  return frame(spec, W, H, g, legend(spec, items, W, H, legPos));
}
function pieSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const s = (spec.series || [])[0]; const cats = (spec.cats || []).map(String);
  if (!s || !cats.length) return frame(spec, W, H, `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" class="ch-empty">${esc(spec.empty || 'Pick some data to chart')}</text>`);
  const vals = s.values.map(v => Math.max(0, +v || 0)); const tot = vals.reduce((a, b) => a + b, 0) || 1; const legPos = spec.legend ?? 'bottom';
  const top = (spec.title ? titleLayout(spec, W).h : 10) + (spec.subtitle ? 18 : 0); const bottom = legPos === 'bottom' ? 30 + (legendRows(cats, W).rows.length - 1) * 18 : legPos === 'top' ? 10 : 10; const right = legPos === 'right' ? 124 : 0;
  const labTxt = (v, j) => spec.labels === 'value' ? fmtNum(v, spec.fmt || 'num', spec.dp) : spec.labels === 'category' ? cats[j] : spec.labels === 'catpct' ? `${cats[j]} ${Math.round(v / tot * 100)}%` : `${Math.round(v / tot * 100)}%`;
  const labW = spec.labels ? Math.max(...vals.map((v, j) => tw(labTxt(v, j), 11))) + 16 : 16;
  const r = Math.max(30, Math.min((W - right) / 2 - Math.min(labW, (W - right) / 2 - 50), (H - top - bottom) / 2 - (spec.labels ? 14 : 6))); const cx = (W - right) / 2, cy = top + (H - top - bottom) / 2;
  const pal = PALETTES[spec.palette] || PALETTES.office; const HL = spec.highlight && spec.highlight.cats; const col = j => HL ? (HL.includes(cats[j]) ? (spec.accent || ACCENT) : GREY) : (s.points && s.points[j]) || pal[j % pal.length];
  let a = -Math.PI / 2 + (spec.firstAngle || 0) * Math.PI / 180; let g = ''; const hole = spec.type === 'doughnut' ? (spec.holeSize ?? 0.55) : 0;
  vals.forEach((v, j) => { const da = v / tot * Math.PI * 2; if (!da) return; const a2 = a + da; const big = da > Math.PI ? 1 : 0; const p = (ang, rr) => [cx + rr * Math.cos(ang), cy + rr * Math.sin(ang)];
    const [x1, y1] = p(a, r), [x2, y2] = p(a2, r); let d;
    if (da >= Math.PI * 2 - 1e-6) d = hole ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0M${cx - r * hole},${cy}a${r * hole},${r * hole} 0 1,1 ${2 * r * hole},0a${r * hole},${r * hole} 0 1,1 ${-2 * r * hole},0` : `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0`;
    else if (hole) { const [x3, y3] = p(a2, r * hole), [x4, y4] = p(a, r * hole); d = `M${x1.toFixed(1)},${y1.toFixed(1)}A${r},${r} 0 ${big},1 ${x2.toFixed(1)},${y2.toFixed(1)}L${x3.toFixed(1)},${y3.toFixed(1)}A${r * hole},${r * hole} 0 ${big},0 ${x4.toFixed(1)},${y4.toFixed(1)}Z`; }
    else d = `M${cx},${cy}L${x1.toFixed(1)},${y1.toFixed(1)}A${r},${r} 0 ${big},1 ${x2.toFixed(1)},${y2.toFixed(1)}Z`;
    g += `<path d="${d}" style="fill:${col(j)}" class="ch-slice"/>`;
    if (spec.labels) { const mid = a + da / 2; const [lx, ly] = p(mid, r + 12); const right2 = Math.cos(mid) >= 0; const txt = labTxt(v, j); const lw = tw(txt, 11); const lx2 = right2 ? Math.min(lx, W - right - 4 - lw) : Math.max(lx, lw + 4);
      g += `<text x="${lx2.toFixed(1)}" y="${(ly + 4).toFixed(1)}" text-anchor="${right2 ? 'start' : 'end'}" class="ch-lab">${esc(txt)}</text>`; }
    a = a2; });
  if (hole && spec.centre) g += `<text x="${cx}" y="${cy + 6}" text-anchor="middle" class="ch-centre">${esc(spec.centre)}</text>`;
  const items = cats.map((c, j) => [c, col(j), false]);
  return frame(spec, W, H, g, legend(spec, items, W, H, legPos));
}
function scatterSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const P = spec.points || ((spec.series || [])[0] ? spec.series[0].values.map((y, j) => [+spec.cats[j], y]) : []);
  const pts = P.filter(p => typeof p[0] === 'number' && !isNaN(p[0]) && typeof p[1] === 'number');
  if (!pts.length) return frame(spec, W, H, `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" class="ch-empty">${esc(spec.empty || 'Pick two columns of numbers')}</text>`);
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]); const SX = niceScale(spec.xMin ?? Math.min(...xs), spec.xMax ?? Math.max(...xs)), SY = niceScale(spec.yMin ?? Math.min(0, ...ys), spec.yMax ?? Math.max(...ys));
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + 6; const left = Math.max(...SY.ticks.map(v => tw(fmtNum(v, spec.fmt || 'num'), 10.5))) + 14 + (spec.yTitle ? 16 : 0); const bottom = 26 + (spec.xTitle ? 18 : 0); const plotW = W - left - 16, plotH = H - top - bottom;
  const X = v => left + (v - SX.lo) / ((SX.hi - SX.lo) || 1) * plotW, Y = v => top + plotH - (v - SY.lo) / ((SY.hi - SY.lo) || 1) * plotH; let g = '';
  if (spec.grid !== false) g += SY.ticks.map(v => `<line x1="${left}" y1="${Y(v).toFixed(1)}" x2="${left + plotW}" y2="${Y(v).toFixed(1)}" class="ch-grid"/>`).join('');
  g += SY.ticks.map(v => `<text x="${left - 6}" y="${(Y(v) + 3.5).toFixed(1)}" text-anchor="end" class="ch-tick">${esc(fmtNum(v, spec.fmt || 'num'))}</text>`).join('') + SX.ticks.map(v => `<text x="${X(v).toFixed(1)}" y="${top + plotH + 15}" text-anchor="middle" class="ch-tick">${esc(fmtNum(v, spec.fmtx || 'num'))}</text>`).join('');
  g += `<line x1="${left}" y1="${top + plotH}" x2="${left + plotW}" y2="${top + plotH}" class="ch-axis"/><line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" class="ch-axis"/>`;
  if (spec.trend && pts.length > 1) { const n = pts.length, mx = xs.reduce((a, b) => a + b) / n, my = ys.reduce((a, b) => a + b) / n; let sxy = 0, sxx = 0; pts.forEach(([x, y]) => { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; }); const b = sxx ? sxy / sxx : 0, a = my - b * mx; const x1 = Math.min(...xs), x2 = Math.max(...xs); g += `<line x1="${X(x1).toFixed(1)}" y1="${Y(a + b * x1).toFixed(1)}" x2="${X(x2).toFixed(1)}" y2="${Y(a + b * x2).toFixed(1)}" class="ch-trend"/>`; }
  const pal = PALETTES[spec.palette] || PALETTES.office; g += pts.map(p => `<circle cx="${X(p[0]).toFixed(1)}" cy="${Y(p[1]).toFixed(1)}" r="${p[3] ? 5.5 : 4}" style="fill:${p[3] === 'hl' ? (spec.accent || ACCENT) : p[3] === 'grey' ? GREY : pal[0]}" opacity=".85"/>${p[2] ? `<text x="${(X(p[0]) + 7).toFixed(1)}" y="${(Y(p[1]) - 6).toFixed(1)}" class="ch-lab">${esc(p[2])}</text>` : ''}`).join('');
  if (spec.yTitle) g += `<text transform="translate(12,${top + plotH / 2}) rotate(-90)" text-anchor="middle" class="ch-axt">${esc(spec.yTitle)}</text>`;
  if (spec.xTitle) g += `<text x="${left + plotW / 2}" y="${H - 6}" text-anchor="middle" class="ch-axt">${esc(spec.xTitle)}</text>`;
  return frame(spec, W, H, g, '');
}
// Read a chart's data from a block of cells, as Excel does: text in the first row names the series, text in the first
// column names the categories; series run down the columns unless byRow (Switch Row/Column).
export function dataFromBlock(vals, byRow = false) {
  if (!vals || !vals.length || !vals[0].length) return { cats: [], series: [] };
  let V = vals.map(r => r.slice()); if (byRow) V = V[0].map((_, j) => V.map(r => r[j]));
  const isText = v => typeof v === 'string' && v.trim() !== '' && isNaN(+v);
  const c0 = V.slice(1).some(r => isText(r[0])) || (V.length === 1 && isText(V[0][0])) ? 1 : 0;
  const head = V.length > 1 && V[0].slice(c0).some(isText); const rows = head ? V.slice(1) : V;
  const cats = c0 ? rows.map(r => String(r[0] ?? '')) : rows.map((_, i) => String(i + 1));
  const series = []; for (let j = c0; j < V[0].length; j++) series.push({ name: head ? String(V[0][j] ?? `Series${j - c0 + 1}`) : `Series${j - c0 + 1}`, values: rows.map(r => typeof r[j] === 'number' ? r[j] : (r[j] === null || r[j] === undefined || r[j] === '' ? null : (isNaN(+r[j]) ? null : +r[j]))) });
  return { cats, series, hasCats: !!c0 };
}
// The lab chart builder's settings → a chart. block(range) returns the cells' values (rows of values) or null.
export const BUILDER_DEFAULTS = { range: '', type: 'column', title: '', labels: false, endLabels: false, legend: 'auto', grid: true, yMin: 'auto', highlight: '', byRow: false, reverse: false, secondary: '', stacked: '', bins: 'auto', binVal: '', underflow: '', overflow: '', totals: [], hideFirst: false, invert: false, yMinCustom: false };
const num = v => v === '' || v === null || v === undefined || !isFinite(+v) ? null : +v;
// the value of a cell for a title linked to it (=B20)
export function titleText(cfg, block) { const t = String(cfg.title || ''); if (!/^=\s*'?[A-Za-z0-9 _]*'?!?\$?[A-Za-z]+\$?\d+\s*$/.test(t)) return t; try { const v = block(t.replace(/^=\s*/, '').replace(/\$/g, '')); const x = v && v[0] && v[0][0]; return x === null || x === undefined ? '' : String(x); } catch (e) { return ''; } }
export function chartFromConfig(cfg0, block, fmt) {
  const cfg = { ...BUILDER_DEFAULTS, ...(cfg0 || {}) }; const vals = cfg.range ? block(cfg.range) : null; const type = cfg.type || 'column';
  const ds = vals ? dataFromBlock(vals, cfg.byRow) : { cats: [], series: [] };
  const spec = { type, title: titleText(cfg, block), cats: ds.cats, series: ds.series, labels: cfg.labels ? (type === 'pie' || type === 'doughnut' ? 'catpct' : true) : false, grid: cfg.grid !== false, reverse: !!cfg.reverse, fmt: fmt || 'num',
    empty: vals ? 'These cells have no numbers to chart' : 'Choose the data: type a range or use your selection' };
  if (cfg.legend !== 'auto') spec.legend = cfg.legend;
  if (num(cfg.yMin) !== null && !['pie', 'doughnut'].includes(type)) spec.yMin = num(cfg.yMin);
  if (cfg.endLabels && (type === 'line' || type === 'area')) spec.labels = 'end';
  if (cfg.palette) spec.palette = cfg.palette;
  if (cfg.stacked) spec.stacked = cfg.stacked;
  if (cfg.highlight) spec.highlight = ds.series.some(s => s.name === cfg.highlight) && ds.series.length > 1 ? { series: [cfg.highlight] } : { cats: [cfg.highlight] };
  if (cfg.invert) spec.invert = true;
  if (cfg.hideFirst && spec.series.length > 1) spec.series = spec.series.map((s, i) => i ? s : { ...s, color: 'none' });
  if (type === 'combo') spec.series = ds.series.map(s => s.name === cfg.secondary ? { ...s, kind: 'line', axis: 2 } : { ...s, kind: 'column' });
  if (type === 'scatter' && ds.series.length >= 2) { spec.points = ds.series[0].values.map((x, j) => [x, ds.series[1].values[j]]); spec.xTitle = ds.series[0].name; spec.yTitle = ds.series[1].name; }
  else if (type === 'scatter' && ds.series.length === 1) spec.points = ds.series[0].values.map((y, j) => [+ds.cats[j], y]);
  if (type === 'histogram') { const s0 = ds.series[ds.series.length - 1]; spec.values = s0 ? s0.values : []; spec.series = s0 ? [{ name: 'Count', values: [] }] : []; spec.fmt = fmt || 'num';
    if (cfg.bins === 'width' && num(cfg.binVal) > 0) spec.binWidth = num(cfg.binVal); if (cfg.bins === 'count' && num(cfg.binVal) > 0) spec.binCount = Math.round(num(cfg.binVal));
    if (num(cfg.underflow) !== null) spec.underflow = num(cfg.underflow); if (num(cfg.overflow) !== null) spec.overflow = num(cfg.overflow); }
  if (type === 'box') spec.groups = ds.hasCats && ds.series.length === 1 ? [...new Set(ds.cats)].map(c => ({ name: c, values: ds.series[0].values.filter((v, j) => ds.cats[j] === c) })) : ds.series.map(s => ({ name: s.name, values: s.values }));
  if (type === 'waterfall') spec.totals = (cfg.totals || []).slice();
  if (type === 'map' && vals) { const rows = vals.filter(r => r.some(v => v !== null && v !== '')); const body = rows.length > 1 && rows[0].every(v => typeof v === 'string') ? rows.slice(1) : rows; const nameCol = r => { let k = -1; r.forEach((v, i) => { if (typeof v === 'string' && isNaN(+v)) k = i; }); return k; };
    const pairs = body.map(r => { const k = nameCol(r); const v = [...r].reverse().find(x => typeof x === 'number'); return k >= 0 && typeof v === 'number' ? [String(r[k]), v] : null; }).filter(Boolean); spec.cats = pairs.map(x => x[0]); spec.series = [{ name: 'Value', values: pairs.map(x => x[1]) }]; }
  if (type === 'treemap' || type === 'sunburst' || type === 'stock') { const rows = (vals || []).filter(r => r.some(v => v !== null && v !== '')); const head = rows.length > 1 && typeof rows[0][rows[0].length - 1] === 'string' && isNaN(+rows[0][rows[0].length - 1]);
    const body = head ? rows.slice(1) : rows; if (type === 'stock') spec.ohlc = body.map(r => r.slice(0, 5).map((v, k) => k && typeof v !== 'number' && v !== null && v !== '' && !isNaN(+v) ? +v : v)); else spec.rows = body.map(r => r.map((v, k) => k === r.length - 1 && typeof v !== 'number' && !isNaN(+v) && v !== '' ? +v : v)); }
  return { spec, data: ds, cfg };
}

// ---------------------------------------------------------------- Excel 2016 chart types (Module 3)
const empty = (spec, W, H, msg) => frame(spec, W, H, `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" class="ch-empty">${esc(spec.empty || msg)}</text>`);
const nums = a => (a || []).filter(v => typeof v === 'number' && isFinite(v));
// Histogram bins as Excel makes them: bin width (or count) from the smallest value; automatic width by Scott's rule
// (3.49 × standard deviation × n^(−1/3)); the first bin is [a, b], the rest (a, b]; optional underflow (≤u) and overflow (>o) bins.
export function histBins(values, o = {}) {
  const v = nums(values).slice().sort((a, b) => a - b); const n = v.length; if (!n) return { bins: [], width: 0 };
  const min = v[0], max = v[n - 1]; const mean = v.reduce((a, b) => a + b, 0) / n; const sd = n > 1 ? Math.sqrt(v.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1)) : 1;
  const lo0 = o.underflow ?? min, hi0 = o.overflow ?? max; let w = o.binWidth ? +o.binWidth : o.binCount ? (hi0 - lo0) / +o.binCount : 3.49 * sd * Math.pow(n, -1 / 3); if (!(w > 0)) w = 1;
  const bins = []; const f = x => fmtNum(x, o.fmt || 'num', o.dp ?? (Math.abs(w) >= 10 ? 0 : 1));
  if (o.underflow !== undefined && o.underflow !== null) bins.push({ label: `≤${f(o.underflow)}`, test: x => x <= o.underflow });
  let a = lo0, first = o.underflow === undefined || o.underflow === null; let guard = 0;
  while (guard++ < 400) { const b = a + w; const a1 = a, fst = first; bins.push({ label: `${fst ? '[' : '('}${f(a1)}, ${f(b)}]`, test: x => (fst ? x >= a1 : x > a1) && x <= b + 1e-9 * Math.abs(b) }); first = false; a = b; if (a >= hi0 - 1e-9 * Math.abs(hi0 || 1)) break; }
  if (o.overflow !== undefined && o.overflow !== null) bins.push({ label: `>${f(o.overflow)}`, test: x => x > o.overflow });
  const used = new Set(); bins.forEach(bn => { bn.count = v.filter((x, i) => !used.has(i) && bn.test(x) && used.add(i)).length; });
  return { bins: bins.map(b => ({ label: b.label, count: b.count })), width: w };
}
function histSpec(spec) {
  const vals = spec.values || ((spec.series || [])[0] || {}).values || []; const H = histBins(vals, spec);
  return { ...spec, type: 'column', cats: H.bins.map(b => b.label), series: [{ name: (spec.series && spec.series[0] && spec.series[0].name) || 'Count', values: H.bins.map(b => b.count) }], gap: spec.gap ?? 0.04, fmt: 'num', dp: 0, legend: spec.legend ?? 'none', empty: 'Pick one column of numbers' };
}
function paretoSpec(spec) {
  const s = (spec.series || [])[0]; if (!s) return { ...spec, type: 'column' };
  const rows = (spec.cats || []).map((c, j) => [c, s.values[j]]).filter(r => typeof r[1] === 'number').sort((a, b) => b[1] - a[1]); const tot = rows.reduce((a, r) => a + r[1], 0) || 1; let run = 0;
  return { ...spec, type: 'combo', cats: rows.map(r => r[0]), series: [{ name: s.name, values: rows.map(r => r[1]), kind: 'column' }, { name: 'Cumulative %', values: rows.map(r => (run += r[1]) / tot), kind: 'line', axis: 2 }], y2Min: 0, y2Max: 1, fmt2: 'pct', gap: spec.gap ?? 0.04, markers: true };
}
// a value axis + category bands, shared by the hand-drawn types below
function bands(spec, W, H, cats, lo, hi, fmt, legH = 0) {
  const S = niceScale(lo, hi); const f = v => fmtNum(v, fmt || spec.fmt || 'num', spec.dp);
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + legH + 8; const left = Math.max(...S.ticks.map(v => tw(f(v), 10.5))) + 12 + (spec.yTitle ? 16 : 0); const plotW = W - left - 14;
  const band = plotW / Math.max(1, cats.length); const maxCat = Math.max(0, ...cats.map(c => tw(c, 10.5))); const rot = maxCat > band - 4; const catH = rot ? Math.min(70, maxCat * 0.7 + 10) : 20;
  const plotH = H - top - 8 - catH - (spec.xTitle ? 16 : 0); const Y = v => top + plotH - (v - S.lo) / ((S.hi - S.lo) || 1) * plotH; let g = '';
  if (spec.grid !== false) g += S.ticks.map(v => `<line x1="${left}" y1="${Y(v).toFixed(1)}" x2="${left + plotW}" y2="${Y(v).toFixed(1)}" class="ch-grid"/>`).join('');
  g += S.ticks.map(v => `<text x="${left - 6}" y="${(Y(v) + 3.5).toFixed(1)}" text-anchor="end" class="ch-tick">${esc(f(v))}</text>`).join('');
  g += `<line x1="${left}" y1="${Y(Math.max(S.lo, Math.min(0, S.hi))).toFixed(1)}" x2="${left + plotW}" y2="${Y(Math.max(S.lo, Math.min(0, S.hi))).toFixed(1)}" class="ch-axis"/>`;
  cats.forEach((c, j) => { const x = left + band * (j + .5); g += rot ? `<text transform="translate(${x.toFixed(1)},${top + plotH + 12}) rotate(-40)" text-anchor="end" class="ch-cat">${esc(c)}</text>` : `<text x="${x.toFixed(1)}" y="${top + plotH + 15}" text-anchor="middle" class="ch-cat">${esc(c)}</text>`; });
  if (spec.yTitle) g += `<text transform="translate(12,${top + plotH / 2}) rotate(-90)" text-anchor="middle" class="ch-axt">${esc(spec.yTitle)}</text>`;
  if (spec.xTitle) g += `<text x="${left + plotW / 2}" y="${H - 6}" text-anchor="middle" class="ch-axt">${esc(spec.xTitle)}</text>`;
  return { S, Y, left, top, plotW, plotH, band, g, f };
}
const WF = { up: '#4472C4', down: '#ED7D31', total: '#A5A5A5' };
// Waterfall: each bar starts where the last one ended; points set as totals stand on the axis.
export function waterfallLevels(values, totals) { let level = 0; return values.map((v, j) => { v = +v || 0; if (totals.includes(j)) { level = v; return { from: 0, to: v, kind: 'total', v }; } const from = level; level += v; return { from, to: level, kind: v >= 0 ? 'up' : 'down', v }; }); }
function waterfallSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const s = (spec.series || [])[0]; const cats = (spec.cats || []).map(String); if (!s || !cats.length) return empty(spec, W, H, 'Pick categories and one column of amounts');
  const tot = (spec.totals || []).map(t => typeof t === 'number' ? t : cats.indexOf(String(t))).filter(j => j >= 0); const L = waterfallLevels(s.values, tot);
  const all = L.flatMap(l => [l.from, l.to]); const lo = spec.yMin ?? Math.min(0, ...all), hi = Math.max(0, ...all); const legH = spec.legend === 'none' ? 0 : 22;
  const B = bands(spec, W, H, cats, lo, hi, spec.fmt, legH); let g = B.g; const gw = B.band * (1 - (spec.gap ?? 0.35));
  L.forEach((l, j) => { const x = B.left + B.band * j + (B.band - gw) / 2; const y0 = B.Y(Math.max(l.from, B.S.lo)), y1 = B.Y(Math.max(l.to, B.S.lo)); const col = spec.highlight && spec.highlight.cats ? (spec.highlight.cats.includes(cats[j]) ? (spec.accent || ACCENT) : GREY) : WF[l.kind];
    g += `<rect x="${x.toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${gw.toFixed(1)}" height="${Math.max(1, Math.abs(y1 - y0)).toFixed(1)}" style="fill:${col}"/>`;
    if (spec.connectors !== false && j < L.length - 1) g += `<line x1="${(x + gw).toFixed(1)}" y1="${y1.toFixed(1)}" x2="${(x + B.band).toFixed(1)}" y2="${y1.toFixed(1)}" class="ch-conn"/>`;
    if (spec.labels !== false) g += `<text x="${(x + gw / 2).toFixed(1)}" y="${(Math.min(y0, y1) - 5).toFixed(1)}" text-anchor="middle" class="ch-lab">${esc(B.f(l.kind === 'total' ? l.to : l.v))}</text>`; });
  const items = [['Increase', WF.up, false], ['Decrease', WF.down, false], ['Total', WF.total, false]];
  return frame(spec, W, H, g, spec.legend === 'none' ? '' : legend(spec, items, W, H, 'top'));
}
// Box and whisker: quartiles by Excel's default “exclusive median” method (like QUARTILE.EXC), whiskers to the furthest
// point within 1.5 × IQR, dots beyond, and × for the mean.
export function boxStats(values) {
  const v = nums(values).slice().sort((a, b) => a - b); const n = v.length; if (!n) return null;
  const at = p => { const pos = p * (n + 1); if (pos <= 1) return v[0]; if (pos >= n) return v[n - 1]; const k = Math.floor(pos); return v[k - 1] + (pos - k) * (v[k] - v[k - 1]); };
  const q1 = at(0.25), med = at(0.5), q3 = at(0.75), iqr = q3 - q1; const inn = v.filter(x => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr);
  return { n, q1, med, q3, mean: v.reduce((a, b) => a + b, 0) / n, lo: Math.min(...inn), hi: Math.max(...inn), out: v.filter(x => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr) };
}
function boxSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const groups = (spec.groups || []).map(gp => ({ name: gp.name, st: boxStats(gp.values) })).filter(gp => gp.st);
  if (!groups.length) return empty(spec, W, H, 'Pick a column of numbers (and, if you like, a column of groups)');
  const all = groups.flatMap(gp => [gp.st.lo, gp.st.hi, ...gp.st.out]); let lo = Math.min(...all), hi = Math.max(...all); lo = spec.yMin ?? autoMin(lo, hi);
  const B = bands(spec, W, H, groups.map(gp => gp.name), lo, hi, spec.fmt); let g = B.g; const pal = PALETTES[spec.palette] || PALETTES.office; const bw = Math.min(70, B.band * 0.5);
  groups.forEach((gp, j) => { const st = gp.st; const cx = B.left + B.band * (j + .5); const col = spec.highlight && spec.highlight.cats ? (spec.highlight.cats.includes(gp.name) ? (spec.accent || ACCENT) : GREY) : (spec.oneColour === false ? pal[j % pal.length] : pal[0]);
    g += `<line x1="${cx}" y1="${B.Y(st.hi).toFixed(1)}" x2="${cx}" y2="${B.Y(st.q3).toFixed(1)}" class="ch-whisk"/><line x1="${cx}" y1="${B.Y(st.q1).toFixed(1)}" x2="${cx}" y2="${B.Y(st.lo).toFixed(1)}" class="ch-whisk"/>`;
    g += `<line x1="${cx - bw / 4}" y1="${B.Y(st.hi).toFixed(1)}" x2="${cx + bw / 4}" y2="${B.Y(st.hi).toFixed(1)}" class="ch-whisk"/><line x1="${cx - bw / 4}" y1="${B.Y(st.lo).toFixed(1)}" x2="${cx + bw / 4}" y2="${B.Y(st.lo).toFixed(1)}" class="ch-whisk"/>`;
    g += `<rect x="${(cx - bw / 2).toFixed(1)}" y="${B.Y(st.q3).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(1, B.Y(st.q1) - B.Y(st.q3)).toFixed(1)}" style="fill:${col}" opacity=".85"/><line x1="${(cx - bw / 2).toFixed(1)}" y1="${B.Y(st.med).toFixed(1)}" x2="${(cx + bw / 2).toFixed(1)}" y2="${B.Y(st.med).toFixed(1)}" class="ch-med"/>`;
    const my = B.Y(st.mean); g += `<path d="M${cx - 4},${my - 4}L${cx + 4},${my + 4}M${cx + 4},${my - 4}L${cx - 4},${my + 4}" class="ch-mean"/>`;
    g += st.out.map(o => `<circle cx="${cx}" cy="${B.Y(o).toFixed(1)}" r="3" class="ch-out" style="stroke:${col}"/>`).join('');
    if (spec.labels) g += `<text x="${(cx + bw / 2 + 4).toFixed(1)}" y="${(B.Y(st.med) + 4).toFixed(1)}" class="ch-lab">${esc(B.f(st.med))}</text>`; });
  return frame(spec, W, H, g, '');
}
// Funnel: stages top to bottom, each bar centred, widths in proportion to the values.
function funnelSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const s = (spec.series || [])[0]; const cats = (spec.cats || []).map(String); if (!s || !cats.length) return empty(spec, W, H, 'Pick the stages and their numbers');
  const vals = s.values.map(v => Math.max(0, +v || 0)); const max = Math.max(...vals) || 1; const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + 6; const left = Math.min(170, Math.max(...cats.map(c => tw(c, 10.5))) + 16);
  const plotW = W - left - 14, plotH = H - top - 10; const band = plotH / cats.length; const cx = left + plotW / 2; let g = ''; const f = v => fmtNum(v, spec.fmt || 'num', spec.dp);
  vals.forEach((v, j) => { const w = Math.max(2, v / max * plotW); const y = top + band * j; const col = spec.highlight && spec.highlight.cats ? (spec.highlight.cats.includes(cats[j]) ? (spec.accent || ACCENT) : GREY) : (spec.accent || PALETTES.office[0]);
    g += `<rect x="${(cx - w / 2).toFixed(1)}" y="${(y + band * 0.08).toFixed(1)}" width="${w.toFixed(1)}" height="${(band * 0.84).toFixed(1)}" style="fill:${col}"/><text x="${left - 8}" y="${(y + band / 2 + 4).toFixed(1)}" text-anchor="end" class="ch-cat">${esc(cats[j])}</text>`;
    if (spec.labels !== false) { const txt = f(v) + (spec.pctOfFirst && j ? ` (${Math.round(v / (vals[0] || 1) * 100)}%)` : ''); g += w > tw(txt, 11) + 8 ? `<text x="${cx}" y="${(y + band / 2 + 4).toFixed(1)}" text-anchor="middle" class="ch-lab in">${esc(txt)}</text>` : `<text x="${(cx + w / 2 + 5).toFixed(1)}" y="${(y + band / 2 + 4).toFixed(1)}" class="ch-lab">${esc(txt)}</text>`; } });
  return frame(spec, W, H, g, '');
}
// Treemap (squarified) and sunburst, from rows of [level 1, level 2 …, value]
function squarify(items, x, y, w, h) { // items: [{v, …}] sorted desc → rects
  const out = []; let rest = items.slice(); const total = rest.reduce((a, b) => a + b.v, 0) || 1; const scale = w * h / total; let X = x, Y = y, Wd = w, Ht = h;
  while (rest.length) { const side = Math.min(Wd, Ht); let row = [], best = Infinity;
    for (let i = 0; i < rest.length; i++) { const r = rest.slice(0, i + 1); const area = r.reduce((a, b) => a + b.v * scale, 0); const worst = Math.max(...r.map(it => Math.max(side * side * it.v * scale / (area * area), area * area / (side * side * it.v * scale)))); if (worst <= best) { best = worst; row = r; } else break; }
    const area = row.reduce((a, b) => a + b.v * scale, 0); const thick = area / side; let off = 0;
    row.forEach(it => { const len = it.v * scale / thick; if (Wd >= Ht) out.push({ ...it, x: X, y: Y + off, w: thick, h: len }); else out.push({ ...it, x: X + off, y: Y, w: len, h: thick }); off += len; });
    if (Wd >= Ht) { X += thick; Wd -= thick; } else { Y += thick; Ht -= thick; } rest = rest.slice(row.length); }
  return out;
}
function hierarchy(spec) { const rows = (spec.rows || []).filter(r => typeof r[r.length - 1] === 'number' && r[r.length - 1] > 0); const top = new Map();
  rows.forEach(r => { const k = String(r[0]); if (!top.has(k)) top.set(k, { name: k, v: 0, kids: [] }); const t = top.get(k); t.v += r[r.length - 1]; if (r.length > 2) t.kids.push({ name: String(r[1]), v: r[r.length - 1] }); });
  return [...top.values()].sort((a, b) => b.v - a.v).map(t => ({ ...t, kids: t.kids.sort((a, b) => b.v - a.v) })); }
function treemapSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const tops = hierarchy(spec); if (!tops.length) return empty(spec, W, H, 'Pick the category columns and a column of values');
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + 4; const pal = PALETTES[spec.palette] || PALETTES.office; let g = ''; const f = v => fmtNum(v, spec.fmt || 'num', spec.dp);
  squarify(tops, 6, top, W - 12, H - top - 6).forEach((t, i) => { const col = spec.highlight && spec.highlight.cats ? (spec.highlight.cats.includes(t.name) ? (spec.accent || ACCENT) : GREY) : pal[i % pal.length]; const ban = t.kids.length && t.h > 40 ? 16 : 0;
    g += `<rect x="${t.x.toFixed(1)}" y="${t.y.toFixed(1)}" width="${t.w.toFixed(1)}" height="${t.h.toFixed(1)}" style="fill:${col}" class="ch-tile"/>`;
    if (ban) g += `<text x="${(t.x + 4).toFixed(1)}" y="${(t.y + 12).toFixed(1)}" class="ch-tl b">${esc(t.w > tw(t.name, 10.5) + 8 ? t.name : '')}</text>`;
    const kids = t.kids.length ? squarify(t.kids, t.x, t.y + ban, t.w, t.h - ban) : [{ ...t, x: t.x, y: t.y, w: t.w, h: t.h }];
    kids.forEach(k => { if (t.kids.length) g += `<rect x="${k.x.toFixed(1)}" y="${k.y.toFixed(1)}" width="${k.w.toFixed(1)}" height="${k.h.toFixed(1)}" style="fill:${col}" class="ch-tile"/>`;
      if (k.w > tw(k.name, 10) + 8 && k.h > 18) g += `<text x="${(k.x + 4).toFixed(1)}" y="${(k.y + 13).toFixed(1)}" class="ch-tl">${esc(k.name)}</text>` + (spec.labels !== false && k.h > 32 && k.w > tw(f(k.v), 10) + 8 ? `<text x="${(k.x + 4).toFixed(1)}" y="${(k.y + 26).toFixed(1)}" class="ch-tl v">${esc(f(k.v))}</text>` : ''); }); });
  return frame(spec, W, H, g, '');
}
function sunburstSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const tops = hierarchy(spec); if (!tops.length) return empty(spec, W, H, 'Pick the category columns and a column of values');
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0); const R = Math.min(W / 2 - 10, (H - top - 10) / 2); const cx = W / 2, cy = top + (H - top) / 2; const tot = tops.reduce((a, t) => a + t.v, 0) || 1;
  const pal = PALETTES[spec.palette] || PALETTES.office; const deep = tops.some(t => t.kids.length); const r0 = R * 0.22, r1 = deep ? R * 0.6 : R, r2 = R; let g = ''; let a = -Math.PI / 2;
  const arc = (a1, a2, ri, ro) => { const p = (an, r) => `${(cx + r * Math.cos(an)).toFixed(1)},${(cy + r * Math.sin(an)).toFixed(1)}`; const big = a2 - a1 > Math.PI ? 1 : 0; return `M${p(a1, ro)}A${ro},${ro} 0 ${big},1 ${p(a2, ro)}L${p(a2, ri)}A${ri},${ri} 0 ${big},0 ${p(a1, ri)}Z`; };
  const lab = (a1, a2, ri, ro, txt) => { const m = (a1 + a2) / 2, rr = (ri + ro) / 2; const len = (a2 - a1) * rr; return len > tw(txt, 10) + 6 && ro - ri > 14 ? `<text x="${(cx + rr * Math.cos(m)).toFixed(1)}" y="${(cy + rr * Math.sin(m) + 4).toFixed(1)}" text-anchor="middle" class="ch-tl">${esc(txt)}</text>` : ''; };
  tops.forEach((t, i) => { const da = t.v / tot * Math.PI * 2; const col = pal[i % pal.length]; g += `<path d="${arc(a, a + da - 0.004, r0, r1)}" style="fill:${col}" class="ch-tile"/>` + lab(a, a + da, r0, r1, t.name); let b = a;
    t.kids.forEach(k => { const db = k.v / tot * Math.PI * 2; g += `<path d="${arc(b, b + db - 0.004, r1 + 1, r2)}" style="fill:${col}" opacity=".78" class="ch-tile"/>` + lab(b, b + db, r1, r2, k.name); b += db; }); a += da; });
  return frame(spec, W, H, g, '');
}
// Stock (open-high-low-close): a thin line from low to high, a box from open to close (hollow when it closed up)
function stockSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const rows = (spec.ohlc || []).filter(r => r.slice(1).every(v => typeof v === 'number')); if (!rows.length) return empty(spec, W, H, 'Pick Date, Open, High, Low and Close, in that order');
  const lo = Math.min(...rows.map(r => r[3])), hi = Math.max(...rows.map(r => r[2])); const cats = rows.map(r => typeof r[0] === 'number' ? fmtNum(r[0], 'date') : String(r[0]));
  const B = bands(spec, W, H, cats, spec.yMin ?? lo - (hi - lo) * 0.1, hi, spec.fmt); let g = B.g; const bw = Math.max(3, B.band * 0.55);
  rows.forEach((r, j) => { const [, o, h, l, c] = r; const x = B.left + B.band * (j + .5); const up = c >= o; g += `<line x1="${x}" y1="${B.Y(h).toFixed(1)}" x2="${x}" y2="${B.Y(l).toFixed(1)}" class="ch-wick"/><rect x="${(x - bw / 2).toFixed(1)}" y="${B.Y(Math.max(o, c)).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.max(1, Math.abs(B.Y(o) - B.Y(c))).toFixed(1)}" class="${up ? 'ch-up' : 'ch-down'}"/>`; });
  return frame(spec, W, H, g, '');
}
// A tile map: each state is a square at roughly its place on the map (Excel's filled map draws the real outlines)
export const IN_TILES = { 'Jammu and Kashmir': [2, 0, 'JK'], 'Ladakh': [3, 0, 'LA'], 'Punjab': [2, 1, 'PB'], 'Himachal Pradesh': [3, 1, 'HP'], 'Uttarakhand': [4, 1, 'UK'], 'Chandigarh': [1, 1, 'CH'], 'Sikkim': [6, 1, 'SK'], 'Arunachal Pradesh': [8, 1, 'AR'],
  'Rajasthan': [1, 2, 'RJ'], 'Haryana': [2, 2, 'HR'], 'Delhi': [3, 2, 'DL'], 'Uttar Pradesh': [4, 2, 'UP'], 'Bihar': [5, 2, 'BR'], 'Assam': [7, 2, 'AS'], 'Nagaland': [8, 2, 'NL'],
  'Gujarat': [0, 3, 'GJ'], 'Madhya Pradesh': [2, 3, 'MP'], 'Chhattisgarh': [3, 3, 'CG'], 'Jharkhand': [4, 3, 'JH'], 'West Bengal': [5, 3, 'WB'], 'Meghalaya': [7, 3, 'ML'], 'Manipur': [8, 3, 'MN'],
  'Maharashtra': [1, 4, 'MH'], 'Telangana': [2, 4, 'TG'], 'Odisha': [4, 4, 'OD'], 'Tripura': [6, 4, 'TR'], 'Mizoram': [7, 4, 'MZ'], 'Goa': [1, 5, 'GA'], 'Karnataka': [2, 5, 'KA'], 'Andhra Pradesh': [3, 5, 'AP'], 'Kerala': [2, 6, 'KL'], 'Tamil Nadu': [3, 6, 'TN'] };
function tilemapSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const s = (spec.series || [])[0]; const cats = (spec.cats || []).map(String); const val = {}; if (s) cats.forEach((c, j) => { if (typeof s.values[j] === 'number') val[c] = s.values[j]; });
  const vs = Object.values(val); if (!vs.length) return empty(spec, W, H, 'Pick the states and their values');
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + 4; const lo = Math.min(...vs), hi = Math.max(...vs); const cell = Math.min((W - 130) / 9, (H - top - 8) / 7); const x0 = 10, y0 = top;
  const shade = v => { const t = hi > lo ? (v - lo) / (hi - lo) : 1; const mix = (a, b) => Math.round(a + (b - a) * t); return `rgb(${mix(222, 31)},${mix(235, 78)},${mix(247, 135)})`; }; let g = '';
  for (const [name, [c, r, code]] of Object.entries(IN_TILES)) { const has = name in val; const x = x0 + c * cell, y = y0 + r * cell; g += `<rect x="${(x + 1).toFixed(1)}" y="${(y + 1).toFixed(1)}" width="${(cell - 2).toFixed(1)}" height="${(cell - 2).toFixed(1)}" rx="3" style="fill:${has ? shade(val[name]) : 'var(--ch-empty-tile,#EEEEEE)'}"><title>${esc(name)}${has ? ': ' + esc(fmtNum(val[name], spec.fmt || 'num', spec.dp)) : ''}</title></rect><text x="${(x + cell / 2).toFixed(1)}" y="${(y + cell / 2 + 4).toFixed(1)}" text-anchor="middle" class="ch-tl${has && (val[name] - lo) / ((hi - lo) || 1) > 0.55 ? '' : ' dk'}">${code}</text>`; }
  const lx = x0 + 9 * cell + 16, ly = y0 + 10; g += `<defs><linearGradient id="tmg" x1="0" y1="1" x2="0" y2="0"><stop offset="0" stop-color="rgb(222,235,247)"/><stop offset="1" stop-color="rgb(31,78,135)"/></linearGradient></defs><rect x="${lx}" y="${ly}" width="14" height="${(cell * 4).toFixed(1)}" fill="url(#tmg)"/>`;
  g += `<text x="${lx + 20}" y="${ly + 10}" class="ch-tick">${esc(fmtNum(hi, spec.fmt || 'num', spec.dp))}</text><text x="${lx + 20}" y="${(ly + cell * 4).toFixed(1)}" class="ch-tick">${esc(fmtNum(lo, spec.fmt || 'num', spec.dp))}</text>`;
  return frame(spec, W, H, g, '');
}
// Bullet chart (built by hand in Excel): grey bands for poor / fair / good, a dark bar for the actual, a tick for the target
function bulletSVG(spec) {
  const W = spec.width || 560, H = spec.height || 320; const rows = spec.rows || []; if (!rows.length) return empty(spec, W, H, 'No measures');
  const top = titleLayout(spec, W).h + (spec.subtitle ? 18 : 0) + 6; const left = Math.min(190, Math.max(...rows.map(r => tw(r.name, 11))) + 16); const plotW = W - left - 20; const band = (H - top - 10) / rows.length; let g = '';
  rows.forEach((r, j) => { const max = r.max ?? Math.max(...r.bands, r.value, r.target) * 1.02; const X = v => left + v / max * plotW; const y = top + band * j; const h = Math.min(30, band * 0.62); const yy = y + (band - h) / 2;
    let prev = 0; r.bands.forEach((b, k) => { g += `<rect x="${X(prev).toFixed(1)}" y="${yy.toFixed(1)}" width="${(X(b) - X(prev)).toFixed(1)}" height="${h.toFixed(1)}" class="ch-band b${k}"/>`; prev = b; });
    g += `<rect x="${left}" y="${(yy + h * 0.35).toFixed(1)}" width="${(X(r.value) - left).toFixed(1)}" height="${(h * 0.3).toFixed(1)}" style="fill:${r.color || 'var(--ch-ink,#262626)'}"/><line x1="${X(r.target).toFixed(1)}" y1="${(yy + h * 0.12).toFixed(1)}" x2="${X(r.target).toFixed(1)}" y2="${(yy + h * 0.88).toFixed(1)}" class="ch-target"/>`;
    g += `<text x="${left - 8}" y="${(yy + h / 2 + 4).toFixed(1)}" text-anchor="end" class="ch-cat">${esc(r.name)}</text>`; if (spec.labels !== false) g += `<text x="${(X(r.value) + 4).toFixed(1)}" y="${(yy + h * 0.35 - 3).toFixed(1)}" class="ch-lab">${esc(fmtNum(r.value, r.fmt || spec.fmt || 'num', spec.dp))}</text>`; });
  return frame(spec, W, H, g, '');
}

// ---- sparklines: a tiny chart that lives in one cell. sp = { type: 'line'|'column'|'winloss', values, high, low, markers, negative, lo, hi, color }
export function sparkSVG(sp, w, h) {
  const vals = sp.values || []; const nums = vals.filter(v => typeof v === 'number'); if (!nums.length) return '';
  const col = sp.color || '#376092', HI = '#00A050', LO = '#D00000', NEG = '#D00000'; const pad = 3; const W = Math.max(10, w - pad * 2), H = Math.max(6, h - pad * 2);
  let lo = sp.lo ?? Math.min(...nums), hi = sp.hi ?? Math.max(...nums); const n = vals.length; const mx = Math.max(...nums), mn = Math.min(...nums); let out = '';
  if (sp.type === 'winloss') { const bw = W / n; vals.forEach((v, i) => { if (typeof v !== 'number' || v === 0) return; out += `<rect x="${(pad + i * bw + bw * .15).toFixed(1)}" y="${v > 0 ? pad : pad + H / 2}" width="${(bw * .7).toFixed(1)}" height="${(H / 2).toFixed(1)}" fill="${v < 0 ? NEG : col}"/>`; }); }
  else if (sp.type === 'column') { if (sp.lo == null && lo > 0) lo = lo - (hi - lo) * 0.15; else lo = Math.min(lo, 0); hi = Math.max(hi, 0); /* like Excel: columns start near their own smallest value unless the axis is shared */ const span = hi - lo || 1; const y = v => pad + H - (v - lo) / span * H; const bw = W / n;
    vals.forEach((v, i) => { if (typeof v !== 'number') return; const y0 = y(Math.max(v, Math.max(lo, 0))), y1 = y(Math.min(v, Math.max(lo, 0))); const c = sp.high && v === mx ? HI : sp.low && v === mn ? LO : v < 0 && sp.negative ? NEG : col; out += `<rect x="${(pad + i * bw + bw * .15).toFixed(1)}" y="${y0.toFixed(1)}" width="${(bw * .7).toFixed(1)}" height="${Math.max(1, y1 - y0).toFixed(1)}" fill="${c}"/>`; }); }
  else { const span = hi - lo || 1; const x = i => pad + (n === 1 ? W / 2 : i / (n - 1) * W), y = v => pad + H - (v - lo) / span * H; let d = '', pen = false;
    vals.forEach((v, i) => { if (typeof v !== 'number') { pen = false; return; } d += `${pen ? 'L' : 'M'}${x(i).toFixed(1)} ${y(v).toFixed(1)}`; pen = true; }); out += `<path d="${d}" fill="none" stroke="${col}" stroke-width="1.5" stroke-linejoin="round"/>`;
    vals.forEach((v, i) => { if (typeof v !== 'number') return; const c = sp.high && v === mx ? HI : sp.low && v === mn ? LO : v < 0 && sp.negative ? NEG : sp.markers ? col : null; if (c) out += `<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="2.2" fill="${c}"/>`; }); }
  return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true">${out}</svg>`;
}
