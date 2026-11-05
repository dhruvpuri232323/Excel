// Sheetah live grid: an Excel-like sheet built on engine.js.
import { Workbook, addr, parseAddr, parseRange, rangeStr, colName, formatValue, fillSeries, shiftFormula, cycleRefAt, tokenize, explainFormula, isErr, parseInput, FUNCTION_NAMES } from './engine.js';

const ROW_H = 24, HEAD_W = 46, HEAD_H = 24, DEF_W = 80;
const REF_COLORS = ['#3557D6', '#D9480F', '#7B3FB0', '#1F8A5B', '#C2255C', '#0B7285'];
export const NUMBER_FORMATS = [
  ['General', 'General'], ['Number', '0.00'], ['Comma', '#,##0.00'], ['Indian comma', '#,##,##0'], ['Rupees', '₹#,##0.00'],
  ['Rupees, Indian grouping', '[$₹-4009] #,##,##0.00'], ['Percentage', '0%'], ['Percentage, 2 dp', '0.00%'],
  ['Short date', 'dd-mm-yyyy'], ['Long date', 'dd mmmm yyyy'], ['Time', 'hh:mm'], ['Text', '@']
];
const FILLS = ['', '#FFF2CC', '#DDEBF7', '#E2EFDA', '#FCE4D6', '#EDE7F6', '#F2F2F2', '#23263A'];
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html !== undefined) e.innerHTML = html; return e; };

export class Grid {
  constructor(host, wb, opts = {}) {
    this.host = host; this.wb = wb instanceof Workbook ? wb : new Workbook(wb); this.o = Object.assign({ toolbar: true, formulaBar: true, tabs: true, status: true, height: 360, readOnly: false, editable: null, highlight: null, lockMessage: 'This cell holds the data. Work in the highlighted cells.' }, opts);
    this.RH = this.o.rowH || (matchMedia('(pointer: coarse)').matches ? 30 : ROW_H); this.host.style.setProperty('--xg-rh', this.RH + 'px');
    this.si = 0; this.sel = { r: 0, c: 0, ar: 0, ac: 0 }; this.editing = null; this.showFormulas = false; this.history = []; this.future = []; this.clip = null; this.refOverlays = [];
    this.build(); this.unsub = this.wb.onChange(() => this.schedule());
    if (opts.select) this.select(opts.select); this.touchKeys(); this.render();
  }
  // a row of formula keys under the sheet on touch screens (phones and tablets have no = $ : keys handy)
  touchKeys() {
    if (this.o.readOnly || !matchMedia('(pointer: coarse)').matches || this.keysBar) return;
    const bar = el('div', 'xg-keys'); bar.innerHTML = ['Type here', '=', '$', ':', '(', ')', ',', '+', '-', '*', '/', 'Enter'].map(k => `<button type="button" data-k="${k}">${k}</button>`).join('');
    bar.addEventListener('mousedown', e => e.preventDefault());
    bar.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; const k = b.dataset.k;
      if (k === 'Type here') { if (!this.editing) this.beginEdit(null, 'cell'); return; }
      if (k === 'Enter') { if (this.editing) { this.commit(); this.move(1, 0); } return; }
      if (!this.editing) { this.beginEdit(k, 'cell'); return; }
      const inp = this.editing.via === 'bar' ? this.fInput : this.cellInput; if (!inp) return; const a = inp.selectionStart ?? inp.value.length, z = inp.selectionEnd ?? a;
      inp.value = inp.value.slice(0, a) + k + inp.value.slice(z); inp.setSelectionRange(a + 1, a + 1); inp.dispatchEvent(new Event('input')); inp.focus(); });
    this.host.insertBefore(bar, this.tabsEl || this.statusEl || null); this.keysBar = bar;
  }
  // ---------- layout ----------
  build() {
    const h = this.host; h.innerHTML = ''; h.classList.add('xg');
    if (this.o.toolbar && !this.o.readOnly) h.appendChild(this.buildToolbar());
    if (this.o.formulaBar) {
      const fb = el('div', 'xg-fbar');
      this.nameBox = el('input', 'xg-name'); this.nameBox.setAttribute('aria-label', 'Name Box'); this.nameBox.spellcheck = false;
      this.fx = el('span', 'xg-fx', 'fx'); this.fInput = el('input', 'xg-finput'); this.fInput.setAttribute('aria-label', 'Formula bar'); this.fInput.spellcheck = false;
      if (this.o.readOnly) this.fInput.readOnly = true;
      fb.append(this.nameBox, this.fx, this.fInput); h.appendChild(fb);
      this.nameBox.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); this.nameBoxGo(this.nameBox.value.trim()); this.focus(); } if (e.key === 'Escape') { this.updateBars(); this.focus(); } });
      this.fInput.addEventListener('focus', () => { if (this.o.readOnly) return; if (!this.editing) this.beginEdit(null, 'bar'); });
      this.fInput.addEventListener('input', () => { if (this.editing) { this.editing.text = this.fInput.value; if (this.cellInput) this.cellInput.value = this.fInput.value; this.showRefs(); this.suggest(); } });
      this.fInput.addEventListener('keydown', e => this.editKey(e, this.fInput));
    }
    this.view = el('div', 'xg-view'); this.view.style.height = this.o.height + 'px'; this.view.tabIndex = 0;
    this.view.setAttribute('role', 'grid'); this.view.setAttribute('aria-label', 'Worksheet');
    this.corner = el('div', 'xg-corner'); this.colHead = el('div', 'xg-colhead'); this.rowHead = el('div', 'xg-rowhead');
    this.scroller = el('div', 'xg-scroll'); this.spacer = el('div', 'xg-spacer'); this.layer = el('div', 'xg-layer');
    this.selBox = el('div', 'xg-sel'); this.actBox = el('div', 'xg-act'); this.handle = el('div', 'xg-handle'); this.handle.title = 'Fill handle: drag to copy or continue a series';
    this.fillGhost = el('div', 'xg-fillghost');
    this.layer.append(this.selBox, this.actBox, this.fillGhost); if (!this.o.readOnly) this.layer.append(this.handle);
    this.scroller.append(this.spacer, this.layer); this.view.append(this.scroller, this.colHead, this.rowHead, this.corner); h.appendChild(this.view);
    this.corner.title = 'Select all'; this.corner.addEventListener('mousedown', e => { e.preventDefault(); const s = this.sheet; this.setSel(0, 0, s.rows - 1, s.cols - 1); this.scroller.scrollTop = 0; this.scroller.scrollLeft = 0; });
    if (this.o.tabs) { this.tabsEl = el('div', 'xg-tabs'); h.appendChild(this.tabsEl); }
    if (this.o.status) { this.statusEl = el('div', 'xg-status'); h.appendChild(this.statusEl); }
    this.toast = el('div', 'xg-toast'); h.appendChild(this.toast);
    this.scroller.addEventListener('scroll', () => this.schedule());
    this.view.addEventListener('mousedown', e => this.mouseDown(e));
    this.view.addEventListener('dblclick', e => { const p = this.hit(e); if (p && !this.o.readOnly) { this.setSel(p.r, p.c); this.beginEdit(null, 'cell'); } });
    this.view.addEventListener('keydown', e => this.key(e));
    this.view.addEventListener('contextmenu', e => this.contextMenu(e));
    this.view.addEventListener('paste', e => this.pasteEvent(e));
    this.view.addEventListener('copy', e => { if (!this.editing) { this.copy(); e.clipboardData.setData('text/plain', this.selectionTSV()); e.preventDefault(); } });
    this.view.addEventListener('cut', e => { if (!this.editing) { this.copy(true); e.clipboardData.setData('text/plain', this.selectionTSV()); e.preventDefault(); } });
    this.ro = new ResizeObserver(() => this.schedule()); this.ro.observe(this.view);
  }
  buildToolbar() {
    const t = el('div', 'xg-tools');
    const btn = (label, title, fn, cls = '') => { const b = el('button', 'xg-btn ' + cls, label); b.type = 'button'; b.title = title; b.setAttribute('aria-label', title); b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => { fn(); this.focus(); }); t.appendChild(b); return b; };
    btn('↶', 'Undo (Ctrl+Z)', () => this.undo()); btn('↷', 'Redo (Ctrl+Y)', () => this.redo());
    t.appendChild(el('span', 'xg-sepr'));
    btn('<b>B</b>', 'Bold (Ctrl+B)', () => this.toggleStyle('bold')); btn('<i>I</i>', 'Italic (Ctrl+I)', () => this.toggleStyle('italic'));
    const fill = btn('<span class="xg-fillicon"></span>', 'Fill colour', () => {}); fill.addEventListener('click', e => this.palette(fill));
    btn('⇤', 'Align left', () => this.applyStyle({ align: 'left' })); btn('≡', 'Centre', () => this.applyStyle({ align: 'center' })); btn('⇥', 'Align right', () => this.applyStyle({ align: 'right' }));
    btn('↵', 'Wrap text', () => this.toggleStyle('wrap'));
    t.appendChild(el('span', 'xg-sepr'));
    const sel = el('select', 'xg-fmt'); sel.title = 'Number format'; sel.setAttribute('aria-label', 'Number format');
    for (const [n, f] of NUMBER_FORMATS) { const o = el('option'); o.value = f; o.textContent = n; sel.appendChild(o); }
    const oc = el('option'); oc.value = '__custom'; oc.textContent = 'Custom…'; sel.appendChild(oc);
    sel.addEventListener('change', () => { let f = sel.value; if (f === '__custom') { f = prompt('Custom number format code, for example #,##0.00 or dd-mmm-yy', this.wb.style(this.si, this.sel.r, this.sel.c).fmt || 'General'); if (!f) { this.updateBars(); return; } } this.applyStyle({ fmt: f === 'General' ? null : f }); this.focus(); });
    t.appendChild(sel); this.fmtSel = sel;
    btn('.0', 'Decrease decimal', () => this.decimals(-1)); btn('.00', 'Increase decimal', () => this.decimals(1));
    t.appendChild(el('span', 'xg-sepr'));
    btn('Σ', 'AutoSum (Alt+=)', () => this.autoSum());
    btn('A→Z', 'Sort smallest to largest', () => this.sort(false)); btn('Z→A', 'Sort largest to smallest', () => this.sort(true));
    btn('fx?', 'Explain this formula step by step', () => this.explain());
    btn('{=}', 'Show formulas (Ctrl+`)', () => { this.showFormulas = !this.showFormulas; this.render(); }, 'xg-showf');
    btn('🔍', 'Find and replace (Ctrl+H)', () => this.findDialog(true));
    return t;
  }
  get sheet() { return this.wb.sheets[this.si]; }
  colW(c) { const w = this.sheet.widths[c]; return w ? (w > 40 ? w : Math.round(w * 7 + 5)) : DEF_W; }
  colX(c) { let x = 0; for (let i = 0; i < c; i++) x += this.colW(i); return x; }
  colAt(x) { let acc = 0; for (let c = 0; c < this.sheet.cols; c++) { const w = this.colW(c); if (x < acc + w) return c; acc += w; } return this.sheet.cols - 1; }
  schedule() { if (this._raf) return; this._raf = requestAnimationFrame(() => { this._raf = null; this.render(); }); }
  destroy() { this.unsub && this.unsub(); this.ro && this.ro.disconnect(); this.host.innerHTML = ''; }
  focus() { if (!this.editing) this.view.focus({ preventScroll: true }); }

  // ---------- rendering ----------
  render() {
    const s = this.sheet; if (!s) return;
    const totalW = this.colX(s.cols), totalH = s.rows * this.RH;
    this.spacer.style.width = totalW + HEAD_W + 'px'; this.spacer.style.height = totalH + HEAD_H + 'px';
    const sl = this.scroller.scrollLeft, st = this.scroller.scrollTop, vw = this.scroller.clientWidth, vh = this.scroller.clientHeight;
    const r1 = Math.max(0, Math.floor(st / this.RH) - 1), r2 = Math.min(s.rows - 1, Math.ceil((st + vh) / this.RH) + 1);
    const c1 = Math.max(0, this.colAt(sl) - 1), c2 = Math.min(s.cols - 1, this.colAt(sl + vw) + 1);
    // column & row headers
    const g = this.selRange();
    let ch = ''; for (let c = c1; c <= c2; c++) { const on = c >= g.c1 && c <= g.c2; ch += `<div class="xg-ch${on ? ' on' : ''}" data-c="${c}" style="left:${this.colX(c) - sl + HEAD_W}px;width:${this.colW(c)}px">${colName(c)}<span class="xg-cresize" data-cr="${c}"></span></div>`; }
    this.colHead.innerHTML = ch;
    let rh = ''; for (let r = r1; r <= r2; r++) { const on = r >= g.r1 && r <= g.r2; rh += `<div class="xg-rh${on ? ' on' : ''}" data-r="${r}" style="top:${r * this.RH - st + HEAD_H}px">${r + 1}</div>`; }
    this.rowHead.innerHTML = rh;
    // cells
    const frag = []; const editable = this.o.editable ? this.editRanges() : null; const hl = this.o.highlight ? (this.o.highlight === this.o.editable ? editable : this.o.highlight.map(parseRange)) : null;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = this.wb.cellObj(s, r, c); const st2 = (cell && cell.s) || {};
        const x = this.colX(c) + HEAD_W, y = r * this.RH + HEAD_H; let w = this.colW(c);
        let text = '', cls = 'xg-c'; let v = null;
        if (cell && (cell.f || cell.v !== null)) {
          if (this.showFormulas && cell.f) { text = cell.f; cls += ' f'; }
          else { v = this.wb.value(this.si, r, c); text = formatValue(v, st2.fmt, typeof v === 'number' ? w : null); }
          if (typeof v === 'number' && !(this.showFormulas && cell.f)) cls += ' n'; else if (typeof v === 'boolean' || isErr(v)) cls += ' b';
          if (isErr(v)) cls += ' err';
          if (typeof v === 'string' && !cell.f && /^\s*-?[\d,]+(\.\d+)?\s*$/.test(v) && !(st2.fmt === '@')) cls += ' tn';
          // text overflow into empty neighbours
          if (typeof v === 'string' && !st2.wrap && st2.align !== 'right' && st2.align !== 'center') { let cc = c + 1; const need = v.length * 7 + 8; while (w < need && cc < s.cols && this.wb.isEmpty(this.si, r, cc)) { w += this.colW(cc); cc++; } }
        }
        let style = `left:${x}px;top:${y}px;width:${w}px;` + (w > this.colW(c) ? 'z-index:2;background:transparent;box-shadow:none;' : '');
        if (st2.bold) style += 'font-weight:700;'; if (st2.italic) style += 'font-style:italic;';
        if (st2.fill) style += `background:${st2.fill};` + (st2.fill === '#23263A' ? 'color:#fff;' : '');
        if (st2.color) style += `color:${st2.color};`;
        if (st2.align) style += `text-align:${st2.align};justify-content:${st2.align === 'center' ? 'center' : st2.align === 'right' ? 'flex-end' : 'flex-start'};`;
        if (st2.wrap) cls += ' wrap';
        if (st2.border) cls += ' bd';
        if (editable && !editable.some(g2 => g2 && r >= g2.r1 && r <= g2.r2 && c >= g2.c1 && c <= g2.c2)) cls += ' locked';
        if (hl && hl.some(g2 => g2 && r >= g2.r1 && r <= g2.r2 && c >= g2.c1 && c <= g2.c2)) cls += ' hl';
        frag.push(`<div class="${cls}" style="${style}"${cls.includes(' tn') ? ' title="Number stored as text"' : ''}>${esc(text)}</div>`);
      }
    }
    this.layer.querySelectorAll('.xg-c').forEach(n => n.remove());
    this.layer.insertAdjacentHTML('afterbegin', frag.join(''));
    this.colHead.style.left = '0'; this.rowHead.style.top = '0';
    this.placeSelection(); this.updateBars(); this.renderTabs(); this.renderStatus();
    if (this.editing && this.cellInput) this.placeEditor();
    if (this.editing) this.showRefs();
  }
  placeSelection() {
    const g = this.selRange(); const x1 = this.colX(g.c1) + HEAD_W, y1 = g.r1 * this.RH + HEAD_H, x2 = this.colX(g.c2 + 1) + HEAD_W, y2 = (g.r2 + 1) * this.RH + HEAD_H;
    Object.assign(this.selBox.style, { left: x1 + 'px', top: y1 + 'px', width: x2 - x1 + 'px', height: y2 - y1 + 'px', display: (g.r1 === g.r2 && g.c1 === g.c2) ? 'none' : 'block' });
    const ax = this.colX(this.sel.ac) + HEAD_W, ay = this.sel.ar * this.RH + HEAD_H;
    Object.assign(this.actBox.style, { left: ax + 'px', top: ay + 'px', width: this.colW(this.sel.ac) + 'px', height: this.RH + 'px' });
    Object.assign(this.handle.style, { left: x2 - 4 + 'px', top: y2 - 4 + 'px' });
  }
  renderTabs() {
    if (!this.tabsEl) return;
    this.tabsEl.innerHTML = this.wb.sheets.map((s, i) => `<button type="button" class="xg-tab${i === this.si ? ' on' : ''}" data-i="${i}">${esc(s.name)}</button>`).join('') + (this.o.readOnly || this.o.noAddSheet ? '' : '<button type="button" class="xg-tab add" title="New sheet (Shift+F11)">+</button>');
    this.tabsEl.querySelectorAll('.xg-tab').forEach(b => b.onclick = () => { if (b.classList.contains('add')) this.addSheet(); else this.setSheet(+b.dataset.i); });
    this.tabsEl.querySelectorAll('.xg-tab:not(.add)').forEach(b => b.ondblclick = () => { if (this.o.readOnly) return; const i = +b.dataset.i; const n = prompt('Rename sheet', this.wb.sheets[i].name); if (n && n.trim()) this.renameSheet(i, n.trim()); });
  }
  renderStatus() {
    if (!this.statusEl) return; const g = this.selRange(); let sum = 0, cnt = 0, nums = 0;
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 1 && (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) < 200000) {
      const u = this.wb.usedRange(this.sheet); if (u) for (let r = Math.max(g.r1, u.r1); r <= Math.min(g.r2, u.r2); r++) for (let c = Math.max(g.c1, u.c1); c <= Math.min(g.c2, u.c2); c++) { const v = this.wb.value(this.si, r, c); if (v !== null && v !== '') cnt++; if (typeof v === 'number') { sum += v; nums++; } }
      this.statusEl.innerHTML = cnt ? (nums ? `<span>Average: <b>${fmtStatus(sum / nums)}</b></span><span>Count: <b>${cnt}</b></span><span>Sum: <b>${fmtStatus(sum)}</b></span>` : `<span>Count: <b>${cnt}</b></span>`) : '<span>Ready</span>';
    } else this.statusEl.innerHTML = this.wb.circular ? '<span class="warn">Circular reference: a formula refers to its own cell</span>' : '<span>Ready</span>';
  }
  updateBars() {
    if (this.nameBox && document.activeElement !== this.nameBox) { const g = this.selRange(); const nm = Object.values(this.wb.names).find(d => d.ref.replace(/\$/g, '').replace(/^.*!/, '') === rangeStr(g)); this.nameBox.value = nm ? nm.name : (g.r1 === g.r2 && g.c1 === g.c2 ? addr(this.sel.ar, this.sel.ac) : `${g.r2 - g.r1 + 1}R × ${g.c2 - g.c1 + 1}C`); }
    if (this.fInput && !this.editing) this.fInput.value = this.wb.input(this.si, this.sel.ar, this.sel.ac);
    if (this.fmtSel) { const f = this.wb.style(this.si, this.sel.ar, this.sel.ac).fmt || 'General'; this.fmtSel.value = NUMBER_FORMATS.some(x => x[1] === f) ? f : 'General'; }
  }
  // ---------- selection ----------
  selRange() { const s = this.sel; return { r1: Math.min(s.ar, s.r), c1: Math.min(s.ac, s.c), r2: Math.max(s.ar, s.r), c2: Math.max(s.ac, s.c) }; }
  setSel(ar, ac, r = ar, c = ac) {
    const s = this.sheet; const clamp = (v, m) => Math.max(0, Math.min(m - 1, v));
    this.sel = { ar: clamp(ar, s.rows), ac: clamp(ac, s.cols), r: clamp(r, s.rows), c: clamp(c, s.cols) };
    this.scrollIntoView(this.sel.r, this.sel.c); this.render(); this.o.onSelect && this.o.onSelect(this.selection());
  }
  select(a) { const g = parseRange(a); if (g) this.setSel(g.r1, g.c1, g.r2, g.c2); }
  selection() { return { sheet: this.si, range: rangeStr(this.selRange()), active: addr(this.sel.ar, this.sel.ac) }; }
  scrollIntoView(r, c) {
    const sc = this.scroller; const x = this.colX(c), w = this.colW(c), y = r * this.RH; const vw = sc.clientWidth - HEAD_W, vh = sc.clientHeight - HEAD_H;
    if (x < sc.scrollLeft) sc.scrollLeft = x; else if (x + w > sc.scrollLeft + vw) sc.scrollLeft = x + w - vw;
    if (y < sc.scrollTop) sc.scrollTop = y; else if (y + this.RH > sc.scrollTop + vh) sc.scrollTop = y + this.RH - vh;
  }
  hit(e) {
    const rect = this.scroller.getBoundingClientRect(); const x = e.clientX - rect.left + this.scroller.scrollLeft - HEAD_W, y = e.clientY - rect.top + this.scroller.scrollTop - HEAD_H;
    const inColHead = e.clientY - rect.top < HEAD_H, inRowHead = e.clientX - rect.left < HEAD_W;
    if (x < 0 && y < 0) return null;
    return { r: Math.max(0, Math.min(this.sheet.rows - 1, Math.floor(y / this.RH))), c: this.colAt(Math.max(0, x)), colHead: inColHead, rowHead: inRowHead };
  }
  mouseDown(e) {
    if (e.button === 2) { const p = this.hit(e); const g = this.selRange(); if (p && !(p.r >= g.r1 && p.r <= g.r2 && p.c >= g.c1 && p.c <= g.c2)) this.setSel(p.r, p.c); return; }
    if (e.target.classList.contains('xg-cresize')) return this.resizeCol(e, +e.target.dataset.cr);
    if (e.target === this.handle) return this.fillDrag(e);
    const p = this.hit(e); if (!p) return;
    // point mode: clicking while writing a formula inserts a reference
    if (this.editing && this.editing.text.startsWith('=') && /[=(,+\-*/^&<>:]\s*$/.test(this.editing.text.slice(0, this.caretPos()))) { e.preventDefault(); return this.pointDrag(e, p); }
    if (this.editing) this.commit();
    e.preventDefault(); this.view.focus({ preventScroll: true });
    if (p.colHead) { this.setSel(0, p.c, this.sheet.rows - 1, p.c); this.sel.ar = 0; return; }
    if (p.rowHead) { this.setSel(p.r, 0, p.r, this.sheet.cols - 1); return; }
    if (e.shiftKey) this.setSel(this.sel.ar, this.sel.ac, p.r, p.c); else this.setSel(p.r, p.c);
    const move = ev => { const q = this.hit(ev); if (q) { this.sel.r = q.r; this.sel.c = q.c; this.scrollIntoView(q.r, q.c); this.schedule(); } };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.o.onSelect && this.o.onSelect(this.selection()); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
  }
  pointDrag(e, p) {
    const input = this.editing.via === 'bar' ? this.fInput : this.cellInput; const pos = this.caretPos(); const before = this.editing.text.slice(0, pos), after = this.editing.text.slice(pos);
    const sheetPrefix = this.si !== this.editing.si ? quoteSheet(this.sheet.name) + '!' : '';
    let start = p; const put = q => { const g = { r1: Math.min(start.r, q.r), c1: Math.min(start.c, q.c), r2: Math.max(start.r, q.r), c2: Math.max(start.c, q.c) }; const ref = sheetPrefix + rangeStr(g); this.editing.text = before + ref + after; input.value = this.editing.text; if (this.fInput) this.fInput.value = this.editing.text; if (this.cellInput) this.cellInput.value = this.editing.text; const cp = (before + ref).length; input.setSelectionRange(cp, cp); this.editing.caret = cp; this.showRefs(); };
    put(p); const move = ev => { const q = this.hit(ev); if (q) put(q); }; const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); input.focus(); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
  }
  caretPos() { const input = this.editing && (this.editing.via === 'bar' ? this.fInput : this.cellInput); return input ? (input.selectionStart ?? this.editing.text.length) : 0; }
  resizeCol(e, c) {
    e.preventDefault(); e.stopPropagation(); const x0 = e.clientX, w0 = this.colW(c);
    const move = ev => { this.sheet.widths[c] = Math.max(16, w0 + ev.clientX - x0); this.schedule(); };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.o.onChange && this.o.onChange(); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
    e.target.ondblclick = () => this.autoFit(c);
  }
  autoFit(c) { let max = 30; const u = this.wb.usedRange(this.sheet); if (u) for (let r = u.r1; r <= u.r2; r++) { const st = this.wb.style(this.si, r, c); const v = this.wb.value(this.si, r, c); const t = formatValue(v, st.fmt); max = Math.max(max, t.length * 7.4 + 14); } this.sheet.widths[c] = Math.round(max); this.render(); this.o.onChange && this.o.onChange(); }

  // ---------- editing ----------
  editRanges() { // 'B2:C5' applies to every sheet; 'Sheet!B2:C5' only to that sheet
    const name = this.wb.sheets[this.si] && this.wb.sheets[this.si].name;
    return (this.o.editable || []).map(e => { const m = /^'?([^'!]+)'?!(.+)$/.exec(e); if (!m) return parseRange(e); return m[1].toLowerCase() === String(name).toLowerCase() ? parseRange(m[2]) : null; });
  }
  canEdit(r, c, quiet) {
    if (this.o.readOnly) return false;
    if (!this.o.editable) return true;
    const ok = this.editRanges().some(g => g && r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);
    if (!ok && !quiet) this.flash(this.o.lockMessage); return ok;
  }
  beginEdit(initial, via = 'cell') {
    const { ar: r, ac: c } = this.sel; if (!this.canEdit(r, c)) return;
    const text = initial !== null && initial !== undefined ? initial : this.wb.input(this.si, r, c);
    this.editing = { r, c, si: this.si, text, via, fresh: initial !== null && initial !== undefined };
    if (via === 'cell') {
      this.cellInput = el('input', 'xg-edit'); this.cellInput.spellcheck = false; this.cellInput.value = text; this.layer.appendChild(this.cellInput); this.placeEditor();
      this.cellInput.addEventListener('input', () => { this.editing.text = this.cellInput.value; if (this.fInput) this.fInput.value = this.cellInput.value; this.showRefs(); this.suggest(); });
      this.cellInput.addEventListener('keydown', e => this.editKey(e, this.cellInput));
      this.cellInput.focus(); const n = text.length; this.cellInput.setSelectionRange(n, n);
    }
    if (this.fInput) this.fInput.value = text; this.showRefs();
  }
  placeEditor() { const { r, c } = this.editing; if (this.editing.si !== this.si) { this.cellInput.style.display = 'none'; return; } this.cellInput.style.display = ''; Object.assign(this.cellInput.style, { left: this.colX(c) + HEAD_W + 'px', top: r * this.RH + HEAD_H + 'px', minWidth: this.colW(c) + 'px', height: this.RH + 'px' }); }
  editKey(e, input) {
    if (!this.editing) { if (e.key === 'Enter') { e.preventDefault(); this.focus(); } return; }
    if (this.sugBox && this.sugBox.style.display !== 'none' && (e.key === 'Tab' || (e.key === 'Enter' && this.sugActive))) { e.preventDefault(); this.acceptSuggest(input); return; }
    if (e.key === 'Enter') { e.preventDefault(); this.commit(); this.move(e.shiftKey ? -1 : 1, 0); this.focus(); }
    else if (e.key === 'Tab') { e.preventDefault(); this.commit(); this.move(0, e.shiftKey ? -1 : 1); this.focus(); }
    else if (e.key === 'Escape') { e.preventDefault(); this.cancel(); this.focus(); }
    else if (e.key === 'F4') { e.preventDefault(); const res = cycleRefAt(input.value, input.selectionStart); if (res) { input.value = res.text; this.editing.text = res.text; if (this.fInput && input !== this.fInput) this.fInput.value = res.text; if (this.cellInput && input !== this.cellInput) this.cellInput.value = res.text; input.setSelectionRange(res.caret, res.caret); this.showRefs(); } }
    else if (this.editing.fresh && !this.editing.text.startsWith('=') && e.key.startsWith('Arrow') && input === this.cellInput) { e.preventDefault(); this.commit(); const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]; this.move(d[0], d[1]); this.focus(); }
    else if (e.key === 'Enter' && e.altKey) { /* line break not supported in this grid */ }
  }
  commit() {
    if (!this.editing) return; const { r, c, si, text } = this.editing; this.closeEditor();
    const before = this.wb.input(si, r, c); if (text === before) return;
    if (text.startsWith('=')) { try { tokenize(text); } catch (err) { this.flash('There is a problem with this formula: ' + err.message); } }
    this.record(); this.wb.setInput(si, r, c, text);
    if (this.wb.circular) this.flash('Circular reference: this formula depends on its own cell, so Excel cannot finish the calculation.');
    this.changed({ type: 'edit', cell: addr(r, c), sheet: si, text });
  }
  cancel() { this.closeEditor(); this.updateBars(); this.render(); }
  closeEditor() { if (this.cellInput) { this.cellInput.remove(); this.cellInput = null; } this.editing = null; this.clearRefs(); if (this.sugBox) this.sugBox.style.display = 'none'; if (this.si !== this.selSheetOnEdit && this.selSheetOnEdit !== undefined) {} }
  showRefs() {
    this.clearRefs(); if (!this.editing || !this.editing.text.startsWith('=')) return;
    let toks; try { toks = tokenize(this.editing.text); } catch (e) { return; }
    let k = 0;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i]; if (t.type !== 'ref') continue; let g = { r1: t.r, c1: t.c, r2: t.r, c2: t.c };
      if (toks[i + 1] && toks[i + 1].val === ':' && toks[i + 2] && toks[i + 2].type === 'ref') { const u = toks[i + 2]; g = { r1: Math.min(t.r, u.r), c1: Math.min(t.c, u.c), r2: Math.max(t.r, u.r), c2: Math.max(t.c, u.c) }; i += 2; }
      const sheetIdx = t.sheet ? this.wb.sheetIndexByName(t.sheet) : this.editing.si; if (sheetIdx !== this.si) { k++; continue; }
      const col = REF_COLORS[k++ % REF_COLORS.length]; const b = el('div', 'xg-refbox');
      Object.assign(b.style, { left: this.colX(g.c1) + HEAD_W + 'px', top: g.r1 * this.RH + HEAD_H + 'px', width: this.colX(g.c2 + 1) - this.colX(g.c1) + 'px', height: (g.r2 - g.r1 + 1) * this.RH + 'px', borderColor: col, background: col + '14' });
      this.layer.appendChild(b); this.refOverlays.push(b);
    }
  }
  clearRefs() { this.refOverlays.forEach(b => b.remove()); this.refOverlays = []; }
  suggest() {
    const input = this.editing.via === 'bar' ? this.fInput : this.cellInput; if (!input) return;
    const pos = input.selectionStart; const m = /(?:^=|[=(,+\-*/^&<>\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(input.value.slice(0, pos));
    if (!this.sugBox) { this.sugBox = el('div', 'xg-sug'); this.host.appendChild(this.sugBox); }
    if (!input.value.startsWith('=') || !m || m[1].length < 1) { this.sugBox.style.display = 'none'; return; }
    const list = FUNCTION_NAMES.filter(n => n.startsWith(m[1].toUpperCase())).slice(0, 6);
    if (!list.length) { this.sugBox.style.display = 'none'; return; }
    this.sugWord = m[1]; this.sugActive = list[0];
    this.sugBox.innerHTML = list.map((n, i) => `<div class="${i ? '' : 'on'}" data-n="${n}">${n}</div>`).join('') + '<small>Tab to insert</small>';
    const ir = input.getBoundingClientRect(), hr = this.host.getBoundingClientRect();
    Object.assign(this.sugBox.style, { display: 'block', left: ir.left - hr.left + 'px', top: ir.bottom - hr.top + 2 + 'px' });
    this.sugBox.querySelectorAll('div').forEach(d => d.onmousedown = ev => { ev.preventDefault(); this.sugActive = d.dataset.n; this.acceptSuggest(input); });
  }
  acceptSuggest(input) { const pos = input.selectionStart; const before = input.value.slice(0, pos - this.sugWord.length) + this.sugActive + '('; const v = before + input.value.slice(pos); input.value = v; this.editing.text = v; if (this.fInput) this.fInput.value = v; if (this.cellInput) this.cellInput.value = v; input.setSelectionRange(before.length, before.length); this.sugBox.style.display = 'none'; }

  // ---------- keyboard ----------
  key(e) {
    if (this.editing) return;
    const ctrl = e.ctrlKey || e.metaKey; const k = e.key;
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { e.preventDefault(); const [dr, dc] = arrows[k]; if (ctrl) this.jump(dr, dc, e.shiftKey); else if (e.shiftKey) this.extend(dr, dc); else this.move(dr, dc); this.emitKey(e); return; }
    if (k === 'Enter') { e.preventDefault(); this.move(e.shiftKey ? -1 : 1, 0); return; }
    if (k === 'Tab') { e.preventDefault(); this.move(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Home') { e.preventDefault(); if (ctrl) this.setSel(0, 0); else this.setSel(this.sel.ar, 0); this.emitKey(e); return; }
    if (k === 'End' && ctrl) { e.preventDefault(); const u = this.wb.usedRange(this.sheet) || { r2: 0, c2: 0 }; if (e.shiftKey) this.setSel(this.sel.ar, this.sel.ac, u.r2, u.c2); else this.setSel(u.r2, u.c2); this.emitKey(e); return; }
    if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); if (ctrl || e.altKey) { this.setSheet(Math.max(0, Math.min(this.wb.sheets.length - 1, this.si + (k === 'PageDown' ? 1 : -1)))); } else { const n = Math.floor(this.scroller.clientHeight / this.RH) - 1; this.move(k === 'PageDown' ? n : -n, 0); } this.emitKey(e); return; }
    if (k === 'F2') { e.preventDefault(); this.beginEdit(null, 'cell'); return; }
    if (k === 'F11' && e.shiftKey) { e.preventDefault(); this.addSheet(); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (k === 'Backspace' && !this.multi()) { if (this.canEdit(this.sel.ar, this.sel.ac)) this.beginEdit('', 'cell'); } else this.clearSelection(); return; }
    if (ctrl) {
      const lk = k.toLowerCase();
      if (lk === 'z') { e.preventDefault(); e.shiftKey ? this.redo() : this.undo(); return; }
      if (lk === 'y') { e.preventDefault(); this.redo(); return; }
      if (lk === 'b') { e.preventDefault(); this.toggleStyle('bold'); return; }
      if (lk === 'i') { e.preventDefault(); this.toggleStyle('italic'); return; }
      if (lk === 'd') { e.preventDefault(); this.fillDir('down'); return; }
      if (lk === 'r') { e.preventDefault(); this.fillDir('right'); return; }
      if (lk === 'a') { e.preventDefault(); const g = this.wb.currentRegion(this.si, this.sel.ar, this.sel.ac); const g0 = this.selRange(); if (rangeStr(g) === rangeStr(g0) || this.wb.isEmpty(this.si, this.sel.ar, this.sel.ac)) this.setSel(0, 0, this.sheet.rows - 1, this.sheet.cols - 1); else { this.sel = { ar: g.r1, ac: g.c1, r: g.r2, c: g.c2 }; this.render(); } this.emitKey(e); return; }
      if (k === '`' || k === '~') { e.preventDefault(); this.showFormulas = !this.showFormulas; this.render(); this.emitKey(e); return; }
      if (k === ';') { e.preventDefault(); if (this.canEdit(this.sel.ar, this.sel.ac)) { const d = new Date(); const t = e.shiftKey ? `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}` : `${d.getDate()}-${d.getMonth() + 1}-${d.getFullYear()}`; this.record(); this.wb.setInput(this.si, this.sel.ar, this.sel.ac, t); this.changed({ type: 'edit' }); } return; }
      if (lk === 'f' || lk === 'h') { e.preventDefault(); this.findDialog(lk === 'h'); return; }
      if (lk === 'c' || lk === 'x' || lk === 'v') return; // handled by copy/cut/paste events
      if (k === '1' ) { e.preventDefault(); this.fmtSel && this.fmtSel.focus(); return; }
    }
    if (e.altKey && (k === '=' )) { e.preventDefault(); this.autoSum(); this.emitKey(e); return; }
    if (k.length === 1 && !ctrl && !e.altKey) { e.preventDefault(); if (this.canEdit(this.sel.ar, this.sel.ac)) this.beginEdit(k, 'cell'); }
  }
  emitKey(e) { this.o.onKey && this.o.onKey({ key: e.key, ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey, alt: e.altKey, sel: this.selection() }); }
  multi() { const g = this.selRange(); return g.r1 !== g.r2 || g.c1 !== g.c2; }
  move(dr, dc) { const g = this.selRange(); if (this.multi() && (dr || dc) && false) return; this.setSel(this.sel.ar + dr, this.sel.ac + dc); }
  extend(dr, dc) { this.sel.r = Math.max(0, Math.min(this.sheet.rows - 1, this.sel.r + dr)); this.sel.c = Math.max(0, Math.min(this.sheet.cols - 1, this.sel.c + dc)); this.scrollIntoView(this.sel.r, this.sel.c); this.render(); this.o.onSelect && this.o.onSelect(this.selection()); }
  jump(dr, dc, extend) {
    // Ctrl+arrow: to the edge of the current block of data, or the next filled cell
    const s = this.sheet; let r = extend ? this.sel.r : this.sel.ar, c = extend ? this.sel.c : this.sel.ac;
    const filled = (rr, cc) => rr >= 0 && cc >= 0 && rr < s.rows && cc < s.cols && !this.wb.isEmpty(this.si, rr, cc);
    const maxR = s.rows - 1, maxC = s.cols - 1;
    const step = () => { r += dr; c += dc; };
    if (filled(r, c) && filled(r + dr, c + dc)) { while (filled(r + dr, c + dc)) step(); }
    else { step(); while (r >= 0 && c >= 0 && r <= maxR && c <= maxC && !filled(r, c)) step(); if (r < 0 || c < 0 || r > maxR || c > maxC) { r = Math.max(0, Math.min(maxR, r)); c = Math.max(0, Math.min(maxC, c)); } }
    if (extend) { this.sel.r = r; this.sel.c = c; this.scrollIntoView(r, c); this.render(); } else this.setSel(r, c);
  }
  nameBoxGo(txt) {
    if (!txt) return; const nm = this.wb.names[txt.toLowerCase()];
    if (nm) { const m = /^(?:'?(.+?)'?!)?(.+)$/.exec(nm.ref); if (m[1]) { const i = this.wb.sheetIndexByName(m[1]); if (i >= 0) this.setSheet(i); } this.select(m[2].replace(/\$/g, '')); return; }
    const m = /^(?:'?(.+?)'?!)?(\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(txt);
    if (m) { if (m[1]) { const i = this.wb.sheetIndexByName(m[1]); if (i >= 0) this.setSheet(i); } this.select(m[2].replace(/\$/g, '').toUpperCase()); return; }
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(txt) && !/^[A-Za-z]{1,3}\d+$/.test(txt)) {
      if (this.o.readOnly) return; const g = this.selRange(); const ref = quoteSheet(this.sheet.name) + '!' + rangeAbs(g);
      this.record(); this.wb.defineName(txt, ref); this.flash(`Named ${rangeStr(g)} as “${txt}”`); this.changed({ type: 'name', name: txt, ref }); return;
    }
    this.flash('Type a cell like B4, a range like A1:C10, or a name.');
  }
  // ---------- changes ----------
  record() { this.history.push(this.wb.snapshot()); if (this.history.length > 60) this.history.shift(); this.future = []; }
  undo() { if (!this.history.length) return this.flash('Nothing to undo'); this.future.push(this.wb.snapshot()); this.wb.restore(this.history.pop()); this.si = Math.min(this.si, this.wb.sheets.length - 1); this.changed({ type: 'undo' }); }
  redo() { if (!this.future.length) return; this.history.push(this.wb.snapshot()); this.wb.restore(this.future.pop()); this.changed({ type: 'redo' }); }
  changed(ev) { this.render(); this.o.onChange && this.o.onChange(ev); }
  flash(msg) { this.toast.textContent = msg; this.toast.classList.add('on'); clearTimeout(this._tt); this._tt = setTimeout(() => this.toast.classList.remove('on'), 3200); }
  eachSel(fn) { const g = this.selRange(); const u = this.wb.usedRange(this.sheet); const r2 = Math.min(g.r2, u ? Math.max(u.r2, g.r1) : g.r1), c2 = Math.min(g.c2, u ? Math.max(u.c2, g.c1) : g.c1); for (let r = g.r1; r <= Math.max(r2, g.r1 === g.r2 ? g.r1 : r2); r++) for (let c = g.c1; c <= Math.max(c2, g.c1 === g.c2 ? g.c1 : c2); c++) fn(r, c); }
  clearSelection() { const g = this.selRange(); let blocked = false; this.record(); this.eachSel((r, c) => { if (this.canEdit(r, c, true)) this.wb.clear(this.si, r, c); else blocked = true; }); if (blocked) this.flash(this.o.lockMessage); this.changed({ type: 'clear', range: rangeStr(g) }); }
  applyStyle(st) { this.record(); this.eachSel((r, c) => { if (this.canEdit(r, c, true) || this.o.styleAnywhere) this.wb.setStyle(this.si, r, c, st); }); this.changed({ type: 'style', style: st }); }
  toggleStyle(k) { const cur = !!this.wb.style(this.si, this.sel.ar, this.sel.ac)[k]; this.applyStyle({ [k]: cur ? null : true }); }
  decimals(d) { const cur = this.wb.style(this.si, this.sel.ar, this.sel.ac).fmt || 'General'; let f = cur; if (f === 'General') f = d > 0 ? '0.0' : '0'; else { const m = /(0)(\.0*)?/.exec(f); if (m) { const dec = m[2] ? m[2].length - 1 : 0; const nd = Math.max(0, dec + d); f = f.replace(/0(\.0*)?(?!.*0)/, '0' + (nd ? '.' + '0'.repeat(nd) : '')); } } this.applyStyle({ fmt: f }); }
  palette(anchor) {
    const p = el('div', 'xg-pal'); p.innerHTML = FILLS.map(f => `<button type="button" title="${f || 'No fill'}" style="background:${f || '#fff'}" data-f="${f}">${f ? '' : '×'}</button>`).join('');
    const r = anchor.getBoundingClientRect(), hr = this.host.getBoundingClientRect(); Object.assign(p.style, { left: r.left - hr.left + 'px', top: r.bottom - hr.top + 4 + 'px' }); this.host.appendChild(p);
    p.querySelectorAll('button').forEach(b => b.onclick = () => { this.applyStyle({ fill: b.dataset.f || null }); p.remove(); this.focus(); });
    setTimeout(() => addEventListener('mousedown', function off(ev) { if (!p.contains(ev.target)) { p.remove(); removeEventListener('mousedown', off); } }), 0);
  }
  autoSum() {
    const { ar: r, ac: c } = this.sel; if (!this.canEdit(r, c)) return;
    let r0 = r - 1; while (r0 >= 0 && typeof this.wb.value(this.si, r0, c) === 'number') r0--;
    let f;
    if (r0 < r - 1) f = `=SUM(${addr(r0 + 1, c)}:${addr(r - 1, c)})`;
    else { let c0 = c - 1; while (c0 >= 0 && typeof this.wb.value(this.si, r, c0) === 'number') c0--; f = c0 < c - 1 ? `=SUM(${addr(r, c0 + 1)}:${addr(r, c - 1)})` : '=SUM()'; }
    this.beginEdit(f, 'cell'); if (f === '=SUM()') this.cellInput.setSelectionRange(5, 5);
  }
  sort(desc) {
    const g0 = this.selRange(); let g = this.multi() ? g0 : this.wb.currentRegion(this.si, this.sel.ar, this.sel.ac);
    if (this.o.editable && !this.o.allowSort) { this.flash('Sorting is switched off in this exercise.'); return; }
    if (this.wb.isEmpty(this.si, this.sel.ar, this.sel.ac) && !this.multi()) return this.flash('Click a cell inside your data first.');
    const header = this.wb.guessHeader(this.si, g); this.record(); this.wb.sortRegion(this.si, g, [{ col: this.sel.ac, desc }], header);
    this.flash(`Sorted ${rangeStr(g)} by column ${colName(this.sel.ac)}, ${desc ? 'largest to smallest' : 'smallest to largest'}${header ? ', keeping the header row' : ''}.`);
    this.changed({ type: 'sort', range: rangeStr(g), col: colName(this.sel.ac), desc });
  }
  explain() {
    const f = this.wb.formula(this.si, this.sel.ar, this.sel.ac); if (!f) return this.flash('Select a cell that holds a formula, then press fx? to see how it is calculated.');
    const steps = explainFormula(this.wb, this.si, f); const d = el('div', 'xg-explain');
    d.innerHTML = `<div class="xg-ex-head"><b>How ${addr(this.sel.ar, this.sel.ac)} is calculated</b><button type="button" aria-label="Close">×</button></div><code>${esc(f)}</code><ol>${steps.map(s => `<li><span>${esc(s.label)}</span> <code>${esc(s.text)}</code> → <b>${esc(s.value)}</b></li>`).join('')}</ol>`;
    this.host.appendChild(d); d.querySelector('button').onclick = () => d.remove(); this.o.onExplain && this.o.onExplain(f);
  }
  findDialog(replace) {
    if (this.findEl) this.findEl.remove();
    const d = el('div', 'xg-find'); d.innerHTML = `<div class="xg-ex-head"><b>${replace ? 'Find and replace' : 'Find'}</b><button type="button" class="x" aria-label="Close">×</button></div>
      <label>Find what <input class="fw"></label>${replace ? '<label>Replace with <input class="rw"></label>' : ''}
      <label class="ck"><input type="checkbox" class="mc"> Match case</label><label class="ck"><input type="checkbox" class="me"> Match entire cell contents</label>
      <div class="row"><button type="button" class="nx">Find next</button>${replace ? '<button type="button" class="r1">Replace</button><button type="button" class="ra">Replace all</button>' : ''}</div><small class="msg">Wildcards: * any characters, ? one character</small>`;
    this.host.appendChild(d); this.findEl = d; const q = s => d.querySelector(s); q('.fw').focus();
    const matcher = () => { const pat = q('.fw').value; if (!pat) return null; const flags = q('.mc').checked ? '' : 'i'; let src = pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'); return new RegExp(q('.me').checked ? '^' + src + '$' : src, flags + 'g'); };
    const cells = () => { const u = this.wb.usedRange(this.sheet); const out = []; if (u) for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) { const t = this.wb.input(this.si, r, c); if (t) out.push({ r, c, t }); } return out; };
    q('.x').onclick = () => { d.remove(); this.findEl = null; this.focus(); };
    q('.nx').onclick = () => { const re = matcher(); if (!re) return; const list = cells(); const cur = this.sel.ar * 16384 + this.sel.ac; const hitc = list.find(x => x.r * 16384 + x.c > cur && (re.lastIndex = 0, re.test(String(this.wb.display(this.si, x.r, x.c)) ) || (re.lastIndex = 0, re.test(x.t)))) || list.find(x => (re.lastIndex = 0, re.test(x.t))); if (hitc) this.setSel(hitc.r, hitc.c); else q('.msg').textContent = 'Excel can’t find what you’re looking for.'; };
    if (replace) {
      const doRep = all => { const re = matcher(); if (!re) return; let n = 0; this.record(); for (const x of all ? cells() : [{ r: this.sel.ar, c: this.sel.ac, t: this.wb.input(this.si, this.sel.ar, this.sel.ac) }]) { re.lastIndex = 0; if (!x.t || !re.test(x.t) || !this.canEdit(x.r, x.c, true)) continue; re.lastIndex = 0; const nt = q('.me').checked ? q('.rw').value : x.t.replace(re, q('.rw').value); this.wb.setInput(this.si, x.r, x.c, nt); n++; } q('.msg').textContent = all ? `All done. We made ${n} replacement${n === 1 ? '' : 's'}.` : (n ? 'Replaced.' : 'This cell doesn’t match.'); this.changed({ type: 'replace', count: n }); };
      q('.r1').onclick = () => doRep(false); q('.ra').onclick = () => doRep(true);
    }
  }
  // ---------- sheets ----------
  setSheet(i) { if (i === this.si || i < 0 || i >= this.wb.sheets.length) return; this.si = i; this.sel = { ar: 0, ac: 0, r: 0, c: 0 }; this.scroller.scrollTop = 0; this.scroller.scrollLeft = 0; if (this.editing && this.cellInput) this.placeEditor(); this.render(); this.o.onSelect && this.o.onSelect(this.selection()); this.o.onSheet && this.o.onSheet(i); }
  addSheet() { if (this.o.readOnly || this.o.noAddSheet) return; this.record(); let n = this.wb.sheets.length + 1; while (this.wb.sheet('Sheet' + n)) n++; this.wb.addSheet('Sheet' + n); this.setSheet(this.wb.sheets.length - 1); this.changed({ type: 'addsheet' }); }
  renameSheet(i, name) {
    if (this.wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) return this.flash('That name is already taken.');
    if (/[\\/?*[\]:]/.test(name) || name.length > 31) return this.flash('Sheet names can’t contain \\ / ? * [ ] : and must be 31 characters or fewer.');
    this.record(); const old = this.wb.sheets[i].name; this.wb.sheets[i].name = name;
    const re = new RegExp("(^|[=(,+\\-*/^&<>:\\s])(" + escRe(old) + "|'" + escRe(old.replace(/'/g, "''")) + "')!", 'gi');
    for (const s of this.wb.sheets) for (const o of s.cells.values()) if (o.f) o.f = o.f.replace(re, (m, p) => p + quoteSheet(name) + '!');
    this.wb.invalidate(); this.changed({ type: 'rename', from: old, to: name });
  }
  // ---------- fill ----------
  fillDrag(e) {
    e.preventDefault(); e.stopPropagation(); const g = this.selRange(); let target = g;
    const move = ev => { const p = this.hit(ev); if (!p) return; const down = p.r - g.r2, right = p.c - g.c2, up = g.r1 - p.r, left = g.c1 - p.c; const m = Math.max(down, right, up, left);
      if (m <= 0) target = g; else if (m === down) target = { ...g, r2: p.r }; else if (m === right) target = { ...g, c2: p.c }; else if (m === up) target = { ...g, r1: p.r }; else target = { ...g, c1: p.c };
      const x1 = this.colX(target.c1) + HEAD_W, y1 = target.r1 * this.RH + HEAD_H; Object.assign(this.fillGhost.style, { display: 'block', left: x1 + 'px', top: y1 + 'px', width: this.colX(target.c2 + 1) + HEAD_W - x1 + 'px', height: (target.r2 - target.r1 + 1) * this.RH + 'px' }); this.scrollIntoView(p.r, p.c); };
    const up = ev => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); this.fillGhost.style.display = 'none'; if (rangeStr(target) !== rangeStr(g)) this.fillTo(g, target, { copy: ev.ctrlKey }); };
    addEventListener('mousemove', move); addEventListener('mouseup', up);
  }
  fillTo(src, target, opts = {}) {
    const s = this.sheet; this.record(); let blocked = false;
    const vertical = target.r1 !== src.r1 || target.r2 !== src.r2;
    if (vertical) {
      const forward = target.r2 > src.r2; const n = forward ? target.r2 - src.r2 : src.r1 - target.r1;
      for (let c = src.c1; c <= src.c2; c++) {
        let line = []; for (let r = src.r1; r <= src.r2; r++) line.push(this.wb.cellObj(s, r, c)); if (!forward) line = line.reverse();
        const res = fillSeries(line.map(o => o ? { v: o.v, f: o.f, s: o.s } : null), n, { copy: opts.copy, series: opts.series });
        res.forEach((x, i) => { const r = forward ? src.r2 + 1 + i : src.r1 - 1 - i; if (!this.canEdit(r, c, true)) { blocked = true; return; } const fromR = forward ? src.r1 + x.from : src.r2 - x.from; this.putFill(r, c, x, fromR, c); });
      }
    } else {
      const forward = target.c2 > src.c2; const n = forward ? target.c2 - src.c2 : src.c1 - target.c1;
      for (let r = src.r1; r <= src.r2; r++) {
        let line = []; for (let c = src.c1; c <= src.c2; c++) line.push(this.wb.cellObj(s, r, c)); if (!forward) line = line.reverse();
        const res = fillSeries(line.map(o => o ? { v: o.v, f: o.f, s: o.s } : null), n, { copy: opts.copy });
        res.forEach((x, i) => { const c = forward ? src.c2 + 1 + i : src.c1 - 1 - i; if (!this.canEdit(r, c, true)) { blocked = true; return; } const fromC = forward ? src.c1 + x.from : src.c2 - x.from; this.putFill(r, c, x, r, fromC); });
      }
    }
    this.wb.invalidate(); if (blocked) this.flash(this.o.lockMessage);
    this.sel = { ar: Math.min(src.r1, target.r1), ac: Math.min(src.c1, target.c1), r: Math.max(src.r2, target.r2), c: Math.max(src.c2, target.c2) };
    this.changed({ type: 'fill', from: rangeStr(src), to: rangeStr(target) });
  }
  putFill(r, c, x, fromR, fromC) { const o = this.wb.cellObj(this.sheet, r, c, true); if (x.f) { o.f = shiftFormula(x.f, r - fromR, c - fromC); o.v = null; } else { o.v = x.v ?? null; o.f = null; } o.s = x.s ? { ...x.s } : null; }
  fillDir(dir) { const g = this.selRange(); if (dir === 'down') { if (g.r1 === g.r2) { if (g.r1 === 0) return; this.fillTo({ ...g, r1: g.r1 - 1, r2: g.r1 - 1 }, { ...g, r1: g.r1 - 1 }, { copy: true }); } else this.fillTo({ ...g, r2: g.r1 }, g, { copy: true }); } else { if (g.c1 === g.c2) { if (g.c1 === 0) return; this.fillTo({ ...g, c1: g.c1 - 1, c2: g.c1 - 1 }, { ...g, c1: g.c1 - 1 }, { copy: true }); } else this.fillTo({ ...g, c2: g.c1 }, g, { copy: true }); } }
  // ---------- clipboard ----------
  copy(cut) { const g = this.selRange(); const cells = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const o = this.wb.cellObj(this.sheet, r, c); row.push(o ? { v: o.v, f: o.f, s: o.s ? { ...o.s } : null } : null); } cells.push(row); } this.clip = { g, si: this.si, cells, cut, text: this.selectionTSV() }; this.flash(cut ? 'Cut. Select where to paste and press Ctrl+V.' : 'Copied. Select where to paste and press Ctrl+V.'); this.markCopy(g); }
  markCopy(g) { if (this.ants) this.ants.remove(); this.ants = el('div', 'xg-ants'); Object.assign(this.ants.style, { left: this.colX(g.c1) + HEAD_W + 'px', top: g.r1 * this.RH + HEAD_H + 'px', width: this.colX(g.c2 + 1) - this.colX(g.c1) + 'px', height: (g.r2 - g.r1 + 1) * this.RH + 'px' }); this.layer.appendChild(this.ants); }
  selectionTSV() { const g = this.selRange(); const lines = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 2000); r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(this.wb.display(this.si, r, c)); lines.push(row.join('\t')); } return lines.join('\n'); }
  pasteEvent(e) {
    if (this.editing) return; e.preventDefault(); const text = e.clipboardData.getData('text/plain');
    if (this.clip && (text === this.clip.text || !text)) return this.paste('all');
    if (text) { const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t')); const { ar, ac } = this.sel; this.record(); rows.forEach((row, i) => row.forEach((t, j) => { if (this.canEdit(ar + i, ac + j, true)) this.wb.setInput(this.si, ar + i, ac + j, t); })); this.changed({ type: 'paste' }); }
  }
  paste(mode = 'all') {
    if (!this.clip) return this.flash('Copy some cells first (Ctrl+C).');
    const { cells, g } = this.clip; const { ar, ac } = this.sel; const s = this.sheet; this.record(); let blocked = false;
    const h = cells.length, w = cells[0].length; const tr = mode === 'transpose';
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
      const r = ar + (tr ? j : i), c = ac + (tr ? i : j); if (!this.canEdit(r, c, true)) { blocked = true; continue; } const src = cells[i][j]; const o = this.wb.cellObj(s, r, c, true);
      if (mode === 'formats') { o.s = src && src.s ? { ...src.s } : null; continue; }
      if (mode === 'values') { const v = this.wb.value(this.clip.si, g.r1 + i, g.c1 + j); o.v = v; o.f = null; continue; }
      if (!src) { o.v = null; o.f = null; o.s = null; continue; }
      o.f = src.f ? (this.clip.cut ? src.f : shiftFormula(src.f, r - (g.r1 + i), c - (g.c1 + j))) : null; o.v = src.f ? null : src.v; o.s = src.s ? { ...src.s } : null;
    }
    if (this.clip.cut) { const cs = this.wb.sheets[this.clip.si]; for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) { const r = g.r1 + i, c = g.c1 + j; if (r >= ar && r < ar + h && c >= ac && c < ac + w && this.clip.si === this.si) continue; cs.cells.delete(r * 16384 + c); } this.clip = null; if (this.ants) this.ants.remove(); }
    this.wb.invalidate(); if (blocked) this.flash(this.o.lockMessage);
    this.sel = { ar, ac, r: ar + (tr ? w : h) - 1, c: ac + (tr ? h : w) - 1 }; this.changed({ type: 'paste', mode });
  }
  // ---------- context menu ----------
  contextMenu(e) {
    e.preventDefault(); if (this.o.readOnly) return; if (this.menu) this.menu.remove();
    const m = el('div', 'xg-menu'); const items = [
      ['Cut', () => this.copy(true)], ['Copy', () => this.copy()], ['Paste', () => this.paste('all')], ['Paste values', () => this.paste('values')], ['Paste formats', () => this.paste('formats')], ['Paste transposed', () => this.paste('transpose')], '-',
      ['Insert row above', () => this.structure('insRow')], ['Delete row', () => this.structure('delRow')], ['Insert column to the left', () => this.structure('insCol')], ['Delete column', () => this.structure('delCol')], '-',
      ['Clear contents', () => this.clearSelection()], ['Sort smallest to largest', () => this.sort(false)], ['Sort largest to smallest', () => this.sort(true)], '-',
      ['Define name…', () => { const n = prompt('Name for ' + rangeStr(this.selRange()) + ' (letters, no spaces)'); if (n) this.nameBoxGo(n.trim()); }], ['Explain formula', () => this.explain()]];
    const tn = typeof this.wb.value(this.si, this.sel.ar, this.sel.ac) === 'string' && /^\s*-?[\d,]+(\.\d+)?\s*$/.test(this.wb.value(this.si, this.sel.ar, this.sel.ac));
    if (tn) items.unshift(['Convert to number', () => { this.record(); this.eachSel((r, c) => { const v = this.wb.value(this.si, r, c); if (typeof v === 'string' && this.canEdit(r, c, true)) { const p = parseInput(v.trim()); if (typeof p.v === 'number') { this.wb.setInput(this.si, r, c, v.trim()); const st = this.wb.style(this.si, r, c); if (st.text) this.wb.setStyle(this.si, r, c, { text: null }); } } }); this.changed({ type: 'convert' }); }], '-');
    for (const it of items) { if (it === '-') { m.appendChild(el('hr')); continue; } const b = el('button', '', it[0]); b.type = 'button'; b.onclick = () => { m.remove(); this.menu = null; it[1](); this.focus(); }; m.appendChild(b); }
    const hr = this.host.getBoundingClientRect(); Object.assign(m.style, { left: Math.min(e.clientX - hr.left, hr.width - 220) + 'px', top: e.clientY - hr.top + 'px' }); this.host.appendChild(m); this.menu = m;
    setTimeout(() => addEventListener('mousedown', function off(ev) { if (!m.contains(ev.target)) { m.remove(); removeEventListener('mousedown', off); } }), 0);
  }
  structure(kind) {
    if (this.o.editable && !this.o.allowStructure) return this.flash('Inserting and deleting is switched off in this exercise.');
    const g = this.selRange(); this.record();
    if (kind === 'insRow') this.wb.insertRows(this.si, g.r1, g.r2 - g.r1 + 1); if (kind === 'delRow') this.wb.deleteRows(this.si, g.r1, g.r2 - g.r1 + 1);
    if (kind === 'insCol') this.wb.insertCols(this.si, g.c1, g.c2 - g.c1 + 1); if (kind === 'delCol') this.wb.deleteCols(this.si, g.c1, g.c2 - g.c1 + 1);
    this.changed({ type: kind, range: rangeStr(g) });
  }
  // ---------- scripted demos ----------
  async demoSelect(a, ms = 500) { this.select(a); this.actBox.classList.add('pulse'); await wait(ms); this.actBox.classList.remove('pulse'); }
  async demoType(a, text, speed = 55) {
    const p = parseAddr(a); this.setSel(p.r, p.c); this.editing = { r: p.r, c: p.c, si: this.si, text: '', via: 'cell' };
    this.cellInput = el('input', 'xg-edit'); this.layer.appendChild(this.cellInput); this.placeEditor();
    for (const ch of text) { this.editing.text += ch; this.cellInput.value = this.editing.text; if (this.fInput) this.fInput.value = this.editing.text; this.showRefs(); await wait(speed); }
    await wait(350); const t = this.editing.text; this.closeEditor(); this.wb.setInput(this.si, p.r, p.c, t); this.render(); await wait(300);
  }
  async demoFill(from, to) { const g = parseRange(from), t = parseRange(to); this.select(from); await wait(400); this.fillTo(g, t); await wait(500); }
  async demoStyle(a, st) { this.select(a); await wait(350); this.applyStyle(st); await wait(450); }
}
function fmtStatus(v) { return Math.abs(v) >= 1000 ? v.toLocaleString('en-IN', { maximumFractionDigits: 2 }) : String(+v.toFixed(4)); }
function quoteSheet(n) { return /^[A-Za-z_][A-Za-z0-9_]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
function rangeAbs(g) { const a = '$' + colName(g.c1) + '$' + (g.r1 + 1), b = '$' + colName(g.c2) + '$' + (g.r2 + 1); return a === b ? a : a + ':' + b; }
function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
const wait = ms => new Promise(r => setTimeout(r, ms));
