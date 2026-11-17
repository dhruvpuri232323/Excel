// The Office Scripts lab: Excel for the web's Automate tab, the sheet, and the Code Editor pane beside it.
// Scripts run on the lab's workbook with oscript.js; checks run them again on fresh copies (oscheck in oscript.js).
import { Workbook } from './engine.js';
import { Grid } from './grid.js';
import { runScript } from './oscript.js';

const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const NEW = `function main(workbook: ExcelScript.Workbook) {\n  // Your code here\n}\n`;
// TypeScript colouring for the static code view (pictures) — the live editor is a plain textarea over a highlighted layer
const OS_KW = /\b(function|let|const|var|for|of|in|if|else|return|while|do|new|true|false|null|undefined|break|continue|interface|type|as|switch|case|default)\b/g;
export function osHL(line) { const parts = []; let rest = line; const re = /(\/\/.*$)|("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|`(?:[^`\\]|\\.)*`)/; let m;
  while ((m = re.exec(rest))) { parts.push([rest.slice(0, m.index), 0]); parts.push([m[0], m[1] ? 'c' : 's']); rest = rest.slice(m.index + m[0].length); } parts.push([rest, 0]);
  return parts.map(([t, k]) => k === 'c' ? `<span class="oc">${esc(t)}</span>` : k === 's' ? `<span class="os">${esc(t)}</span>` : esc(t).replace(OS_KW, '<span class="ok">$1</span>').replace(/\b(ExcelScript)\b/g, '<span class="ot">$1</span>')).join(''); }
export function osCodeHTML(code, o = {}) { const lines = String(code).replace(/\n$/, '').split('\n'); const mark = new Set(o.mark || []);
  return `<div class="osc">${lines.map((l, i) => `<div class="osl${mark.has(i + 1) ? ' mk' : ''}${o.err === i + 1 ? ' er' : ''}" data-z="osline:${i + 1}"><span class="osn">${i + 1}</span><span class="ost">${osHL(l) || '&nbsp;'}</span></div>`).join('')}</div>`; }
// the parameters main takes after the workbook (what a flow fills in): [{name, type, opt}]
export function osParams(code) { const m = /function\s+main\s*\(([^)]*)\)/.exec(String(code || '')); if (!m) return []; return m[1].split(',').slice(1).map(x => x.trim()).filter(Boolean).map(x => { const k = /^(\w+)\s*(\?)?\s*(?::\s*([^=]+?))?\s*(?:=\s*(.+))?$/.exec(x) || []; return { name: k[1] || x, type: (k[3] || 'string').trim(), opt: !!(k[2] || k[4]) }; }); }
const showRet = v => typeof v === 'string' ? JSON.stringify(v) : (() => { try { return JSON.stringify(v); } catch (e) { return String(v); } })();
const parsHTML = (pars, live) => pars.length ? `<div class="osp-par" data-z="osparams"><span class="mini">main also takes:</span>${pars.map((p, i) => `<label>${esc(p.name)} <i>${esc(p.type)}</i><input data-par="${i}" value="${esc(p.value ?? '')}" spellcheck="false" ${live ? '' : 'readonly'}></label>`).join('')}</div>` : '';
// the Code Editor pane (also used in pictures)
export function osPaneHTML(o) { const out = o.out || []; const err = o.error; const pars = o.params || [];
  return `<div class="ospane" data-z="ospane"><div class="osp-h"><b>Code Editor</b><span class="x">✕</span></div>
  <div class="osp-bar"><span class="osp-name" data-z="osname">${esc(o.name || 'Script')}</span><button type="button" class="osrun" data-os="run" data-z="osrun">▷ Run</button><button type="button" class="osmore" data-os="scripts" title="All Scripts">☰</button></div>
  ${parsHTML(pars, o.live)}
  <div class="osp-code" data-z="oscode">${o.live ? `<div class="osed"><pre class="oshl" aria-hidden="true"></pre><div class="osgut" aria-hidden="true"></div><textarea class="osta" spellcheck="false" autocapitalize="off" autocomplete="off" aria-label="Script code">${esc(o.code)}</textarea></div>` : osCodeHTML(o.code || NEW, { mark: o.mark, err: err && err.line })}</div>
  <div class="osp-out" data-z="osout"><div class="osp-oh">Output${o.ran ? ` <span class="osok">${err ? '✕ The script stopped' : '✓ Ran'}</span>` : ''}</div><div class="osp-ol">${out.map(l => `<div>${esc(l)}</div>`).join('')}${err ? `<div class="oserr">${err.line ? `Line ${err.line}: ` : ''}${esc(err.message)}</div>` : ''}${o.ret !== undefined ? `<div class="osret">↩ main gave back ${typeof o.ret === 'string' && o.ret.includes('\n') ? 'this text' : esc(showRet(o.ret))} <span class="mini">(what a flow gets as the script’s result)</span></div>${typeof o.ret === 'string' && o.ret.includes('\n') ? `<div class="osretb">${esc(o.ret)}</div>` : ''}` : ''}${!out.length && !err && o.ret === undefined ? '<div class="mini">console.log output appears here.</div>' : ''}</div></div></div>`; }
const RIBBON = `<div class="osrib" data-z="osribbon"><div class="ostabs">${['File', 'Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate', 'Help'].map(t => `<span class="${t === 'Automate' ? 'on' : ''}" data-z="ostab:${t}">${t}</span>`).join('')}</div>
  <div class="osgroups"><button type="button" data-os="new" data-z="osnew"><i>✎</i>New Script</button><button type="button" data-os="scripts" data-z="osall"><i>☰</i>All Scripts</button><button type="button" data-os="record" data-z="osrecord"><i>●</i>Record Actions</button><span class="ossep"></span><button type="button" data-os="flow" data-z="osflow"><i>⚡</i>Automate a Task</button></div></div>`;

export class OsLab {
  get isOs() { return true; }
  constructor(el, lab, state, save) { this.el = el; this.host = el; this.lab = lab; this.v = lab.os || {}; this.saveFn = save;
    const st = state && state.scripts ? state : { scripts: (this.v.scripts || []).map(s => ({ ...s })), cur: (this.v.scripts || [])[0] ? this.v.scripts[0].name : null, sheet: null, active: 0 };
    this.st = st; this.wb = new Workbook(JSON.parse(JSON.stringify(st.sheet || lab.sheet))); this.out = []; this.err = null; this.ran = false; this.ret = undefined; this.render(); }
  ctx() { this.flush(); return { lab: this.lab, state: JSON.parse(JSON.stringify({ ...this.st, sheet: this.wb.toSpec() })) }; }
  save() { this.st.sheet = this.wb.toSpec(); this.saveFn && this.saveFn(this.st); }
  script() { return this.st.scripts.find(s => s.name === this.st.cur) || null; }
  flush() { const s = this.script(); if (s && this.ta) s.code = this.ta.value; }
  render() { const s = this.script(); const h = this.v.height || 560;
    this.el.innerHTML = `<div class="oslab vlab" style="--osh:${h}px"><div class="oswin"><div class="ostitle"><span class="osapp">Excel</span><span>${esc(this.v.book || 'Nilgiri daily report.xlsx')}</span><span class="osnote">Excel for the web, in Sheetah</span></div>${RIBBON}
      <div class="osbody"><div class="osgrid"></div>${s ? osPaneHTML({ name: s.name, code: s.code, out: this.out, error: this.err, ran: this.ran, live: true, ret: this.ret, params: osParams(s.code).map((p, i) => ({ ...p, value: (this.parVals || {})[s.name + '|' + p.name] ?? '' })) }) : `<div class="ospane osempty"><div class="osp-h"><b>Code Editor</b></div><p>No script open. Click <b>New Script</b> on the Automate tab, or <b>All Scripts</b> to open one.</p></div>`}</div><div class="osmsg" data-msg></div></div></div>`;
    const narrow = this.el.clientWidth && this.el.clientWidth < 760; this.grid = new Grid(this.el.querySelector('.osgrid'), this.wb, { height: narrow ? 250 : h - 110, styleAnywhere: true, allowStructure: true, allowSort: true, onChange: () => this.save() });
    this._fitTries = 0; if (!narrow) this.fit();
    if (this.st.active && this.st.active < this.wb.sheets.length) this.grid.setSheet(this.st.active);
    this.el.querySelectorAll('[data-os]').forEach(b => b.onclick = () => this.cmd(b.dataset.os));
    const x = this.el.querySelector('.ospane .x'); if (x) x.onclick = () => { this.flush(); this.st.cur = null; this.save(); this.render(); };
    const nm = this.el.querySelector('.osp-name'); if (nm) { nm.ondblclick = () => this.rename(); nm.title = 'Double-click to rename'; }
    this.wirePars();
    this.ta = this.el.querySelector('.osta'); if (this.ta) this.wireEditor(); }
  fit() { const gEl = this.el.querySelector('.osgrid'); if (!gEl) return; if (!gEl.isConnected || !gEl.clientHeight) { if (++this._fitTries < 120 && typeof requestAnimationFrame === 'function') requestAnimationFrame(() => this.fit()); return; }
    const over = gEl.scrollHeight - gEl.clientHeight; if (over > 0 && this.grid.o && this.grid.o.height) { this.grid.o.height = Math.max(140, this.grid.o.height - over); if (this.grid.view) this.grid.view.style.height = this.grid.o.height + 'px'; this.grid.render(); } }
  parsNow() { const s = this.script(); return s ? osParams(s.code).map(p => ({ ...p, value: (this.parVals || {})[s.name + '|' + p.name] ?? '' })) : []; }
  wirePars() { this._parSig = JSON.stringify(this.parsNow().map(p => [p.name, p.type])); this.el.querySelectorAll('[data-par]').forEach(inp => inp.oninput = () => { const s2 = this.script(); const p = osParams(s2.code)[+inp.dataset.par]; if (!p) return; this.parVals = this.parVals || {}; this.parVals[s2.name + '|' + p.name] = inp.value; }); }
  repaintPars() { const pars = this.parsNow(); const sig = JSON.stringify(pars.map(p => [p.name, p.type])); if (sig === this._parSig) return; const old = this.el.querySelector('.osp-par'); const code = this.el.querySelector('.osp-code'); if (old) old.remove(); if (pars.length && code) code.insertAdjacentHTML('beforebegin', parsHTML(pars, true)); this.wirePars(); }
  wireEditor() { const ta = this.ta, hl = this.el.querySelector('.oshl'), gut = this.el.querySelector('.osgut');
    const paint = () => { const lines = ta.value.split('\n'); hl.innerHTML = lines.map(l => osHL(l) || ' ').join('\n') + '\n'; gut.innerHTML = lines.map((_, i) => `<div class="${this.err && this.err.line === i + 1 ? 'er' : ''}">${i + 1}</div>`).join(''); hl.scrollTop = ta.scrollTop; hl.scrollLeft = ta.scrollLeft; gut.scrollTop = ta.scrollTop; };
    paint(); ta.oninput = () => { paint(); this.flush(); this.repaintPars(); clearTimeout(this._sv); this._sv = setTimeout(() => this.save(), 400); }; ta.onscroll = () => { hl.scrollTop = ta.scrollTop; hl.scrollLeft = ta.scrollLeft; gut.scrollTop = ta.scrollTop; };
    ta.onkeydown = e => { if (e.key === 'Tab') { e.preventDefault(); const a = ta.selectionStart; ta.setRangeText('  ', a, ta.selectionEnd, 'end'); paint(); this.flush(); }
      if (e.key === 'Enter' && !e.ctrlKey) { const a = ta.selectionStart; const line = ta.value.slice(0, a).split('\n').pop(); const ind = /^\s*/.exec(line)[0] + (/[{(\[]\s*$/.test(line) ? '  ' : ''); e.preventDefault(); ta.setRangeText('\n' + ind, a, ta.selectionEnd, 'end'); paint(); this.flush(); }
      if (e.key === '}' && /^\s+$/.test(ta.value.slice(0, ta.selectionStart).split('\n').pop())) { const a = ta.selectionStart; const line = ta.value.slice(0, a).split('\n').pop(); const back = Math.min(2, line.length); e.preventDefault(); ta.setRangeText('}', a - back, ta.selectionEnd, 'end'); paint(); this.flush(); }
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); this.run(); } }; }
  msg(t, bad) { const m = this.el.querySelector('[data-msg]'); if (!m) return; m.textContent = t; m.className = 'osmsg' + (bad ? ' bad' : ''); clearTimeout(this._mt); this._mt = setTimeout(() => { if (m) m.textContent = ''; }, 6000); }
  cmd(k) { this.flush(); if (k === 'run') return this.run(); if (k === 'new') return this.newScript(); if (k === 'scripts') return this.allScripts();
    if (k === 'record') return this.msg('Record Actions writes a script as you click; in this lab you write the script yourself.'); if (k === 'flow') return this.msg('Automate a Task builds a Power Automate flow that runs a script (Lesson 9.2). It needs Microsoft 365, so it isn’t in this lab.'); }
  newScript() { let k = 1; while (this.st.scripts.some(s => s.name === 'Script ' + k)) k++; const s = { name: 'Script ' + k, code: NEW }; this.st.scripts.push(s); this.st.cur = s.name; this.out = []; this.err = null; this.ran = false; this.ret = undefined; this.save(); this.render(); this.msg('New script: double-click its name to rename it.'); }
  rename() { const s = this.script(); const nm = this.el.querySelector('.osp-name'); if (!s || !nm) return; const inp = document.createElement('input'); inp.className = 'osp-rn'; inp.value = s.name; inp.setAttribute('aria-label', 'Script name'); nm.replaceWith(inp); inp.focus(); inp.select();
    let done = false; const commit = ok => { if (done) return; done = true; const t = inp.value.trim(); if (ok && t && t !== s.name) { if (this.st.scripts.some(x => x !== s && x.name.toLowerCase() === t.toLowerCase())) { this.render(); this.msg(`There’s already a script called ${t}.`, true); return; } s.name = t; this.st.cur = t; this.save(); } this.render(); };
    inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); commit(true); } if (e.key === 'Escape') commit(false); }; inp.onblur = () => commit(true); }
  allScripts() { const sh = document.createElement('div'); sh.className = 'vshade'; sh.innerHTML = `<div class="vmb osall"><div class="vmb-t"><span>All Scripts</span><span class="x">✕</span></div><div class="vmac-b"><div class="vmac-list">${this.st.scripts.map(s => `<div data-sn="${esc(s.name)}" class="${s.name === this.st.cur ? 'on' : ''}">📜 ${esc(s.name)}</div>`).join('') || '<div class="mini">No scripts yet.</div>'}</div></div><div class="vmb-f"><button type="button" data-mb="new">New Script</button><button type="button" data-mb="ren">Rename…</button><button type="button" data-mb="del">Delete</button><button type="button" data-mb="close">Close</button></div></div>`;
    this.el.querySelector('.oswin').appendChild(sh); let pick = this.st.cur;
    sh.querySelectorAll('[data-sn]').forEach(d => { d.onclick = () => { sh.querySelectorAll('[data-sn]').forEach(x => x.classList.remove('on')); d.classList.add('on'); pick = d.dataset.sn; }; d.ondblclick = () => { this.st.cur = d.dataset.sn; this.out = []; this.err = null; this.ran = false; this.ret = undefined; this.save(); this.render(); }; });
    sh.querySelector('.x').onclick = () => sh.remove();
    sh.querySelectorAll('[data-mb]').forEach(b => b.onclick = () => { const k = b.dataset.mb; sh.remove(); if (k === 'new') this.newScript(); else if (k === 'ren') { this.st.cur = pick; this.render(); this.rename(); } else if (k === 'del') { if (!pick) return; this.st.scripts = this.st.scripts.filter(s => s.name !== pick); if (this.st.cur === pick) this.st.cur = this.st.scripts[0] ? this.st.scripts[0].name : null; this.save(); this.render(); } else if (pick && pick !== this.st.cur) { this.st.cur = pick; this.out = []; this.err = null; this.ran = false; this.ret = undefined; this.save(); this.render(); } }); }
  run() { const s = this.script(); if (!s) { this.msg('Open or make a script first.', true); return; }
    const sel = this.grid ? { si: this.grid.si, g: this.grid.selRange() } : null;
    const args = osParams(s.code).map(p => { const v = (this.parVals || {})[s.name + '|' + p.name]; if (v === undefined || v === '') return undefined; if (/^number$/i.test(p.type)) return Number(v); if (/^boolean$/i.test(p.type)) return /^true$/i.test(v); return v; });
    const r = runScript(s.code, this.wb, { args, active: this.grid ? this.grid.si : 0, sel: sel && { s: this.wb.sheets[sel.si], g: sel.g }, name: this.v.book });
    this.out = r.out; this.err = r.error; this.ran = true; this.ret = r.error ? undefined : r.value; this.st.active = r.active >= 0 ? r.active : 0; this.save(); this.render(); this.msg(r.error ? 'The script stopped: see Output.' : 'Script ran.', !!r.error); }
}

// ---------------------------------------------------------------- a still picture: Excel for the web with the Code Editor (the sheet is drawn by the shot tool)
export { RIBBON as OS_RIBBON };
