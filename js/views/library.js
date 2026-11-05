// Free tools for everyone: the function library, the shortcut trainer and the playground.
import { register } from '../app.js';
import { store } from '../store.js';
import { getJSON } from '../content.js';
import { esc, md, $, $$, mascot, toast, download, shuffle } from '../ui.js';
import { Grid } from '../grid.js';
import { Workbook } from '../engine.js';
import { readXlsx, toWorkbook } from '../xlsxcheck.js';
import { workbookToXlsx } from '../xlsxwrite.js';

register('library', async (main, name) => {
  const lib = await getJSON('content/library.json'); const fns = lib.functions;
  if (name) {
    const f = fns.find(x => x.name.toLowerCase() === name.toLowerCase()); if (!f) throw new Error('Function not found');
    main.innerHTML = `<a class="crumb" href="#/library">← Function library</a><section class="card"><p class="eyebrow">${esc(f.cat)} · taught in ${esc(f.course)}</p><h1><code>${esc(f.name)}</code></h1><p class="lead">${md(f.what)}</p><p><code class="big">${esc(f.syntax)}</code></p>
     <table class="tbl"><thead><tr><th>Argument</th><th>What it is</th></tr></thead><tbody>${f.args.map(([a, d]) => `<tr><td><code>${esc(a)}</code></td><td>${md(d)}</td></tr>`).join('')}</tbody></table>
     <h2>Try it</h2><p>${md(f.example.say)}</p><div id="tg"></div>
     <h2>Common mistakes</h2><ul class="ticks warn">${f.mistakes.map(m => `<li>${md(m)}</li>`).join('')}</ul>${f.related && f.related.length ? `<p>Related: ${f.related.map(r => `<a href="#/library/${r}">${r}</a>`).join(', ')}</p>` : ''}</section>`;
    new Grid($('#tg'), new Workbook(f.example.sheet), { height: 240, select: f.example.cell, tabs: false }); return;
  }
  const cats = [...new Set(fns.map(f => f.cat))];
  main.innerHTML = `<h1>Function library</h1><p class="lead">Every function in the course, with what it does, a live grid to try it in, and the mistakes people make. Free for everyone.</p><input class="search big-input" id="q" placeholder="Search functions, e.g. average" aria-label="Search functions">
   ${cats.map(c => `<section class="card fcat"><h2>${esc(c)}</h2><div class="fgrid">${fns.filter(f => f.cat === c).map(f => `<a class="fcard" href="#/library/${f.name}" data-q="${esc((f.name + ' ' + f.what + ' ' + f.cat).toLowerCase())}"><code>${esc(f.name)}</code><span>${md(f.short)}</span></a>`).join('')}</div></section>`).join('')}
   <section class="card"><h2>Glossary</h2><dl class="gloss">${lib.glossary.map(([t, d]) => `<dt>${esc(t)}</dt><dd>${md(d)}</dd>`).join('')}</dl></section>`;
  $('#q').oninput = e => { const v = e.target.value.toLowerCase(); $$('.fcard').forEach(a => a.hidden = !a.dataset.q.includes(v)); $$('.fcat').forEach(s => s.hidden = !$$('.fcard', s).some(a => !a.hidden)); };
});

register('shortcuts', async main => {
  const lib = await getJSON('content/library.json'); const os = store.d.settings.os; const S = lib.shortcuts; const cats = [...new Set(S.map(s => s.cat))];
  main.innerHTML = `<h1>Shortcuts</h1><p class="lead">The keys that save the most time, for ${os === 'mac' ? 'Mac' : 'Windows'} (change this in Settings). Practise them below: press the real keys and Sheetah checks them.</p>
   <section class="card trainer"><div class="row between"><h2>Shortcut trainer</h2><span id="ts" class="muted"></span></div><div id="tbox" class="tbox" tabindex="0" aria-live="polite"></div><div class="row"><button class="btn" id="tstart">Start a round of 10</button><button class="btn ghost" id="tskip" hidden>Show me</button></div><p class="muted small">Some shortcuts belong to the browser (like Ctrl+T or Ctrl+W), so they’re shown as flashcards instead of keypresses. They work in Excel.</p></section>
   ${cats.map(c => `<section class="card"><h2>${esc(c)}</h2><table class="keys"><tbody>${S.filter(s => s.cat === c).map(s => `<tr><td>${md('[[' + (os === 'mac' ? s.mac : s.win) + ']]')}</td><td>${md(s.what)}</td></tr>`).join('')}</tbody></table></section>`).join('')}`;
  let round = [], i = 0, right = 0; const box = $('#tbox');
  const norm = e => [e.ctrlKey && 'ctrl', e.metaKey && 'cmd', e.altKey && 'alt', e.shiftKey && 'shift', e.key.length === 1 ? e.key.toLowerCase() : e.key.toLowerCase()].filter(Boolean).filter((v, k, a) => !(['control', 'shift', 'alt', 'meta'].includes(v))).join('+');
  const show = () => { if (i >= round.length) { box.innerHTML = `<div class="say">${mascot(right >= 8 ? 'celebrate' : 'encourage', 'sm')}<div class="bubble">${right} of ${round.length}. ${right >= 8 ? 'Fast fingers!' : 'Run another round and they’ll stick.'}</div></div>`; $('#tskip').hidden = true; return; } const s = round[i]; box.innerHTML = `<p class="eyebrow">${i + 1} of ${round.length}</p><h2>${md(s.what)}</h2><p class="muted">${s.practise ? 'Press the keys now.' : 'This one is taken by the browser. Say it, then reveal.'}</p>${s.practise ? '' : '<button class="btn sm" id="rv">Reveal</button>'}`; $('#tskip').hidden = false; box.focus(); const rv = $('#rv'); if (rv) rv.onclick = () => { box.innerHTML += `<p class="answer">${md('[[' + (os === 'mac' ? s.mac : s.win) + ']]')}</p><div class="row"><button class="btn ghost sm" id="no">I didn’t know</button><button class="btn sm" id="yes">I knew it</button></div>`; $('#yes').onclick = () => { right++; i++; show(); }; $('#no').onclick = () => { i++; show(); }; }; $('#ts').textContent = `${right} right`; };
  box.onkeydown = e => { const s = round[i]; if (!s || !s.practise) return; if (['Control', 'Shift', 'Alt', 'Meta'].includes(e.key)) return; e.preventDefault(); const got = norm(e); const want = (os === 'mac' ? s.comboMac || s.combo : s.combo); if (got === want || (os === 'mac' && got.replace('cmd', 'ctrl') === s.combo)) { right++; box.innerHTML = `<p class="ok big">✓ ${md('[[' + (os === 'mac' ? s.mac : s.win) + ']]')}</p>`; i++; setTimeout(show, 700); } else { box.querySelector('.muted').innerHTML = `You pressed <code>${esc(got)}</code>. Try again, or click “Show me”.`; } };
  $('#tstart').onclick = () => { round = shuffle(S).slice(0, 10); i = 0; right = 0; show(); };
  $('#tskip').onclick = () => { const s = round[i]; box.innerHTML = `<p class="answer">${md(s.what)}: ${md('[[' + (os === 'mac' ? s.mac : s.win) + ']]')}</p>`; i++; setTimeout(show, 1600); };
});

register('playground', async main => {
  const lib = await getJSON('content/library.json');
  main.innerHTML = `<h1>Playground</h1><p class="lead">A free sheet to experiment in. Try any formula, press <b>fx?</b> to see how it’s calculated, open your own Excel file, or download your work as .xlsx.</p>
  <div class="row"><select id="ds" aria-label="Sample data">${lib.playground.map((p, i) => `<option value="${i}">${esc(p.name)}</option>`).join('')}</select><label class="btn ghost">Open an Excel file<input type="file" accept=".xlsx" hidden id="op"></label><button class="btn ghost" id="sv">⬇ Download as .xlsx</button></div><div id="pg"></div>`;
  let g; const load = spec => { if (g) g.destroy(); g = new Grid($('#pg'), new Workbook(spec), { height: 520, allowSort: true, allowStructure: true }); };
  load(lib.playground[0].sheet); $('#ds').onchange = e => load(lib.playground[+e.target.value].sheet);
  $('#op').onchange = async e => { const f = e.target.files[0]; if (!f) return; const x = await readXlsx(f); if (x.error) return toast(esc(x.error)); const wb = toWorkbook(x); if (g) g.destroy(); g = new Grid($('#pg'), wb, { height: 520, allowSort: true, allowStructure: true }); toast(`Opened ${esc(f.name)}. Functions Sheetah doesn’t know yet show Excel’s saved result.`); };
  $('#sv').onclick = async () => download('sheetah-playground.xlsx', await workbookToXlsx(g.wb));
});
