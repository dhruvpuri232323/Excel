// The learning map: every module as a stop, the module tests as diamonds between them, the project and the final exam.
// Wide screens get a winding path; phones get a single column.
import { esc, popover, closePopover, toast } from './ui.js';
import { moduleInfo, passedTest, projectOpen, examReqs, currentModule, nextStep, hasKey, fmtMin, isDone } from './access.js';
import { store } from './store.js';

const wrap = (s, n = 17) => { const w = s.split(' '); const out = ['']; for (const x of w) { if ((out.at(-1) + ' ' + x).trim().length > n && out.at(-1)) out.push(x); else out[out.length - 1] = (out.at(-1) + ' ' + x).trim(); } return out.slice(0, 3); };
const STATUS = { done: 'Done', testedout: 'Tested out', progress: 'In progress', open: 'Open', locked: 'Locked' };
function stops(c) {
  const cur = currentModule(c); const step = nextStep(c);
  const list = c.modules.map(m => { const i = moduleInfo(c, m); let sub = STATUS[i.status]; if (i.status === 'progress' || i.status === 'open') sub = `${i.done} of ${i.total} lessons`; if (i.status === 'done' || i.status === 'testedout') { const b = Math.max(...store.d.tests.filter(t => t.kind === 'module' && t.module === `${c.id}-m${m.n}`).map(t => t.pct)); sub += `, test ${b}%`; } if (i.status === 'locked') sub = i.reason === 'login' ? 'Needs a login' : `After the Module ${i.prev.n} test`;
    return { key: 'm' + m.n, label: String(m.n), name: m.title, sub, status: i.status, href: `#/module/${c.id}/${m.n}`, here: cur === m, info: i, m }; });
  const pOpen = projectOpen(c); const R = examReqs(c); const examDone = (store.d.examAttempts[c.id] || []).some(a => a.passed);
  list.push({ key: 'project', label: 'P', name: 'Course project', sub: R.projOk ? `Passed, ${R.proj.best}%` : pOpen ? 'Open' : `After Module ${c.modules.at(-1).n}`, status: R.projOk ? 'done' : pOpen ? 'open' : 'locked', href: `#/project/${c.id}`, here: step.kind === 'project' });
  list.push({ key: 'exam', label: 'E', name: 'Final exam', sub: examDone ? 'Passed' : R.all ? 'Open' : 'After the project', status: examDone ? 'done' : R.all ? 'open' : 'locked', href: `#/exam/${c.id}`, here: step.kind === 'exam' });
  return list;
}
export function mapSVG(c, width) {
  const S = stops(c); const vertical = width < 620; const cols = width >= 900 ? 5 : 4;
  const pos = []; let W, H;
  if (vertical) { W = 360; S.forEach((s, k) => pos.push([46, 60 + k * 92])); H = 60 + S.length * 92; }
  else { W = 1000; const gap = (W - 180) / (cols - 1); S.forEach((s, k) => { const row = Math.floor(k / cols), col = k % cols; const x = 90 + (row % 2 ? cols - 1 - col : col) * gap; pos.push([x, 96 + row * 176]); }); H = 96 + Math.ceil(S.length / cols) * 176 - 40; }
  let path = '', gates = '';
  for (let k = 0; k < S.length - 1; k++) {
    const [x1, y1] = pos[k], [x2, y2] = pos[k + 1]; const passed = S[k].status === 'done' || S[k].status === 'testedout'; const cls = passed ? 'seg done' : 'seg';
    if (y1 === y2) path += `<path class="${cls}" d="M${x1 + (x2 > x1 ? 32 : -32)} ${y1}H${x2 + (x2 > x1 ? -32 : 32)}"/>`;
    else if (vertical) path += `<path class="${cls}" d="M${x1} ${y1 + 32}V${y2 - 32}"/>`;
    else { const side = x1 > 500 ? 1 : -1; const ex = x1 + side * 60; path += `<path class="${cls}" d="M${x1 + side * 32} ${y1}H${ex}V${y2}H${x2 + side * 32}"/>`; }
    if (S[k].m && k < S.length - 2) { // a module test sits between this module and the next stop
      const gx = y1 === y2 ? (x1 + x2) / 2 : vertical ? x1 : x1 + (x1 > 500 ? 60 : -60), gy = y1 === y2 ? y1 : (y1 + y2) / 2;
      const p = passedTest(c, S[k].m.n); gates += `<a href="#/mtest/${c.id}/${S[k].m.n}" class="gate ${p ? 'passed' : ''}" aria-label="Module ${S[k].m.n} test${p ? ', passed' : ''}"><title>Module ${S[k].m.n} test${p ? ' · passed' : ' · pass 70% to open the next module'}</title><path d="M${gx} ${gy - 11}l11 11l-11 11l-11 -11z"/></a>`;
    }
  }
  const nodes = S.map((s, k) => { const [x, y] = pos[k]; const lines = wrap(s.name, vertical ? 30 : 19); const tx = vertical ? x + 48 : x, anchor = vertical ? 'start' : 'middle', ty = vertical ? y - 4 : y + 54;
    return `<a href="${s.href}" class="stop st-${s.status} ${s.here ? 'here' : ''}" data-stop="${s.key}" aria-label="${esc(s.name)}: ${esc(s.sub)}"><circle cx="${x}" cy="${y}" r="30"/><text class="sl" x="${x}" y="${y + 7}" text-anchor="middle">${s.status === 'done' || s.status === 'testedout' ? '✓' : esc(s.label)}</text>
      <text class="sn" x="${tx}" y="${ty}" text-anchor="${anchor}">${lines.map((l, i) => `<tspan x="${tx}" dy="${i ? 19 : 0}">${esc(l)}</tspan>`).join('')}</text><text class="ss" x="${tx}" y="${ty + lines.length * 19}" text-anchor="${anchor}">${esc(s.sub)}</text>
      ${s.here ? (vertical ? `<text class="here-l" x="${x + 48}" y="${y - 26}">▼ You are here</text>` : `<text class="here-l" x="${x}" y="${y - 44}" text-anchor="middle">You are here ▼</text>`) : ''}</a>`; }).join('');
  return `<svg class="lmap ${vertical ? 'vert' : ''}" viewBox="0 0 ${W} ${H}" role="group" aria-label="Learning map for ${esc(c.title)}">${path}${gates}${nodes}</svg>`;
}
export function mountMap(el, c) {
  const draw = () => { el.innerHTML = mapSVG(c, el.clientWidth || innerWidth); wireHover(el, c); };
  draw(); let w = el.clientWidth; celebrate(el, c); const ro = new ResizeObserver(() => { if (Math.abs(el.clientWidth - w) > 40) { w = el.clientWidth; draw(); } }); ro.observe(el);
}
function wireHover(el, c) {
  el.querySelectorAll('.stop').forEach(a => {
    const show = () => { const k = a.dataset.stop; if (!k.startsWith('m')) return; const m = c.modules.find(x => 'm' + x.n === k); const i = moduleInfo(c, m);
      popover(a, `<b>Module ${m.n}: ${esc(m.title)}</b><ul class="pop-list">${m.lessons.map(l => `<li class="${isDone(l.id) ? 'ok' : ''}">${isDone(l.id) ? '✓' : '○'} ${esc(l.title)} <small>${l.minutes} min</small></li>`).join('')}</ul><p class="small muted">${i.passed ? 'Module test passed.' : i.open ? `${fmtMin(i.minutesLeft)} left, including the test.` : i.reason === 'login' ? 'Needs a student login.' : `Opens when you pass the Module ${i.prev.n} test, or test out.`}</p>`); };
    a.addEventListener('mouseenter', () => { if (matchMedia('(hover: hover)').matches) show(); }); a.addEventListener('mouseleave', () => setTimeout(() => { if (!document.querySelector('.popover:hover')) closePopover(); }, 250)); a.addEventListener('focus', show);
  });
}
// a compact strip of module dots for the home page
export function miniMap(c) {
  const cur = currentModule(c);
  return `<div class="minimap" aria-label="Modules">${c.modules.map(m => { const i = moduleInfo(c, m); return `<a href="#/module/${c.id}/${m.n}" class="mm st-${i.status} ${cur === m ? 'here' : ''}" title="Module ${m.n}: ${esc(m.title)} · ${STATUS[i.status]}">${i.status === 'done' || i.status === 'testedout' ? '✓' : m.n}</a>`; }).join('<i></i>')}</div>`;
}
// U1: the first time the map shows a module that just opened, it pulses and says so
function celebrate(el, c) {
  if (store.preview) return; const open = c.modules.filter(m => !m.free && moduleInfo(c, m).open).map(m => c.id + '-m' + m.n);
  const seen = store.d.unlockSeen; if (!seen) { store.d.unlockSeen = open; store.save(); return; }
  const fresh = open.filter(k => !seen.includes(k)); if (!fresh.length) return;
  store.d.unlockSeen = seen.concat(fresh); store.save();
  fresh.forEach(k => { const n = k.split('-m')[1]; const s = el.querySelector(`.stop[data-stop="m${n}"]`); if (s) s.classList.add('just'); });
  const n = fresh.map(k => k.split('-m')[1]); toast(`🔓 Module ${n.join(' and ')} ${n.length > 1 ? 'are' : 'is'} now open.`);
}
