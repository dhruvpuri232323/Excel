// The study plan builds itself: the student picks a start date, a finish date, their study days and any days
// off, and every remaining lesson, module test, the project and the final exam are spread evenly across the
// days available. The plan is always worked out from today, so a missed day simply moves things along.
import { store, today } from './store.js';
import { lessonList, isDone, passedTest, projectOpen, examReqs, TEST_MINUTES, comingSoon } from './access.js';

export const iso = d => { const x = new Date(d); x.setMinutes(x.getMinutes() - x.getTimezoneOffset()); return x.toISOString().slice(0, 10); };
export const addDays = (s, n) => { const d = new Date(s + 'T00:00:00'); d.setDate(d.getDate() + n); return iso(d); };
export const dayOf = s => new Date(s + 'T00:00:00').getDay();
export const fmtD = (s, o = { day: 'numeric', month: 'short' }) => new Date(s + 'T00:00:00').toLocaleDateString('en-IN', o);
export const PROJECT_MIN = 120, EXAM_MIN = 90;

// everything still to do, in course order
export function planItems(c) {
  const out = [];
  for (const m of c.modules) {
    for (const l of m.lessons) if (!isDone(l.id)) out.push({ key: 'l:' + l.id, kind: 'lesson', id: l.id, short: `${m.n}.${l.num.split('.')[1]}`, title: l.title, mod: m.n, min: +l.minutes || 30, href: `#/lesson/${c.id}/${l.id}` });
    if (!passedTest(c, m.n)) out.push({ key: 't:' + m.n, kind: 'test', short: `Test ${m.n}`, title: `Module ${m.n} test`, mod: m.n, min: TEST_MINUTES, href: `#/mtest/${c.id}/${m.n}` });
  }
  const R = examReqs(c); if (comingSoon(c)) return out; // the project and exam arrive with the course's last part
  if (!R.projOk) out.push({ key: 'p', kind: 'project', short: 'Project', title: 'Course project', min: PROJECT_MIN, href: `#/project/${c.id}` });
  if (!(store.d.examAttempts[c.id] || []).some(a => a.passed)) out.push({ key: 'e', kind: 'exam', short: 'Exam', title: 'Final exam', min: EXAM_MIN, href: `#/exam/${c.id}` });
  return out;
}
const isStudyDay = (p, s) => p.days.includes(dayOf(s)) && !(p.off || []).includes(s);
// spread the items over the study days between start (or today) and the finish date
export function buildPlan(c, p = store.d.plan) {
  const items = planItems(c); const t = today(); const start = p.start && p.start > t ? p.start : t;
  const total = items.reduce((a, x) => a + x.min, 0);
  let end = p.end;
  if (!end) { // no finish date yet: use the daily minutes to work one out
    let need = Math.ceil(total / Math.max(10, p.minutesGoal || 30)), d = start, guard = 0; end = start;
    while (need > 0 && guard++ < 1500 && p.days.length) { if (isStudyDay(p, d)) { need--; end = d; } d = addDays(d, 1); }
  }
  const dates = []; if (p.days.length) for (let d = start, g = 0; d <= end && g < 1500; d = addDays(d, 1), g++) if (isStudyDay(p, d)) dates.push(d);
  const res = { items, total, start, end, dates, days: {}, perDay: 0, fixedEnd: !!p.end, error: null };
  if (!items.length) return res;
  if (!p.days.length) { res.error = 'Pick at least one study day.'; return res; }
  if (!dates.length) { res.error = end < start ? 'The finish date is before the start date.' : 'No study days between those dates. Pick a later finish date or more study days.'; return res; }
  res.perDay = Math.ceil(total / dates.length);
  let cum = 0; for (const it of items) { const k = Math.min(dates.length - 1, Math.floor((cum + it.min / 2) / total * dates.length)); (res.days[dates[k]] = res.days[dates[k]] || []).push(it); it.date = dates[k]; cum += it.min; }
  res.last = items.at(-1).date;
  return res;
}
export const dayMinutes = list => (list || []).reduce((a, x) => a + x.min, 0);
// when the student fixes their dates, remember where each item was meant to be, so we can say on track / behind
export function setBaseline(c) { const P = buildPlan(c); store.d.plan.baseline = { made: today(), at: Object.fromEntries(P.items.map(x => [x.key, x.date])) }; store.d.plan.minutesGoal = Math.max(10, Math.min(180, Math.ceil((P.perDay || 30) / 5) * 5)); store.save(); return P; }
export function planStatus(c) {
  const b = store.d.plan.baseline; if (!b) return null; const t = today(); const left = new Set(planItems(c).map(x => x.key));
  const behind = Object.entries(b.at).filter(([k, d]) => d < t && left.has(k)).length;
  const ahead = Object.entries(b.at).filter(([k, d]) => d > t && !left.has(k)).length;
  return { behind, ahead };
}
export function todayItems(c) { const P = buildPlan(c); return { P, list: P.days[today()] || [], next: P.dates.find(d => d > today()) }; }
