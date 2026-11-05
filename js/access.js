// Who can open what. Free modules are open to everyone. With a login, modules open one at a time:
// pass a module's test and the next one opens, or pass a locked module's test to "test out" of it.
import { session } from './content.js';
import { store } from './store.js';
import { site } from './site.js';

export const PARTS = ['lecture', 'demo', 'lab', 'quiz', 'recap'];
export const PART_NAMES = ['Lecture', 'Demo', 'Lab', 'Quiz', 'Recap'];
const SHARE = [0.3, 0.15, 0.35, 0.12, 0.08];
const CUM = SHARE.reduce((a, x) => a.concat(a[a.length - 1] + x), [0]);
// minutes per part, rounded so the five parts add up to the lesson's minutes
export const partMinutes = (lesson, i) => { const m = +lesson.minutes || 30; return Math.max(1, Math.round(m * CUM[i + 1]) - Math.round(m * CUM[i])); };
export const TEST_MINUTES = 25;

export function lessonList(c) { const out = []; for (const m of c.modules) for (const l of m.lessons) out.push({ ...l, module: m }); return out; }
export const lessonState = id => store.d.lessons[id] || null;
export const isDone = id => !!(store.d.lessons[id] && store.d.lessons[id].status === 'done');
export const lessonNo = l => l.num ? +String(l.num).split('.')[1] : 1;
export const lessonLabel = (l, m) => `Module ${(m || l.module).n} · Lesson ${lessonNo(l)}`;

// "visitor" | "student" | "admin", or the mode an admin is previewing
export const who = () => store.preview ? (store.preview.mode === 'visitor' ? 'visitor' : 'student') : session.isAdmin ? 'admin' : session.signedIn ? 'student' : 'visitor';
export const hasKey = c => who() !== 'visitor' && session.has(c.id);
// the teacher can open every module for one student (by their login's id) or for everyone, from the Admin page
export const override = () => who() === 'student' && !store.preview && (!!(session.s && session.s.open) || !!site.data.openAll || (site.data.openFor || []).includes(session.s && session.s.id));

export const testsFor = (c, n) => store.d.tests.filter(t => t.kind === 'module' && t.module === `${c.id}-m${n}`);
export const passedTest = (c, n) => testsFor(c, n).some(t => t.passed);

export function moduleInfo(c, m) {
  const idx = c.modules.indexOf(m); const prev = idx > 0 ? c.modules[idx - 1] : null;
  const done = m.lessons.filter(l => isDone(l.id)).length, total = m.lessons.length;
  const touched = m.lessons.some(l => store.d.lessons[l.id] && store.d.lessons[l.id].status && store.d.lessons[l.id].status !== 'new');
  const passed = passedTest(c, m.n);
  let open = false, reason = null;
  if (m.free) open = true;
  else if (!hasKey(c)) reason = 'login';
  else if (who() === 'admin' || override() || passed || !prev || passedTest(c, prev.n) || touched) open = true;
  else reason = 'test';
  const testedOut = passed && done < total;
  const status = passed ? (testedOut ? 'testedout' : 'done') : !open ? 'locked' : (done || touched) ? 'progress' : 'open';
  const minutesLeft = m.lessons.filter(l => !isDone(l.id)).reduce((a, l) => a + lessonMinutesLeft(l), 0) + (passed ? 0 : TEST_MINUTES);
  // what's left before the previous module's test, for "you're 1 lesson away" messages
  const prevLeft = prev ? prev.lessons.filter(l => !isDone(l.id)).length : 0;
  return { m, n: m.n, open, reason, passed, testedOut, done, total, status, prev, prevLeft, minutesLeft };
}
export const canOpenLesson = (c, l) => moduleInfo(c, l.module || c.modules.find(m => m.lessons.some(x => x.id === l.id))).open;
// module tests: free modules for everyone; with a login, any module (taking a locked module's test = testing out)
export const canTakeTest = (c, m) => m.free || hasKey(c);
export const lastModule = c => c.modules[c.modules.length - 1];
export const projectOpen = c => hasKey(c) && (who() === 'admin' || override() || passedTest(c, lastModule(c).n));
export function examReqs(c) {
  const mods = c.modules.map(m => moduleInfo(c, m));
  const lessonsOk = mods.every(x => x.passed || x.done === x.total);
  const modsOk = mods.every(x => x.passed);
  const proj = store.d.projects[c.id]; const projOk = !!(proj && proj.best >= 70);
  return { lessonsOk, modsOk, projOk, proj, all: lessonsOk && modsOk && projOk };
}
export function lessonMinutesLeft(l) {
  const s = store.d.lessons[l.id]; if (s && s.status === 'done') return 0;
  const from = s ? (s.maxPart || 0) : 0; let t = 0; for (let i = from; i < 5; i++) t += partMinutes(l, i); return t;
}
// The one next thing to do in a course: a lesson (and which part), a module test, the project or the exam.
export function nextStep(c) {
  const all = lessonList(c);
  const started = all.find(l => { const s = store.d.lessons[l.id]; return s && s.status === 'started' && canOpenLesson(c, l); });
  if (started) { const s = store.d.lessons[started.id]; return { kind: 'lesson', lesson: started, part: s.part || 0, minutes: lessonMinutesLeft(started), continue: true }; }
  for (const m of c.modules) {
    const info = moduleInfo(c, m); if (info.passed) continue;
    if (!info.open) return { kind: 'locked', module: m, info };
    const l = m.lessons.map(x => all.find(y => y.id === x.id)).find(x => !isDone(x.id));
    if (l) return { kind: 'lesson', lesson: l, part: 0, minutes: lessonMinutesLeft(l) };
    return { kind: 'test', module: m, minutes: TEST_MINUTES };
  }
  if (!projectOpen(c)) return { kind: 'done' };
  if (!examReqs(c).projOk) return { kind: 'project', minutes: 180 };
  if (!(store.d.examAttempts[c.id] || []).some(a => a.passed)) return { kind: 'exam', minutes: c.exam.minutes };
  return { kind: 'finished' };
}
export function currentModule(c) { const s = nextStep(c); return s.lesson ? s.lesson.module : s.module || null; }
export function stepLabel(s) {
  if (!s) return '';
  if (s.kind === 'lesson') return `${s.continue ? 'Continue' : 'Start'}: ${lessonLabel(s.lesson)} · ${s.lesson.title}`;
  if (s.kind === 'test') return `Take the Module ${s.module.n} test`;
  if (s.kind === 'project') return 'Start the course project';
  if (s.kind === 'exam') return 'Take the final exam';
  if (s.kind === 'locked') return s.info.reason === 'login' ? `Module ${s.module.n} needs a student login` : `Pass the Module ${s.info.prev.n} test to open Module ${s.module.n}`;
  return 'Course complete';
}
export function stepHref(c, s) {
  if (s.kind === 'lesson') return `#/lesson/${c.id}/${s.lesson.id}/${PARTS[s.part || 0]}`;
  if (s.kind === 'test') return `#/mtest/${c.id}/${s.module.n}`;
  if (s.kind === 'project') return `#/project/${c.id}`;
  if (s.kind === 'exam') return `#/exam/${c.id}`;
  if (s.kind === 'locked') return s.info.reason === 'login' ? '#/login' : `#/mtest/${c.id}/${s.info.prev.n}`;
  return `#/certificates`;
}
export const fmtMin = m => m >= 90 ? `about ${Math.round(m / 60 * 2) / 2} hours` : m >= 60 ? 'about 1 hour' : `about ${m} min`;
