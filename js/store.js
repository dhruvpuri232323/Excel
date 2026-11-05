// Local progress store. Each person on a laptop has their own profile, kept in the browser.
import { dayNum, dayStr } from './srs.js';
const KEY = 'sheetah:v1:';
const blank = () => ({
  createdAt: new Date().toISOString(), name: '', lessons: {}, cards: {}, errors: [], activity: {}, streak: { freezesUsed: [] },
  plan: { days: [1, 3, 5], time: '20:00', target: '', minutesGoal: 30, missed: {} }, checklist: {}, journal: {}, weekly: {}, tests: [], projects: {}, certs: [],
  notes: {}, bookmarks: [], milestones: [], retention: {}, examAttempts: {},
  settings: { theme: 'auto', tone: 'warm', narration: false, reduceMotion: false, os: navigator.platform && /Mac/i.test(navigator.platform) ? 'mac' : 'win', remind: true, newPerDay: 15, textSize: 'normal' }
});
export const today = () => dayStr(dayNum());
export const store = {
  pid: 'guest', d: blank(), listeners: new Set(),
  load(pid) { this.pid = pid; try { const raw = localStorage.getItem(KEY + pid); this.d = Object.assign(blank(), raw ? JSON.parse(raw) : {}); this.d.settings = Object.assign(blank().settings, this.d.settings); this.d.plan = Object.assign(blank().plan, this.d.plan); } catch (e) { this.d = blank(); } this.emit(); },
  save() { try { localStorage.setItem(KEY + this.pid, JSON.stringify(this.d)); } catch (e) { console.warn('Could not save progress', e); } this.emit(); },
  on(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }, emit() { for (const f of this.listeners) f(this.d); },
  lesson(id) { return this.d.lessons[id] || (this.d.lessons[id] = { status: 'new', part: 0, tasks: {}, minutes: 0 }); },
  act(field, n = 1) { const t = today(); const a = this.d.activity[t] || (this.d.activity[t] = { minutes: 0, lessons: 0, reviews: 0, tasks: 0 }); a[field] = (a[field] || 0) + n; },
  // a study day = any finished lesson, review or 10+ minutes of study
  studied(date) { const a = this.d.activity[date]; return !!a && (a.lessons > 0 || a.reviews >= 5 || a.minutes >= 10) || this.d.streak.freezesUsed.includes(date); },
  streak() { let n = 0, day = dayNum(); if (!this.studied(dayStr(day))) day--; while (this.studied(dayStr(day))) { n++; day--; } return n; },
  bestStreak() { const days = Object.keys(this.d.activity).filter(d => this.studied(d)).concat(this.d.streak.freezesUsed).map(d => Math.floor(Date.parse(d) / 864e5)).sort((a, b) => a - b); let best = 0, run = 0, prev = null; for (const d of [...new Set(days)]) { run = prev !== null && d === prev + 1 ? run + 1 : 1; best = Math.max(best, run); prev = d; } return best; },
  freezesLeft() { const m = today().slice(0, 7); return 2 - this.d.streak.freezesUsed.filter(d => d.startsWith(m)).length; },
  useFreeze() { const y = dayStr(dayNum() - 1); if (this.freezesLeft() <= 0 || this.studied(y)) return false; this.d.streak.freezesUsed.push(y); this.save(); return true; },
  logError(e) { const id = 'e' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); const d0 = dayNum(); this.d.errors.unshift({ id, at: new Date().toISOString(), due: [d0 + 1, d0 + 3, d0 + 7], step: 0, resolved: false, cause: e.cause || '', ...e }); this.d.errors = this.d.errors.slice(0, 500); return id; },
  exportJSON() { return JSON.stringify({ app: 'sheetah', version: 1, profile: this.pid, exportedAt: new Date().toISOString(), data: this.d }, null, 1); },
  importJSON(text) { const j = JSON.parse(text); if (j.app !== 'sheetah' || !j.data) throw new Error('This is not a Sheetah backup file.'); this.d = Object.assign(blank(), j.data); this.save(); },
};
// Count active study minutes: while the page is visible and the learner did something in the last 2 minutes
let lastInput = Date.now();
['keydown', 'mousedown', 'wheel', 'touchstart'].forEach(ev => addEventListener(ev, () => { lastInput = Date.now(); }, { passive: true }));
setInterval(() => { if (document.visibilityState === 'visible' && Date.now() - lastInput < 120000 && store.studying) { store.act('minutes', 0.5); if (store.currentLesson) { const l = store.lesson(store.currentLesson); l.minutes = (l.minutes || 0) + 0.5; } store.save(); } }, 30000);
