// FSRS-4.5 scheduler (open algorithm by Jarrett Ye et al.), written from the published formulas.
const W = [0.4072, 1.1829, 3.1262, 15.4722, 7.2102, 0.5316, 1.0651, 0.0234, 1.616, 0.1544, 1.0824, 1.9813, 0.0953, 0.2975, 2.2042, 0.2407, 2.9466, 0.5034, 0.6567];
const DECAY = -0.5, FACTOR = 19 / 81, RETAIN = 0.9, MAX_IVL = 3650;
export const AGAIN = 1, HARD = 2, GOOD = 3, EASY = 4;
const clampD = d => Math.min(10, Math.max(1, d));
const initS = g => Math.max(0.1, W[g - 1]);
const initD = g => clampD(W[4] - (g - 3) * W[5]);
export const retrievability = (t, s) => Math.pow(1 + FACTOR * t / s, DECAY);
const interval = s => Math.min(MAX_IVL, Math.max(1, Math.round(s / FACTOR * (Math.pow(RETAIN, 1 / DECAY) - 1))));
const nextD = (d, g) => clampD(W[7] * initD(4) + (1 - W[7]) * (d - W[6] * (g - 3)));
const recallS = (d, s, r, g) => s * (1 + Math.exp(W[8]) * (11 - d) * Math.pow(s, -W[9]) * (Math.exp((1 - r) * W[10]) - 1) * (g === HARD ? W[15] : 1) * (g === EASY ? W[16] : 1));
const forgetS = (d, s, r) => Math.min(s, W[11] * Math.pow(d, -W[12]) * (Math.pow(s + 1, W[13]) - 1) * Math.exp((1 - r) * W[14]));
export const dayNum = (date = new Date()) => Math.floor((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())) / 864e5);
export function dayStr(n) { const d = new Date(n * 864e5); return d.toISOString().slice(0, 10); }

// card: {s, d, reps, lapses, last (day number), due (day number)} ; returns new card state
export function review(card, grade, today = dayNum()) {
  const c = { ...card };
  if (!c.reps) { c.s = initS(grade); c.d = initD(grade); }
  else { const t = Math.max(0, today - (c.last ?? today)); const r = retrievability(t, c.s); c.d = nextD(c.d, grade); c.s = grade === AGAIN ? forgetS(card.d, card.s, r) : recallS(card.d, card.s, r, grade); if (grade === AGAIN) c.lapses = (c.lapses || 0) + 1; }
  c.reps = (c.reps || 0) + 1; c.last = today;
  c.due = grade === AGAIN ? today : today + interval(c.s);
  return c;
}
export function previewIntervals(card, today = dayNum()) { return [AGAIN, HARD, GOOD, EASY].map(g => { const n = review(card, g, today); return n.due - today; }); }
// chance of remembering in `days` from now
export function recallIn(card, days, today = dayNum()) { if (!card || !card.reps) return 0; return retrievability(today - card.last + days, card.s); }
