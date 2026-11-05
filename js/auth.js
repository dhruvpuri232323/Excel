// Access without payments or servers.
// data/students.json lists one entry per learner. Each entry holds the course keys,
// encrypted with a key derived from that learner's own email + password (PBKDF2 + AES-GCM).
// Only someone who knows the password can unlock the paid lessons, even though the site is public.
const enc = new TextEncoder(), dec = new TextDecoder();
export const b64 = buf => btoa(String.fromCharCode(...new Uint8Array(buf)));
export const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const ITER = 250000;
export const normEmail = e => String(e || '').trim().toLowerCase();
export async function sha256hex(s) { const h = await crypto.subtle.digest('SHA-256', enc.encode(s)); return [...new Uint8Array(h)].map(b => b.toString(16).padStart(2, '0')).join(''); }
export const entryId = email => sha256hex('sheetah:' + normEmail(email));
async function passKey(email, password, salt, iter = ITER) {
  const base = await crypto.subtle.importKey('raw', enc.encode(normEmail(email) + '\u0000' + password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
export async function aesKey(rawB64) { return crypto.subtle.importKey('raw', unb64(rawB64), 'AES-GCM', false, ['encrypt', 'decrypt']); }
export async function encryptBytes(key, bytes) { const iv = crypto.getRandomValues(new Uint8Array(12)); const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes); const out = new Uint8Array(12 + ct.byteLength); out.set(iv); out.set(new Uint8Array(ct), 12); return out; }
export async function decryptBytes(key, buf) { const u = new Uint8Array(buf); return crypto.subtle.decrypt({ name: 'AES-GCM', iv: u.slice(0, 12) }, key, u.slice(12)); }
export async function decryptJSON(key, b64s) { return JSON.parse(dec.decode(await decryptBytes(key, unb64(b64s)))); }
export async function encryptJSON(key, obj) { return b64(await encryptBytes(key, enc.encode(JSON.stringify(obj)))); }

export async function loadRoster() {
  const r = await fetch('data/students.json', { cache: 'no-store' }); if (!r.ok) throw new Error('Could not load the student list');
  return r.json();
}
// returns the decrypted payload {email, name, role, keys:{c1:..}, adminKey?, expires?}
export async function login(email, password, roster) {
  roster = roster || await loadRoster(); const id = await entryId(email);
  const entry = roster.entries.find(e => e.id === id);
  if (!entry) throw new Error('That email or password is not right. Check with your teacher if you were given access.');
  try {
    const k = await passKey(email, password, unb64(entry.salt), entry.iter || ITER);
    const payload = JSON.parse(dec.decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(entry.iv) }, k, unb64(entry.data))));
    if (payload.expires && payload.expires < new Date().toISOString().slice(0, 10)) throw new Error('expired');
    if (entry.disabled) throw new Error('disabled');
    return { ...payload, id };
  } catch (e) {
    if (e.message === 'expired') throw new Error('Your access period has ended. Ask your teacher to renew it.');
    if (e.message === 'disabled') throw new Error('This account has been paused. Ask your teacher.');
    throw new Error('That email or password is not right. Check with your teacher if you were given access.');
  }
}
// Build an entry for a learner (used by the admin page)
export async function makeEntry({ email, password, name, role = 'student', keys, adminKey, courses, expires }) {
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12));
  const k = await passKey(email, password, salt);
  const payload = { email: normEmail(email), name, role, keys: Object.fromEntries(Object.entries(keys).filter(([c]) => !courses || courses.includes(c))), expires: expires || null };
  if (role === 'admin') payload.adminKey = adminKey;
  const data = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, k, enc.encode(JSON.stringify(payload)));
  return { id: await entryId(email), role, salt: b64(salt), iv: b64(iv), iter: ITER, data: b64(data) };
}
export function randomPassword() { const words = ['Tiger', 'Mango', 'Delta', 'River', 'Pixel', 'Comet', 'Lotus', 'Cedar', 'Nova', 'Orbit', 'Prism', 'Quartz']; const a = crypto.getRandomValues(new Uint32Array(3)); return words[a[0] % words.length] + '-' + (1000 + a[1] % 9000) + '-' + words[a[2] % words.length]; }
