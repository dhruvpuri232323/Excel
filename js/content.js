// Loads course content. Free lessons are plain JSON; locked lessons and their files are
// AES-GCM encrypted and open only with the course key in a signed-in learner's session.
import { aesKey, decryptBytes } from './auth.js';
const SKEY = 'sheetah:session';
const keyCache = new Map(); const jsonCache = new Map(); const urlCache = new Map(); const packCache = new Map();
export const session = {
  s: null,
  load() { try { this.s = JSON.parse(localStorage.getItem(SKEY) || sessionStorage.getItem(SKEY) || 'null'); } catch (e) { this.s = null; } if (this.s && this.s.expires && this.s.expires < new Date().toISOString().slice(0, 10)) this.clear(); return this.s; },
  set(s, remember) { this.s = s; const j = JSON.stringify(s); if (remember) localStorage.setItem(SKEY, j); else sessionStorage.setItem(SKEY, j); },
  clear() { this.s = null; localStorage.removeItem(SKEY); sessionStorage.removeItem(SKEY); keyCache.clear(); packCache.clear(); jsonCache.clear(); },
  get signedIn() { return !!this.s; }, get isAdmin() { return !!(this.s && this.s.role === 'admin'); },
  has(course) { return !!(this.s && this.s.keys && this.s.keys[course]); },
  profileId() { return this.s ? this.s.id.slice(0, 24) : 'guest'; }
};
async function courseKey(course) { if (!session.has(course)) throw new Error('locked'); if (!keyCache.has(course)) keyCache.set(course, await aesKey(session.s.keys[course])); return keyCache.get(course); }
export async function getJSON(path) { if (jsonCache.has(path)) return jsonCache.get(path); const r = await fetch(path); if (!r.ok) throw new Error('Missing ' + path); const j = await r.json(); jsonCache.set(path, j); return j; }
// Locked files are bundled into one encrypted pack per lesson (plus a "course" pack for the exam and project).
// s/index.json says which pack holds which file; a pack is: 4-byte header length, JSON header {name: [offset, length]}, bytes.
async function openPack(course, pid, key) {
  const k = course + '/' + pid; if (!packCache.has(k)) packCache.set(k, (async () => { const r = await fetch(`content/${course}/s/${pid}.bin`); if (!r.ok) throw new Error('Missing locked file ' + pid); const u = new Uint8Array(await decryptBytes(key, await r.arrayBuffer())); const hl = new DataView(u.buffer, u.byteOffset, 4).getUint32(0); return { u, base: 4 + hl, head: JSON.parse(new TextDecoder().decode(u.subarray(4, 4 + hl))) }; })().catch(e => { packCache.delete(k); throw e; }));
  return packCache.get(k);
}
export async function getSecure(course, name, as = 'json') {
  const ck = course + '/' + name + ':' + as; if (jsonCache.has(ck)) return jsonCache.get(ck);
  const key = await courseKey(course); const idx = await getJSON(`content/${course}/s/index.json`); const pid = idx[name]; if (!pid) throw new Error('Missing locked file ' + name);
  const p = await openPack(course, pid, key); const e = p.head[name]; if (!e) throw new Error('Missing locked file ' + name);
  const bytes = p.u.slice(p.base + e[0], p.base + e[0] + e[1]); const out = as === 'json' ? JSON.parse(new TextDecoder().decode(bytes)) : bytes.buffer; jsonCache.set(ck, out); return out;
}
export const catalog = () => getJSON('content/catalog.json');
export const course = id => getJSON(`content/${id}/course.json`);
export function lessonMeta(c, id) { for (const m of c.modules) for (const l of m.lessons) if (l.id === id) return { ...l, module: m }; return null; }
export async function lesson(c, id) { const meta = lessonMeta(c, id); if (!meta) throw new Error('No such lesson'); if (meta.free) return getJSON(`content/${c.id}/l/${id}.json`); return getSecure(c.id, id); }
export const canOpen = (c, meta) => meta.free || session.has(c.id);
const MIME = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', png: 'image/png', pdf: 'application/pdf', svg: 'image/svg+xml' };
// returns a URL usable in <img src> or a download link
export async function assetURL(courseId, file) {
  if (file.startsWith('a/')) return `content/${courseId}/${file}`;
  const ck = courseId + '/' + file; if (urlCache.has(ck)) return urlCache.get(ck);
  const bytes = await getSecure(courseId, file.replace(/^s\//, ''), 'bytes'); const ext = file.split('.').pop();
  const url = URL.createObjectURL(new Blob([bytes], { type: MIME[ext] || 'application/octet-stream' })); urlCache.set(ck, url); return url;
}
