// Public site settings the teacher publishes inside data/students.json (field "site"):
// contact details, an announcement, "open everything", and the certificate register.
export const site = { data: {}, loaded: false };
const KEY = 'sheetah:site';
export async function loadSite() {
  try { const r = await fetch('data/students.json', { cache: 'no-store' }); if (!r.ok) throw 0; const j = await r.json(); site.data = j.site || {}; localStorage.setItem(KEY, JSON.stringify(site.data)); }
  catch (e) { try { site.data = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e2) { site.data = {}; } }
  site.loaded = true; return site.data;
}
export const teacher = () => site.data.teacher || {};
