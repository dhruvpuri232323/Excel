// Reading the VBA code out of an .xlsm: xl/vbaProject.bin is a Compound File (MS-CFB) whose VBA storage holds a
// compressed "dir" stream (MS-OVBA 2.3.4) listing the modules, and one stream per module whose source starts at
// the module's text offset, compressed the same way. We only read: nothing is ever run from the file itself.

const le16 = (b, o) => b[o] | (b[o + 1] << 8);
const le32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

// ---------------------------------------------------------------- MS-CFB: find a stream by its path
function cfb(bytes) {
  const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]; for (let i = 0; i < 8; i++) if (bytes[i] !== sig[i]) throw new Error('not a compound file');
  const ss = 1 << le16(bytes, 0x1E), mss = 1 << le16(bytes, 0x20); const nFat = le32(bytes, 0x2C), dir0 = le32(bytes, 0x30), cutoff = le32(bytes, 0x38), mini0 = le32(bytes, 0x3C), nMini = le32(bytes, 0x40);
  let difat0 = le32(bytes, 0x44); const nDifat = le32(bytes, 0x48);
  const sec = n => 512 + n * ss; // sector n starts after the 512-byte header (version 3; version 4 headers fill a whole 4096 sector)
  const SEC = n => ss === 4096 ? (n + 1) * 4096 : sec(n);
  const fatSecs = []; for (let i = 0; i < 109 && fatSecs.length < nFat; i++) { const s = le32(bytes, 0x4C + i * 4); if (s < 0xFFFFFFFA) fatSecs.push(s); }
  for (let k = 0; k < nDifat && difat0 < 0xFFFFFFFA; k++) { const at = SEC(difat0); for (let i = 0; i < ss / 4 - 1 && fatSecs.length < nFat; i++) { const s = le32(bytes, at + i * 4); if (s < 0xFFFFFFFA) fatSecs.push(s); } difat0 = le32(bytes, at + ss - 4); }
  const fat = []; for (const s of fatSecs) { const at = SEC(s); for (let i = 0; i < ss / 4; i++) fat.push(le32(bytes, at + i * 4)); }
  const chain = (start, max = 1e6) => { const out = []; let s = start; while (s < 0xFFFFFFFA && out.length < max) { out.push(s); s = fat[s]; } return out; };
  const readBig = (start, size) => { const out = new Uint8Array(size); let o = 0; for (const s of chain(start)) { const at = SEC(s); const n = Math.min(ss, size - o); if (n <= 0) break; out.set(bytes.subarray(at, at + n), o); o += n; } return out; };
  // the directory: a red-black tree of 128-byte entries per storage
  const dirBytes = readBig(dir0, chain(dir0).length * ss); const entries = [];
  for (let o = 0; o + 128 <= dirBytes.length; o += 128) { const nl = le16(dirBytes, o + 0x40); let name = ''; for (let i = 0; i < Math.max(0, nl - 2); i += 2) name += String.fromCharCode(le16(dirBytes, o + i));
    entries.push({ name, type: dirBytes[o + 0x42], left: le32(dirBytes, o + 0x44), right: le32(dirBytes, o + 0x48), child: le32(dirBytes, o + 0x4C), start: le32(dirBytes, o + 0x74), size: le32(dirBytes, o + 0x78) }); }
  const root = entries[0]; const miniStream = readBig(root.start, root.size);
  const miniFat = []; for (const s of chain(mini0, nMini || 1e6)) { const at = SEC(s); for (let i = 0; i < ss / 4; i++) miniFat.push(le32(bytes, at + i * 4)); }
  const readMini = (start, size) => { const out = new Uint8Array(size); let o = 0, s = start; while (s < 0xFFFFFFFA && o < size) { const n = Math.min(mss, size - o); out.set(miniStream.subarray(s * mss, s * mss + n), o); o += n; s = miniFat[s]; } return out; };
  const kids = e => { const out = []; const walk = i => { if (i >= entries.length || i === 0xFFFFFFFF) return; const x = entries[i]; walk(x.left); out.push(x); walk(x.right); }; walk(e.child); return out; };
  const find = path => { let e = root; for (const part of path) { e = kids(e).find(x => x.name.toLowerCase() === part.toLowerCase()); if (!e) return null; } return e; };
  return { read(path) { const e = find(path); if (!e || e.type !== 2) return null; return e.size < cutoff ? readMini(e.start, e.size) : readBig(e.start, e.size); }, list(path) { const e = find(path); return e ? kids(e).map(x => x.name) : []; } };
}

// ---------------------------------------------------------------- MS-OVBA 2.4.1: decompress a CompressedContainer
export function decompress(b, at = 0) {
  if (b[at] !== 1) throw new Error('not a compressed VBA container'); const out = []; let p = at + 1;
  while (p < b.length - 1) { const hdr = le16(b, p); const size = (hdr & 0x0FFF) + 3, compressed = (hdr & 0x8000) !== 0; const end = Math.min(p + size, b.length); p += 2; const start = out.length;
    if (!compressed) { for (let i = 0; i < 4096 && p < end; i++) out.push(b[p++]); continue; }
    while (p < end) { const flags = b[p++]; for (let bit = 0; bit < 8 && p < end; bit++) {
      if (!(flags & (1 << bit))) { out.push(b[p++]); continue; }
      const tok = le16(b, p); p += 2; const done = out.length - start; let bits = 4; while ((1 << bits) < done) bits++; bits = Math.max(bits, 4);
      const lenMask = 0xFFFF >> bits; const len = (tok & lenMask) + 3, off = (tok >> (16 - bits)) + 1; const from = out.length - off; for (let i = 0; i < len; i++) out.push(out[from + i]); } }
    p = end; }
  return Uint8Array.from(out); }

const cp1252 = bytes => { try { return new TextDecoder('windows-1252').decode(bytes); } catch (e) { return String.fromCharCode(...bytes); } };

// ---------------------------------------------------------------- the dir stream: module names, stream names, text offsets and types
function modulesOf(dir) { const mods = []; let cur = null; let p = 0;
  while (p + 6 <= dir.length) { const id = le16(dir, p), size = le32(dir, p + 2); p += 6;
    if (id === 0x0009) { p += 6; continue; } // PROJECTVERSION: its size field lies (4), the record is 6 bytes long
    const data = dir.subarray(p, p + size); p += size;
    if (id === 0x0019) { cur = { name: cp1252(data), stream: null, offset: 0, kind: 'std' }; mods.push(cur); }
    else if (cur && id === 0x001A) cur.stream = cp1252(data);
    else if (cur && id === 0x0031) cur.offset = le32(data, 0);
    else if (cur && id === 0x0022) cur.kind = 'class'; // document modules (sheets, ThisWorkbook) and classes; told apart by VB_Base below
    else if (id === 0x0010) break; } // end of the dir stream
  return mods; }

// readVbaProject(bytes of xl/vbaProject.bin) → [{ name, kind: 'std' | 'class' | 'document' | 'form', code }]
export function readVbaProject(bytes) { const f = cfb(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)); const dirRaw = f.read(['VBA', 'dir']); if (!dirRaw) throw new Error('no VBA project in this file');
  const out = [];
  for (const m of modulesOf(decompress(dirRaw))) { const s = f.read(['VBA', m.stream || m.name]); if (!s) continue; let text = ''; try { text = cp1252(decompress(s, m.offset)); } catch (e) { continue; }
    const attrs = /^Attribute VB_Base = "0\{00020820-|^Attribute VB_Base = "0\{00020819-/m.test(text) ? 'document' : /^Attribute VB_Base = "0\{[0-9A-F-]+\}\{[0-9A-F-]+\}"/m.test(text) ? 'form' : m.kind;
    const code = text.replace(/\r\n?/g, '\n').split('\n').filter(l => !/^Attribute VB_/i.test(l)).join('\n').replace(/^\n+/, '');
    out.push({ name: m.name, kind: attrs, code }); }
  return out; }
