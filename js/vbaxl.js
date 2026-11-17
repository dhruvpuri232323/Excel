// The Excel object model for the VBA interpreter, on Sheetah's own Workbook engine: Application, Workbooks, Worksheets,
// Range (Value, Formula, Offset, Resize, End, CurrentRegion, Font, Interior, NumberFormat, Copy, Delete, Sort,
// AutoFilter, SpecialCells, Find), ListObjects, WorksheetFunction, and a pretend folder of files for Workbooks.Open,
// SaveAs, Dir, Kill and PDFs. Messages are Excel's. Coordinates inside are 0-based; VBA sees 1-based.
import { VObject, VArray, VDate, vbErr, bind, toNum, toStr, toBool, deref, isObj, EMPTY, NOTHING, MISSING, NULLV, CONSTS } from './vba.js';
import { Workbook, parseRange, parseAddr, colName, colIndex, addr, isDateFmt, isErr, shiftFormula, dateToSerial, formatValue, FN } from './engine.js';

export const MAXR = 1048576, MAXC = 16384;
export const XL = { xlup: -4162, xldown: -4121, xltoleft: -4159, xltoright: -4161, xlcalculationmanual: -4135, xlcalculationautomatic: -4105, xlcalculationsemiautomatic: 2, xlcenter: -4108, xlleft: -4131, xlright: -4152, xlgeneral: 1, xltop: -4160, xlbottom: -4107,
  xlnone: -4142, xlcontinuous: 1, xldash: -4115, xldot: -4118, xldouble: -4119, xlthin: 2, xlmedium: -4138, xlthick: 4, xlhairline: 1, xledgeleft: 7, xledgetop: 8, xledgebottom: 9, xledgeright: 10, xlinsidevertical: 11, xlinsidehorizontal: 12, xldiagonaldown: 5, xldiagonalup: 6,
  xltypepdf: 0, xltypexps: 1, xlqualitystandard: 0, xlopenxmlworkbook: 51, xlopenxmlworkbookmacroenabled: 52, xlworkbookdefault: 51, xlcsv: 6, xlcsvutf8: 62, xlexcel8: 56, xlworkbooknormal: -4143,
  xlyes: 1, xlno: 2, xlguess: 0, xlascending: 1, xldescending: 2, xlsortonvalues: 0, xltoptobottom: 1, xlpinyin: 1, xlsortnormal: 0,
  xlcelltypevisible: 12, xlcelltypeconstants: 2, xlcelltypeformulas: -4123, xlcelltypeblanks: 4, xlcelltypelastcell: 11, xlwhole: 1, xlpart: 2, xlvalues: -4163, xlformulas: -4123, xlbyrows: 1, xlbycolumns: 2, xlnext: 1, xlprevious: 2,
  xlpastevalues: -4163, xlpasteformats: -4122, xlpasteall: -4104, xlpasteformulas: -4123, xlpastevaluesandnumberformats: 12, xlpastecolumnwidths: 8, xlpastespecialoperationnone: -4142,
  xland: 1, xlor: 2, xlfiltervalues: 7, xltop10items: 3, xlshiftup: -4162, xlshiftdown: -4121, xlshifttoleft: -4159, xlshifttoright: -4161, xlformatfromleftorabove: 0,
  xlsheetvisible: -1, xlsheethidden: 0, xlsheetveryhidden: 2, xlworksheet: -4167, xlsrcrange: 1, xlsrcexternal: 0, xllistobjecthasheaders: 1, xldatabase: 1, xla1: 1, xlr1c1: -4150,
  xlcolorindexnone: -4142, xlcolorindexautomatic: -4105, xlautomatic: -4105, xlthemecolordark1: 1, xlthemecolorlight1: 2, xlsolid: 1, xlpatternnone: -4142,
  xlmaximized: -4137, xlnormal: -4143, xlwait: 2, xldefault: -4143, xlnochange: 1, olmailitem: 0, olformathtml: 2, olformatplain: 1, xlportrait: 1, xllandscape: 2, xlpapera4: 9, xlpaperletter: 1, olimportancehigh: 2, xlerrdiv0: 2007, xlerrna: 2042, xlerrname: 2029, xlerrnull: 2000, xlerrnum: 2036, xlerrref: 2023, xlerrvalue: 2015, xlopenxmladdin: 55, xladdin: 18, msoFileDialogFilePicker: 3, msofiledialogfolderpicker: 4 };
const COLORS56 = [0, 0, 16777215, 255, 65280, 16711680, 65535, 16711935, 16776960, 128, 32768, 8388608, 32896, 8388736, 8421376, 12632256, 8421504, 16751001, 6697881, 13434879, 16777164, 6684774, 8421631, 13395456, 16764108, 8388608, 16711935, 65535, 16776960, 8388736, 128, 8421376, 16711680, 16763904, 16777164, 13434828, 10092543, 16764057, 13408767, 16751052, 10079487, 16737843, 13421619, 52479, 39423, 26367, 10053222, 9868950, 6697728, 6723891, 13056, 13107, 13209, 6697881, 10040115, 3355443];
export const bgrToHex = v => { const n = Math.round(toNum(v)); const r = n & 255, g = (n >> 8) & 255, b = (n >> 16) & 255; return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase(); };
export const hexToBgr = h => { if (!h) return 16777215; const x = h.replace('#', ''); const r = parseInt(x.slice(0, 2), 16), g = parseInt(x.slice(2, 4), 16), b = parseInt(x.slice(4, 6), 16); return r + g * 256 + b * 65536; };
const ALIGN = { [-4108]: 'center', [-4131]: 'left', [-4152]: 'right' }, ALIGN_BACK = { center: -4108, left: -4131, right: -4152 };
const lc = s => String(s).toLowerCase();
const has = a => a !== undefined && a !== MISSING;
const tail = p => String(p).split(/[\\/]/).pop();
const dirOf = p => { const s = String(p).replace(/\//g, '\\'); const i = s.lastIndexOf('\\'); return i < 0 ? '' : s.slice(0, i); };
const wild = (pat) => new RegExp('^' + pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');

// ---------------------------------------------------------------- the host: what the interpreter calls for anything Excel
export class ExcelHost {
  // opts: { wb: Workbook (ThisWorkbook), name, path, files: { 'C:\\…\\x.xlsx': {kind, spec|text} }, now: serial, screen: 'on' }
  constructor(opts = {}) {
    this.o = opts; this.consts = XL; this.clip = null; this.cost = { reads: 0, writes: 0 };
    this.fs = new Map(); for (const [p, f] of Object.entries(opts.files || {})) this.fs.set(lc(p), { path: p, ...f });
    for (const d of opts.folders || []) this.fs.set(lc(d), { path: d, kind: 'folder' });
    this.path = opts.path || 'C:\\Users\\you\\Documents'; this.app = new App(this);
    this.this = new Book(this, opts.wb || new Workbook(), opts.name || 'Book1.xlsm', this.path); this.books = [this.this]; this.active = this.this;
    this.sel = { book: this.this, si: 0, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 };
    this.screen = true; this.alerts = true; this.events = true; this.pending = []; this.calc = -4105; this.status = false; this.log = []; this.outbox = []; this.dirState = null;
  }
  now() { return this.o.now !== undefined ? this.o.now : dateToSerial(2026, 11, 16) + 10 / 24; }
  globalNames() { return ['range', 'cells', 'rows', 'columns', 'activesheet', 'activecell', 'activeworkbook', 'thisworkbook', 'workbooks', 'worksheets', 'sheets', 'selection', 'application', 'worksheetfunction', 'intersect', 'union', 'calculate', 'names', 'activewindow', 'evaluate']; }
  fnNames() { return ['dir', 'kill', 'filecopy', 'mkdir', 'rmdir', 'filelen', 'filedatetime', 'curdir', 'chdir', 'getattr']; }
  global(n, args, rt) { if (n === 'application') return this.app; if (n === 'thisworkbook') return this.this; if (this.app.globalOk(n)) return this.app.get(n, args, rt); return undefined; }
  globalPut(n, args, v, rt) { if (n === 'activecell' || n === 'selection') { this.app.get(n, [], rt).putDflt([], v, rt); return true; } return false; }
  me(mod) { if (mod.kind === 'sheet') { const si = this.this.wb.sheets.findIndex(s => lc(s.name) === lc(mod.sheet || mod.name)); return si >= 0 ? new Sheet(this, this.this, si) : null; } if (mod.kind === 'workbook') return this.this; if (mod.kind === 'form' && this.forms) return this.forms(mod.name); return null; }
  fn(n) { return FS_FNS[n] ? (args, rt) => FS_FNS[n](this, args.map(a => deref(a.v, rt)), rt) : null; }
  newObj(cls, rt, byProgId) { if (this.o.newObj) { const o = this.o.newObj(cls, this, rt); if (o) return o; } if (cls === 'outlook.application') return this.outlook || (this.outlook = new OlApp(this)); return null; }
  // costs, in seconds, for Timer: writing a cell with the screen updating is slow, reading a whole range at once is fast
  // a macro changed cells of this workbook: Worksheet_Change runs after the statement (vba.js flushEvents), unless EnableEvents is False
  changed(r) { if (this.events !== false && r.book === this.this) this.pending.push({ sheet: r.sheet.name, target: r }); }
  tick(kind, n = 1, rt) { if (!rt) return; const auto = this.calc === -4105 && kind === 'write' ? 0.0004 : 0; rt.clock += n * (kind === 'write' ? (this.screen ? 0.0012 : 0.0002) + auto : kind === 'read' ? 0.00015 : kind === 'bulk' ? 0.002 : kind === 'select' ? (this.screen ? 0.004 : 0.0004) : 0.0001); } // Select and Activate redraw the screen: slow unless ScreenUpdating is off
  bookByName(n) { const k = lc(n); return this.books.find(b => lc(b.name) === k || lc(b.name.replace(/\.[^.]+$/, '')) === k); }
  file(path) { return this.fs.get(lc(path)) || null; }
  writeFile(path, f) { const d = dirOf(path); if (d && !this.folderExists(d)) throw vbErr(1004, `Microsoft Excel cannot access the file '${path}'. The file name or path does not exist.`); this.fs.set(lc(path), { path, ...f, made: true }); this.log.push({ kind: 'file', path, type: f.kind }); }
  folderExists(d) { const k = lc(d).replace(/\\$/, ''); if (k === lc(this.path) || /^[a-z]:$/.test(k)) return true; for (const [p, f] of this.fs) { if (p === k && f.kind === 'folder') return true; if (p.startsWith(k + '\\')) return true; } return false; }
  sheetOf(book, si) { return book.wb.sheets[si]; }
}

// ---------------------------------------------------------------- Application
const APP_GLOBAL = new Set(['range', 'cells', 'rows', 'columns', 'activesheet', 'activecell', 'activeworkbook', 'workbooks', 'worksheets', 'sheets', 'selection', 'worksheetfunction', 'intersect', 'union', 'calculate', 'names', 'activewindow', 'evaluate']);
class App extends VObject {
  constructor(h) { super(); this.h = h; } get tn() { return 'Application'; }
  globalOk(n) { return APP_GLOBAL.has(n); }
  sheetObj(rt) { const b = this.h.active; return new Sheet(this.h, b, b.activeSi); }
  get(n, a, rt) { const h = this.h;
    switch (n) {
      case 'activeworkbook': return h.active; case 'thisworkbook': return h.this; case 'workbooks': return new Books(h);
      case 'activesheet': return this.sheetObj(rt); case 'worksheets': case 'sheets': { const c = new SheetsColl(h, h.active); return a.length ? c.dflt(a, rt) : c; }
      case 'range': case 'cells': case 'rows': case 'columns': return this.sheetObj(rt).get(n, a, rt);
      case 'activecell': { const s = h.sel; return new Rng(h, s.book, s.si, { r1: s.ar, c1: s.ac, r2: s.ar, c2: s.ac }); }
      case 'selection': { const s = h.sel; return new Rng(h, s.book, s.si, { ...s.g }); }
      case 'screenupdating': return h.screen; case 'displayalerts': return h.alerts; case 'enableevents': return h.events; case 'calculation': return h.calc; case 'statusbar': return h.status; case 'cutcopymode': return h.clip ? 1 : false;
      case 'worksheetfunction': return new WsFn(h, true); case 'username': return 'you'; case 'userlibrarypath': return 'C:\\Users\\you\\AppData\\Roaming\\Microsoft\\AddIns\\'; case 'librarypath': return 'C:\\Program Files\\Microsoft Office\\root\\Office16\\Library'; case 'name': return 'Microsoft Excel'; case 'version': return '16.0'; case 'pathseparator': return '\\'; case 'defaultfilepath': return h.path; case 'operatingsystem': return 'Windows (64-bit) NT 10.00';
      case 'macrooptions': { const [mac, desc, , , , , cat] = bind(a, ['macro', 'description', 'hasmenu', 'menutext', 'hasshortcutkey', 'shortcutkey', 'category', 'statusbar', 'helpcontextid', 'helpfile', 'argumentdescriptions']); const k = toStr(mac).toLowerCase(); h.macroOptions = h.macroOptions || {}; h.macroOptions[k] = { description: has(desc) ? toStr(desc) : '', category: has(cat) ? deref(cat) : '' }; h.log.push({ kind: 'macrooptions', macro: k }); return EMPTY; }
      case 'wait': case 'ontime': case 'volatile': case 'calculatefull': case 'goto': case 'quit': return EMPTY; case 'calculate': h.this.wb.invalidate(); return EMPTY;
      case 'intersect': { const rs = a.map(x => x.v).filter(x => x !== MISSING); let g = null; for (const r of rs) { if (!(r instanceof Rng)) throw vbErr(13); if (!g) g = { ...r.g, si: r.si, book: r.book }; else { if (r.si !== g.si) throw vbErr(1004); g = { si: g.si, book: g.book, r1: Math.max(g.r1, r.g.r1), c1: Math.max(g.c1, r.g.c1), r2: Math.min(g.r2, r.g.r2), c2: Math.min(g.c2, r.g.c2) }; } if (g.r1 > g.r2 || g.c1 > g.c2) return NOTHING; } return new Rng(h, g.book, g.si, { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }); }
      case 'union': { const rs = a.map(x => x.v).filter(x => x instanceof Rng); const rows = []; let g = null; for (const r of rs) { for (const rr of r.rowList()) if (!rows.includes(rr)) rows.push(rr); g = g ? { r1: Math.min(g.r1, r.g.r1), c1: Math.min(g.c1, r.g.c1), r2: Math.max(g.r2, r.g.r2), c2: Math.max(g.c2, r.g.c2) } : { ...r.g }; } return new Rng(h, rs[0].book, rs[0].si, g, rows.sort((x, y) => x - y)); }
      case 'run': { const [name, ...rest] = a.map(x => x.v); return { ui: 'run', name: toStr(name), args: rest, vbaRun: true }; }
      case 'inputbox': { const [p, t, d, , , , , ty] = bind(a, ['prompt', 'title', 'default', 'left', 'top', 'helpfile', 'helpcontextid', 'type']); return { ui: 'inputbox', prompt: toStr(p), title: has(t) ? toStr(t) : 'Input', def: has(d) ? toStr(d) : '', type: has(ty) ? toNum(ty) : 2, app: true, after: r => r === null || r === undefined || r === false ? false : (has(ty) && (toNum(ty) & 1) ? toNum(r) : String(r)) }; }
      case 'getopenfilename': { return { ui: 'openfile', after: r => r || false }; }
      case 'evaluate': return this.h.this.evaluate(toStr(a[0].v), this.h.active.activeSi);
      case 'activewindow': return new Win(h); case 'names': return new NamesColl(h, h.active);
    }
    if (WsFn.has(n)) return new WsFn(h, false).get(n, a, rt);
    throw vbErr(438); }
  put(n, a, v) { const h = this.h; switch (n) { case 'screenupdating': h.screen = toBool(v); return; case 'displayalerts': h.alerts = toBool(v); return; case 'enableevents': h.events = toBool(v); return; case 'calculation': h.calc = toNum(v); return; case 'statusbar': h.status = v === false ? false : toStr(v); h.log.push({ kind: 'status', text: h.status }); return; case 'cutcopymode': if (!toBool(v)) h.clip = null; return; case 'windowstate': case 'cursor': case 'askToUpdateLinks': case 'asktoupdatelinks': return; } throw vbErr(438); }
}
class Win extends VObject { constructor(h) { super(); this.h = h; } get tn() { return 'Window'; } get(n) { if (['freezepanes', 'splitrow', 'splitcolumn', 'zoom', 'displaygridlines'].includes(n)) return false; throw vbErr(438); } put(n, a, v) { if (n === 'freezepanes') { const s = this.h.sel; const sh = s.book.wb.sheets[s.si]; sh.freeze = toBool(v) ? addr(s.ar, s.ac) : null; return; } if (['splitrow', 'splitcolumn', 'zoom', 'displaygridlines'].includes(n)) return; throw vbErr(438); } }

// ---------------------------------------------------------------- Workbooks and Workbook
class Books extends VObject {
  constructor(h) { super(); this.h = h; } get tn() { return 'Workbooks'; }
  get(n, a, rt) { const h = this.h;
    if (n === 'count') return h.books.length;
    if (n === 'item') return this.dflt(a, rt);
    if (n === 'add') { let k = 1; while (h.bookByName('Book' + k)) k++; const b = new Book(h, new Workbook({ sheets: [{ name: 'Sheet1' }] }), 'Book' + k, ''); b.unsaved = true; h.books.push(b); h.active = b; h.sel = { book: b, si: 0, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; return b; }
    if (n === 'open') { const [fn] = bind(a, ['filename', 'updatelinks', 'readonly']); const p = toStr(fn); const already = h.books.find(b => lc(b.fullName) === lc(p)); if (already) { h.active = already; return already; }
      const f = h.file(p); if (!f || f.kind === 'folder') throw vbErr(1004, `Sorry, we couldn't find ${p}. Is it possible it was moved, renamed or deleted?`);
      let wb; if (f.kind === 'csv') { const rows = f.text.replace(/\r/g, '').split('\n').filter((l, i, arr) => l || i < arr.length - 1).map(l => l.split(',')); wb = new Workbook({ sheets: [{ name: tail(p).replace(/\.[^.]+$/, ''), cells: {} }] }); rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(0, r, c, v); })); }
      else wb = new Workbook(JSON.parse(JSON.stringify(f.spec))); const b = new Book(this.h, wb, tail(p), dirOf(p)); h.books.push(b); h.active = b; h.sel = { book: b, si: 0, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; h.log.push({ kind: 'open', path: p }); return b; }
    if (n === 'close') { h.books = [h.this]; h.active = h.this; return EMPTY; }
    throw vbErr(438); }
  dflt(a) { const k = deref(a[0].v); const h = this.h; if (typeof k === 'number') { const b = h.books[k - 1]; if (!b) throw vbErr(9); return b; } const b = h.bookByName(toStr(k)); if (!b) throw vbErr(9); return b; }
  each() { return this.h.books.slice(); }
}
class Book extends VObject {
  constructor(h, wb, name, path) { super(); this.h = h; this.wb = wb; this.name = name; this.path = path; this.activeSi = 0; this.saved = true; }
  get tn() { return 'Workbook'; } get fullName() { return this.path ? this.path + '\\' + this.name : this.name; }
  same(o) { return o === this; }
  evaluate(f, si) { const v = this.wb.evaluate(f.startsWith('=') ? f : '=' + f, si); return fromCell(v, null); }
  get(n, a, rt) { const h = this.h;
    switch (n) {
      case 'name': return this.name; case 'path': return this.path; case 'fullname': return this.fullName; case 'saved': return this.saved; case 'isaddin': return !!this.isAddin; case 'readonly': return false;
      case 'worksheets': case 'sheets': { const c = new SheetsColl(h, this); return a.length ? c.dflt(a, rt) : c; }
      case 'activesheet': return new Sheet(h, this, this.activeSi);
      case 'activate': h.active = this; return EMPTY;
      case 'save': this.saved = true; if (this.path) h.writeFile(this.fullName, { kind: /\.xlsm$/i.test(this.name) ? 'xlsm' : 'xlsx', spec: this.wb.toSpec() }); return EMPTY;
      case 'saveas': case 'savecopyas': { const [fn, ff] = bind(a, n === 'saveas' ? ['filename', 'fileformat'] : ['filename']); let p = toStr(fn); if (!/[\\/]/.test(p)) p = h.path + '\\' + p; const fmt = has(ff) ? toNum(ff) : null;
        if (fmt === 6 || fmt === 62 || /\.csv$/i.test(p)) { if (!/\.csv$/i.test(p)) p += '.csv'; const s = this.wb.sheets[this.activeSi]; const u = this.wb.usedRange(s); const lines = []; if (u) for (let r = 0; r <= u.r2; r++) { const row = []; for (let c = 0; c <= u.c2; c++) { const v = this.wb.value(this.activeSi, r, c); row.push(v === null || v === undefined ? '' : this.wb.display(this.activeSi, r, c)); } lines.push(row.join(',')); } h.writeFile(p, { kind: 'csv', text: lines.join('\r\n') + '\r\n' }); }
        else { if (!/\.(xlsx|xlsm|xls|xlsb|xlam)$/i.test(p)) p += fmt === 52 ? '.xlsm' : fmt === 55 ? '.xlam' : '.xlsx'; if ((/\.xlam$/i.test(p) && fmt !== 55 && fmt !== 18) || (/\.xlsm$/i.test(p) && fmt === 51) || (/\.xlsx$/i.test(p) && (fmt === 52 || fmt === 55))) throw vbErr(1004, 'This extension can not be used with the selected file type. Change the file extension in the File name text box or select a different file type by changing the Save as type.'); if (/\.xlsx$/i.test(p) && h.this === this && !/\.xlsx$/i.test(this.name)) { if (h.alerts) return { ui: 'msgbox', prompt: 'The following features cannot be saved in macro-free workbooks:\n\n• VB project\n\nTo save a file with these features, click No, and then choose a macro-enabled file type in the File Type list.\nTo continue saving as a macro-free workbook, click Yes.', buttons: 4 + 48, title: 'Microsoft Excel', after: r => { if (r === 6) this.saveTo(p, 'xlsx'); return EMPTY; } }; }
          const ex = h.file(p); if (ex && h.alerts && n === 'saveas' && lc(p) !== lc(this.fullName)) return { ui: 'msgbox', prompt: `A file named '${p}' already exists in this location. Do you want to replace it?`, buttons: 3 + 48, title: 'Microsoft Excel', after: r => { if (r === 6 || r === undefined) { this.saveTo(p, /\.xlsm$/i.test(p) ? 'xlsm' : /\.xlam$/i.test(p) ? 'xlam' : 'xlsx'); return EMPTY; } throw vbErr(1004, "Method 'SaveAs' of object '_Workbook' failed"); } };
          this.saveTo(p, /\.xlsm$/i.test(p) ? 'xlsm' : /\.xlam$/i.test(p) ? 'xlam' : 'xlsx'); }
        if (n === 'saveas') { this.name = tail(p); this.path = dirOf(p); } return EMPTY; }
      case 'close': { const [sc] = bind(a, ['savechanges', 'filename']); if (has(sc) && toBool(sc) && this.path) this.saveTo(this.fullName, /\.xlsm$/i.test(this.name) ? 'xlsm' : 'xlsx'); if (this === h.this) throw vbErr(1004, 'Sheetah keeps this workbook open: close the other workbooks instead.'); h.books = h.books.filter(b => b !== this); if (h.active === this) h.active = h.this; if (h.sel.book === this) h.sel = { book: h.this, si: h.this.activeSi, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; h.log.push({ kind: 'close', name: this.name }); return EMPTY; }
      case 'exportasfixedformat': { const [ty, fn] = bind(a, ['type', 'filename', 'quality', 'includedocproperties', 'ignoreprintareas', 'from', 'to', 'openafterpublish']); return exportPdf(h, this, null, fn); }
      case 'names': return new NamesColl(h, this); case 'parent': return h.app; case 'protect': case 'unprotect': case 'refreshall': return EMPTY;
      case 'windows': return EMPTY;
    } throw vbErr(438); }
  saveTo(p, kind) { this.h.writeFile(p, { kind, spec: this.wb.toSpec() }); this.saved = true; }
  put(n, a, v) { if (n === 'saved') { this.saved = toBool(v); return; } if (n === 'isaddin') { this.isAddin = toBool(v); return; } throw vbErr(438); }
}
function exportPdf(h, book, si, fn, g) { let p = toStr(fn); if (!p) throw vbErr(1004, 'Document not saved.'); if (!/[\\/]/.test(p)) p = h.path + '\\' + p; if (!/\.pdf$/i.test(p)) p += '.pdf';
  const sheets = si === null ? book.wb.sheets.map((s, i) => i) : [si]; const pages = sheets.map(i => { const s = book.wb.sheets[i]; const pa = !g && s.page && s.page.printarea ? (() => { try { return parseRange(String(s.page.printarea).replace(/\$/g, '')); } catch (e) { return null; } })() : null; const u = g || pa || book.wb.usedRange(s); const rows = []; if (u) for (let r = u.r1; r <= u.r2; r++) { if (s.hidden && s.hidden.has(r)) continue; const row = []; for (let c = u.c1; c <= u.c2; c++) row.push(book.wb.display(i, r, c)); rows.push(row); } return { sheet: s.name, rows, ...(s.page ? { setup: { ...s.page } } : {}) }; });
  h.writeFile(p, { kind: 'pdf', pages }); return EMPTY; }
class NamesColl extends VObject { constructor(h, b) { super(); this.h = h; this.b = b; } get tn() { return 'Names'; } get(n, a) { if (n === 'count') return Object.keys(this.b.wb.names).length; if (n === 'add') { const [nm, ref] = bind(a, ['name', 'referstorange', 'refersto']); const r = a.find(x => x.name === 'referstorange') ? a.find(x => x.name === 'referstorange').v : ref; if (r instanceof Rng) this.b.wb.defineName(toStr(nm), `'${r.sheet.name}'!${r.address(true, true)}`); else this.b.wb.defineName(toStr(nm), toStr(r).replace(/^=/, '')); return EMPTY; } throw vbErr(438); } }

// ---------------------------------------------------------------- Worksheets and Worksheet
class SheetsColl extends VObject {
  constructor(h, b) { super(); this.h = h; this.b = b; } get tn() { return 'Sheets'; }
  get(n, a, rt) { const wb = this.b.wb;
    if (n === 'count') return wb.sheets.length; if (n === 'item') return this.dflt(a, rt);
    if (n === 'add') { const [before, after, count] = bind(a, ['before', 'after', 'count', 'type']); let k = wb.sheets.length + 1; while (wb.sheet('Sheet' + k)) k++; const s = wb.addSheet('Sheet' + k); wb.sheets.pop();
      let at = this.b.activeSi; if (has(after)) at = after.si + 1; else if (has(before)) at = before.si; wb.sheets.splice(at, 0, s); wb.invalidate(); this.b.activeSi = at; this.h.sel = { book: this.b, si: at, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; this.h.log.push({ kind: 'addsheet', name: s.name }); return new Sheet(this.h, this.b, at); }
    if (n === 'select' || n === 'delete') return EMPTY; throw vbErr(438); }
  dflt(a) { const k = deref(a[0].v); const wb = this.b.wb; if (typeof k === 'number') { if (k < 1 || k > wb.sheets.length) throw vbErr(9); return new Sheet(this.h, this.b, k - 1); } const i = wb.sheets.findIndex(s => lc(s.name) === lc(toStr(k))); if (i < 0) throw vbErr(9); return new Sheet(this.h, this.b, i); }
  each() { return this.b.wb.sheets.map((s, i) => new Sheet(this.h, this.b, i)); }
}
class Sheet extends VObject {
  constructor(h, b, si) { super(); this.h = h; this.b = b; this.si = si; this.s = b.wb.sheets[si]; if (!this.s) throw vbErr(9); }
  get tn() { return 'Worksheet'; } same(o) { return o instanceof Sheet && o.s === this.s; }
  idx() { const i = this.b.wb.sheets.indexOf(this.s); if (i < 0) throw vbErr(424, 'Object required: the sheet was deleted'); this.si = i; return i; }
  rng(g, rows) { return new Rng(this.h, this.b, this.idx(), g, rows); }
  get(n, a, rt) { const h = this.h, wb = this.b.wb; const si = this.idx();
    switch (n) {
      case 'name': return this.s.name; case 'codename': return this.s.code || ('Sheet' + (si + 1)); case 'index': return si + 1; case 'parent': return this.b; case 'visible': return this.s.hiddenSheet ? 0 : -1;
      case 'range': { if (!a.length) throw vbErr(449); const [x, y] = bind(a, ['cell1', 'cell2']); if (has(y)) { const p = x instanceof Rng ? x.g : parseRef(toStr(x), this), q = y instanceof Rng ? y.g : parseRef(toStr(y), this); return this.rng({ r1: Math.min(p.r1, q.r1), c1: Math.min(p.c1, q.c1), r2: Math.max(p.r2, q.r2), c2: Math.max(p.c2, q.c2) }); } if (x instanceof Rng) return x; if (/,/.test(toStr(x))) return new MultiRng(toStr(x).split(',').map(t => this.rng(parseRef(t, this)))); return this.rng(parseRef(toStr(x), this)); }
      case 'cells': { if (!a.length) return this.rng({ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }); const [r, c] = bind(a, ['rowindex', 'columnindex']); if (!has(c)) { const k = Math.round(toNum(r)) - 1; return this.rng({ r1: Math.floor(k / MAXC), c1: k % MAXC, r2: Math.floor(k / MAXC), c2: k % MAXC }); } const rr = Math.round(toNum(r)) - 1; const cc = colArg(c); if (rr < 0 || rr >= MAXR || cc < 0 || cc >= MAXC) throw vbErr(1004); return this.rng({ r1: rr, c1: cc, r2: rr, c2: cc }); }
      case 'rows': { if (!a.length) return this.rng({ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }, null).mode('rows'); const k = deref(a[0].v); if (typeof k === 'string') { const m = /^(\d+)(?::(\d+))?$/.exec(k); if (!m) throw vbErr(13); return this.rng({ r1: +m[1] - 1, c1: 0, r2: +(m[2] || m[1]) - 1, c2: MAXC - 1 }); } const r = Math.round(toNum(k)) - 1; return this.rng({ r1: r, c1: 0, r2: r, c2: MAXC - 1 }); }
      case 'columns': { if (!a.length) return this.rng({ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }).mode('cols'); const k = deref(a[0].v); if (typeof k === 'string') { const m = /^([A-Z]+)(?::([A-Z]+))?$/i.exec(k); if (!m) throw vbErr(13); return this.rng({ r1: 0, c1: colIndex(m[1]), r2: MAXR - 1, c2: colIndex(m[2] || m[1]) }); } const c = Math.round(toNum(k)) - 1; return this.rng({ r1: 0, c1: c, r2: MAXR - 1, c2: c }); }
      case 'usedrange': { const u = wb.usedRange(this.s) || { r1: 0, c1: 0, r2: 0, c2: 0 }; return this.rng(u); }
      case 'activate': case 'select': h.tick('select', 1, rt); this.b.activeSi = si; h.active = this.b; if (h.sel.book !== this.b || h.sel.si !== si) h.sel = { book: this.b, si, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; return EMPTY;
      case 'delete': { if (wb.sheets.length === 1) throw vbErr(1004, 'A workbook must contain at least one visible worksheet.'); if (h.alerts) return { ui: 'msgbox', prompt: 'Microsoft Excel will permanently delete this sheet. Do you want to continue?', buttons: 1 + 48, title: 'Microsoft Excel', after: r => { if (r === 1) this.remove(); return r === 1; } }; this.remove(); return true; }
      case 'copy': { const [before, after] = bind(a, ['before', 'after']); const copy = JSON.parse(JSON.stringify(wb.toSpec().sheets[si]));
        if (!has(before) && !has(after)) { copy.name = this.s.name; let k = 1; while (h.bookByName('Book' + k)) k++; const nb = new Book(h, new Workbook({ sheets: [copy] }), 'Book' + k, ''); h.books.push(nb); h.active = nb; h.sel = { book: nb, si: 0, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; return EMPTY; } let nm = this.s.name + ' (2)'; let k = 2; while (wb.sheet(nm)) nm = `${this.s.name} (${++k})`; copy.name = nm; const tmp = new Workbook({ sheets: [copy] }); const ns = tmp.sheets[0]; ns.name = nm; let at = has(after) ? after.idx() + 1 : has(before) ? before.idx() : wb.sheets.length; wb.sheets.splice(at, 0, ns); wb.invalidate(); this.b.activeSi = at; return EMPTY; }
      case 'listobjects': { const c = new Tables(h, this); return a.length ? c.dflt(a, rt) : c; }
      case 'autofiltermode': return !!this.s.filter; case 'filtermode': return !!(this.s.filter && this.s.hidden && this.s.hidden.size);
      case 'showalldata': if (this.s.hidden) this.s.hidden.clear(); if (this.s.filter) this.s.filter.fields = {}; return EMPTY;
      case 'pagesetup': return new PageSetup(this.s); case 'printarea': return (this.s.page || {}).printarea || '';
      case 'autofilter': return new AutoFilterObj(h, this); case 'sort': return this.s.sortObj || (this.s.sortObj = new SortObj(h, this));
      case 'exportasfixedformat': { const [ty, fn] = bind(a, ['type', 'filename', 'quality', 'includedocproperties', 'ignoreprintareas', 'from', 'to', 'openafterpublish']); return exportPdf(h, this.b, si, fn); }
      case 'protect': this.s.protectedBy = true; return EMPTY; case 'unprotect': this.s.protectedBy = false; return EMPTY; case 'protectcontents': return !!this.s.protectedBy;
      case 'calculate': wb.invalidate(); return EMPTY; case 'printout': case 'printpreview': return EMPTY;
      case 'paste': { const [dest] = bind(a, ['destination', 'link']); const to = has(dest) ? dest : new Rng(h, h.sel.book, h.sel.si, { r1: h.sel.ar, c1: h.sel.ac, r2: h.sel.ar, c2: h.sel.ac }); if (!h.clip) throw vbErr(1004, 'Paste method of Worksheet class failed'); pasteFrom(h, h.clip, to, 'all'); return EMPTY; }
      case 'next': { const i = si + 1; if (i >= wb.sheets.length) return NOTHING; return new Sheet(h, this.b, i); } case 'previous': return si ? new Sheet(h, this.b, si - 1) : NOTHING;
      case 'move': return EMPTY; case 'shapes': return new Shapes(h, this);
      case 'evaluate': return this.b.evaluate(toStr(a[0].v), si);
    } throw vbErr(438); }
  put(n, a, v) { const wb = this.b.wb; if (n === 'name') { const nm = toStr(v); if (!nm || nm.length > 31 || /[\\/?*[\]:]/.test(nm)) throw vbErr(1004, "You typed an invalid name for a sheet or chart."); if (wb.sheets.some(s => s !== this.s && lc(s.name) === lc(nm))) throw vbErr(1004, 'That name is already taken. Try a different one.'); const old = this.s.name; this.s.name = nm; wb.invalidate(); this.h.log.push({ kind: 'rename', from: old, to: nm }); return; } if (n === 'visible') { this.s.hiddenSheet = toNum(v) !== -1 && v !== true; return; } if (n === 'autofiltermode') { if (!toBool(v)) { this.s.filter = null; if (this.s.hidden) this.s.hidden.clear(); } return; } if (n === 'tab' || n === 'displaypagebreaks' || n === 'enablecalculation') return; throw vbErr(438); }
  remove() { const wb = this.b.wb; const i = this.idx(); wb.sheets.splice(i, 1); wb.invalidate(); this.b.activeSi = Math.min(this.b.activeSi, wb.sheets.length - 1); if (this.h.sel.book === this.b) this.h.sel = { book: this.b, si: this.b.activeSi, g: { r1: 0, c1: 0, r2: 0, c2: 0 }, ar: 0, ac: 0 }; this.h.log.push({ kind: 'deletesheet', name: this.s.name }); }
}
// Range("A1,A3:D3,A8:D8"): several areas at once. Setting a property (or a Font/Interior/Borders property) sets it on each area; reading gives the first area's.
class MultiRng extends VObject { constructor(parts) { super(); this.parts = parts; } get tn() { return 'Range'; }
  get(n, a = [], rt) { const P = this.parts; switch (n) {
      case 'font': case 'interior': case 'borders': { const subs = P.map(r => r.get(n, a, rt)); return new MultiProxy(subs); }
      case 'areas': return new MultiAreas(P); case 'count': case 'countlarge': return P.reduce((k, r) => k + toNum(r.get('count', [])), 0);
      case 'address': return P.map(r => toStr(r.get('address', a))).join(',');
      case 'clearcontents': case 'clear': case 'clearformats': case 'autofit': { for (const r of P) r.get(n, a, rt); return EMPTY; }
      default: return P[0].get(n, a, rt); } }
  put(n, a, v, rt, isSet) { for (const r of this.parts) r.put(n, a, v, rt, isSet); } }
class MultiProxy extends VObject { constructor(subs) { super(); this.subs = subs; } get tn() { return this.subs[0].tn; } get(n, a, rt) { return this.subs[0].get(n, a, rt); } put(n, a, v, rt, isSet) { for (const x of this.subs) x.put(n, a, v, rt, isSet); } }
class MultiAreas extends VObject { constructor(P) { super(); this.P = P; } get tn() { return 'Areas'; } get(n, a = []) { if (n === 'count') return this.P.length; if (n === 'item' || n === '') return this.P[Math.round(toNum(a[0])) - 1]; throw vbErr(438); } }
class Shapes extends VObject { constructor(h, s) { super(); this.h = h; this.s = s; } get tn() { return 'Shapes'; } get(n) { if (n === 'count') return (this.s.s.buttons || []).length; throw vbErr(438); } }
function colArg(c) { c = deref(c); if (typeof c === 'string' && /^[A-Z]+$/i.test(c)) return colIndex(c); return Math.round(toNum(c)) - 1; }
function parseRef(t, sheet) { t = t.trim().replace(/\$/g, ''); let m;
  if ((m = /^([A-Z]+)(\d+)$/i.exec(t))) return { r1: +m[2] - 1, c1: colIndex(m[1]), r2: +m[2] - 1, c2: colIndex(m[1]) };
  if ((m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(t))) { const a = { r: +m[2] - 1, c: colIndex(m[1]) }, b = { r: +m[4] - 1, c: colIndex(m[3]) }; return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }; }
  if ((m = /^([A-Z]+):([A-Z]+)$/i.exec(t))) return { r1: 0, c1: Math.min(colIndex(m[1]), colIndex(m[2])), r2: MAXR - 1, c2: Math.max(colIndex(m[1]), colIndex(m[2])) };
  if ((m = /^(\d+):(\d+)$/.exec(t))) return { r1: +m[1] - 1, c1: 0, r2: +m[2] - 1, c2: MAXC - 1 };
  if (sheet) { const tb = sheet.b.wb.tableByName(t); if (tb) return parseRange(tb.t.ref); const nm = sheet.b.wb.names[lc(t)]; if (nm) { const r = String(nm.ref).replace(/^=/, '').replace(/^'?[^'!]*'?!/, ''); return parseRef(r); } }
  throw vbErr(1004, "Method 'Range' of object '_Worksheet' failed"); }

// ---------------------------------------------------------------- cell values in and out
function fromCell(v, st) { if (v === null || v === undefined || v === '') return EMPTY; if (isErr(v)) return v; if (typeof v === 'number' && st && st.fmt && isDateFmt(st.fmt)) return new VDate(v); if (v && v.range) return v.single !== undefined ? fromCell(v.single) : EMPTY; return v; }
function toCell(wb, si, r, c, v) {
  v = deref(v); if (v === EMPTY || v === MISSING || v === NULLV || v === undefined) { wb.clear(si, r, c, true); return; }
  if (v instanceof VDate) { const cur = wb.style(si, r, c).fmt; wb.setRaw(wb.sheets[si], r, c, v.n, true); if (!cur || !isDateFmt(cur)) wb.setStyle(si, r, c, { fmt: v.n % 1 ? 'dd-mm-yyyy hh:mm' : 'dd-mm-yyyy' }); return; }
  if (typeof v === 'string') { const st = wb.style(si, r, c); if (st.fmt === '@') { wb.setRaw(wb.sheets[si], r, c, v, true); return; } wb.setInput(si, r, c, v); return; }
  if (typeof v === 'number' || typeof v === 'boolean') { wb.setRaw(wb.sheets[si], r, c, v, true); return; }
  if (v && v.code) { wb.setRaw(wb.sheets[si], r, c, v, true); return; }
  if (isObj(v)) throw vbErr(13); throw vbErr(13); }
function grow(wb, si, g) { const s = wb.sheets[si]; if (g.r2 < MAXR - 1 && g.r2 + 2 > s.rows) s.rows = Math.min(MAXR, g.r2 + 21); if (g.c2 < MAXC - 1 && g.c2 + 2 > s.cols) s.cols = Math.min(MAXC, g.c2 + 6); }

// ---------------------------------------------------------------- Range
export class Rng extends VObject {
  // rows: an explicit list of row numbers (visible cells after a filter, blanks, a Union of rows); null means every row g.r1..g.r2
  constructor(h, b, si, g, rows) { super(); this.h = h; this.book = b; this.si = si; this.g = { ...g }; this.rows = rows || null; this._mode = null; }
  get tn() { return 'Range'; } mode(m) { this._mode = m; return this; }
  get wb() { return this.book.wb; } get sheet() { return this.wb.sheets[this.si]; }
  same(o) { return o instanceof Rng && o.book === this.book && o.si === this.si && ['r1', 'c1', 'r2', 'c2'].every(k => o.g[k] === this.g[k]); }
  rowList() { if (this.rows) return this.rows.slice(); const out = []; const lim = this.limitRows(); for (let r = this.g.r1; r <= lim; r++) out.push(r); return out; }
  limitRows() { const u = this.wb.usedRange(this.sheet); return Math.min(this.g.r2, u ? Math.max(u.r2, this.g.r1) : this.g.r1); } // whole columns stop at the data
  limitCols() { const u = this.wb.usedRange(this.sheet); return Math.min(this.g.c2, u ? Math.max(u.c2, this.g.c1) : this.g.c1); }
  cells(fn, opts = {}) { const rows = this.rows || null; const r2 = opts.all ? this.g.r2 : this.limitRows(), c2 = opts.all ? this.g.c2 : this.limitCols(); for (const r of rows || range(this.g.r1, r2)) for (let c = this.g.c1; c <= c2; c++) fn(r, c); }
  address(ra = true, ca = true, ext) { const one = (r, c) => (ca ? '$' : '') + colName(c) + (ra ? '$' : '') + (r + 1); const g = this.g; let s;
    if (g.c1 === 0 && g.c2 === MAXC - 1) s = `${ra ? '$' : ''}${g.r1 + 1}:${ra ? '$' : ''}${g.r2 + 1}`; else if (g.r1 === 0 && g.r2 === MAXR - 1) s = `${ca ? '$' : ''}${colName(g.c1)}:${ca ? '$' : ''}${colName(g.c2)}`; else s = g.r1 === g.r2 && g.c1 === g.c2 ? one(g.r1, g.c1) : one(g.r1, g.c1) + ':' + one(g.r2, g.c2);
    if (this.rows) { const parts = groupRows(this.rows).map(([a, b]) => (g.c1 === 0 && g.c2 === MAXC - 1) ? `${ra ? '$' : ''}${a + 1}:${ra ? '$' : ''}${b + 1}` : one(a, g.c1) + (a === b && g.c1 === g.c2 ? '' : ':' + one(b, g.c2))); s = parts.join(','); }
    return ext ? `[${this.book.name}]${this.sheet.name}!${s}` : s; }
  value(v2) { const g = this.g; if (g.r1 === g.r2 && g.c1 === g.c2 && !this.rows) { this.h.tick('read', 1, this.rt); const v = this.wb.value(this.si, g.r1, g.c1); return v2 ? (v === null || v === undefined ? EMPTY : (isErr(v) ? v : v)) : fromCell(v, this.wb.style(this.si, g.r1, g.c1)); }
    const rows = this.rows || range(g.r1, Math.min(g.r2, g.r1 + 199999)); const nc = Math.min(g.c2, g.c1 + 16383) - g.c1 + 1; const a = new VArray([[1, rows.length], [1, nc]], 'variant'); let k = 0; for (const r of rows) for (let c = 0; c < nc; c++) { const v = this.wb.value(this.si, r, g.c1 + c); a.d[k++] = v2 ? (v === null || v === undefined ? EMPTY : v) : fromCell(v, this.wb.style(this.si, r, g.c1 + c)); } this.h.tick('bulk', 1, this.rt); return a; }
  setValue(v, asFormula) { const g = this.g; const wb = this.wb; if (this.h.inUdf) throw vbErr(1004, 'A function used in a cell can’t change cells.'); if (this.sheet.protectedBy) throw vbErr(1004, 'The cell or chart you’re trying to change is on a protected sheet.'); grow(wb, this.si, g);
    if (v instanceof VArray) { const two = v.b.length === 2; const nr = two ? v.b[0][1] - v.b[0][0] + 1 : 1, nc = two ? v.b[1][1] - v.b[1][0] + 1 : v.b[0][1] - v.b[0][0] + 1; const R = this.rows || range(g.r1, g.r2); let i = 0; for (const r of R) { for (let c = g.c1; c <= g.c2; c++) { const rr = i, cc = c - g.c1; const x = rr < nr && cc < nc ? (two ? v.get1([v.b[0][0] + rr, v.b[1][0] + cc]) : v.get1([v.b[0][0] + cc])) : { code: '#N/A' }; toCell(wb, this.si, r, c, x && x.code ? x : x); } i++; } this.h.tick('bulk', 1, this.rt); wb.invalidate(); this.h.changed(this); return; }
    let n = 0; const R = this.rows || range(g.r1, g.r2); const c2 = g.c2 - g.c1 > 16383 ? g.c1 : g.c2; const v0 = deref(v), shift = typeof v0 === 'string' && v0.startsWith('=') && (R.length > 1 || c2 > g.c1); const r0 = R[0]; for (const r of R) for (let c = g.c1; c <= c2; c++) { const x = shift && (r !== r0 || c !== g.c1) ? shiftFormula(v0, r - r0, c - g.c1) : v0; if (asFormula && typeof x === 'string') wb.setInput(this.si, r, c, x); else toCell(wb, this.si, r, c, typeof x === 'string' ? x : v); n++; } /* one formula into many cells: relative references move, as when you fill */ this.h.tick('write', n, this.rt); wb.invalidate(); this.h.changed(this); }
  style(st) { const wb = this.wb; this.cells((r, c) => wb.setStyle(this.si, r, c, st), { all: this.g.r2 - this.g.r1 < 5000 && this.g.c2 - this.g.c1 < 200 }); let n = 0; this.cells(() => n++); this.h.tick('write', Math.max(1, n / 4), this.rt); }
  first() { return this.wb.style(this.si, this.rows ? this.rows[0] : this.g.r1, this.g.c1); }
  sub(g) { return new Rng(this.h, this.book, this.si, g); }
  cell(r, c) { return this.sub({ r1: r, c1: c, r2: r, c2: c }); }
  count() { const nc = this.g.c2 - this.g.c1 + 1; const nr = this.rows ? this.rows.length : this.g.r2 - this.g.r1 + 1; if (this._mode === 'rows') return nr; if (this._mode === 'cols') return nc; return nr * nc; }
  get(n, a, rt) { this.rt = rt; const h = this.h, wb = this.wb, g = this.g;
    switch (n) {
      case 'value': return this.value(false); case 'value2': return this.value(true);
      case 'formula': case 'formulalocal': case 'formular1c1': { if (g.r1 === g.r2 && g.c1 === g.c2) { const t = wb.input(this.si, g.r1, g.c1); return t; } const a2 = new VArray([[1, g.r2 - g.r1 + 1], [1, g.c2 - g.c1 + 1]]); let k = 0; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) a2.d[k++] = wb.input(this.si, r, c); return a2; }
      case 'text': return wb.display(this.si, g.r1, g.c1);
      case 'hasformula': return !!wb.formula(this.si, g.r1, g.c1);
      case 'address': { const [ra, ca, rs, ext] = bind(a, ['rowabsolute', 'columnabsolute', 'referencestyle', 'external']); return this.address(has(ra) ? toBool(ra) : true, has(ca) ? toBool(ca) : true, has(ext) && toBool(ext)); }
      case 'row': return (this.rows ? this.rows[0] : g.r1) + 1; case 'column': return g.c1 + 1; case 'count': case 'countlarge': return this.count();
      case 'rows': { if (!a.length) return this.sub(g).withRows(this.rows).mode('rows'); const k = Math.round(toNum(a[0].v)) - 1; const R = this.rows || null; const r = R ? R[k] : g.r1 + k; if (r === undefined) throw vbErr(1004); return this.sub({ r1: r, c1: g.c1, r2: r, c2: g.c2 }); }
      case 'columns': { if (!a.length) return this.sub(g).mode('cols'); const kv = deref(a[0].v); if (typeof kv === 'string' && /^[A-Z]+(:[A-Z]+)?$/i.test(kv)) { const [x, y] = kv.split(':'); return this.sub({ r1: g.r1, c1: g.c1 + colIndex(x), r2: g.r2, c2: g.c1 + colIndex(y || x) }); } const k = Math.round(toNum(kv)) - 1; return this.sub({ r1: g.r1, c1: g.c1 + k, r2: g.r2, c2: g.c1 + k }); } // Columns("G:H") of a range counts from the range's own first column, like Excel
      case 'cells': case 'item': { if (!a.length) return this; const [r, c] = bind(a, ['rowindex', 'columnindex']); if (!has(c)) { const k = Math.round(toNum(r)) - 1; const w = g.c2 - g.c1 + 1; if (this._mode === 'rows') return this.sub({ r1: g.r1 + k, c1: g.c1, r2: g.r1 + k, c2: g.c2 }); if (this._mode === 'cols') return this.sub({ r1: g.r1, c1: g.c1 + k, r2: g.r2, c2: g.c1 + k }); return this.cell(g.r1 + Math.floor(k / w), g.c1 + k % w); } return this.cell(g.r1 + Math.round(toNum(r)) - 1, g.c1 + colArg(c)); }
      case 'range': { const [x, y] = bind(a, ['cell1', 'cell2']); const p = parseRef(toStr(x)); return this.sub({ r1: g.r1 + p.r1, c1: g.c1 + p.c1, r2: g.r1 + p.r2, c2: g.c1 + p.c2 }); }
      case 'offset': { const [dr, dc] = bind(a, ['rowoffset', 'columnoffset']); const x = has(dr) ? Math.round(toNum(dr)) : 0, y = has(dc) ? Math.round(toNum(dc)) : 0; if (g.r1 + x < 0 || g.c1 + y < 0 || g.r2 + x >= MAXR || g.c2 + y >= MAXC) throw vbErr(1004); return new Rng(h, this.book, this.si, { r1: g.r1 + x, c1: g.c1 + y, r2: g.r2 + x, c2: g.c2 + y }, this.rows ? this.rows.map(r => r + x) : null); }
      case 'resize': { const [nr, nc] = bind(a, ['rowsize', 'columnsize']); const R = has(nr) ? Math.round(toNum(nr)) : g.r2 - g.r1 + 1, C = has(nc) ? Math.round(toNum(nc)) : g.c2 - g.c1 + 1; if (R < 1 || C < 1) throw vbErr(1004); return this.sub({ r1: g.r1, c1: g.c1, r2: g.r1 + R - 1, c2: g.c1 + C - 1 }); }
      case 'end': { const d = toNum(a[0].v); return this.end(d); }
      case 'currentregion': return this.sub(currentRegion(wb, this.si, g.r1, g.c1));
      case 'entirerow': return new Rng(h, this.book, this.si, { r1: g.r1, c1: 0, r2: g.r2, c2: MAXC - 1 }, this.rows);
      case 'entirecolumn': return this.sub({ r1: 0, c1: g.c1, r2: MAXR - 1, c2: g.c2 });
      case 'worksheet': case 'parent': return new Sheet(h, this.book, this.si);
      case 'font': return new Font(this); case 'interior': return new Interior(this); case 'borders': return new Borders(this);
      case 'numberformat': case 'numberformatlocal': return this.first().fmt || 'General';
      case 'horizontalalignment': return ALIGN_BACK[this.first().align] || 1; case 'verticalalignment': return -4107; case 'wraptext': return !!this.first().wrap;
      case 'columnwidth': return (this.sheet.widths[g.c1] || 64) / 7.5; case 'rowheight': return 15; case 'hidden': return !!(this.sheet.hidden && this.sheet.hidden.has(g.r1));
      case 'clearcontents': if (this.sheet.protectedBy) throw vbErr(1004, 'The cell or chart you’re trying to change is on a protected sheet.'); this.cells((r, c) => wb.clear(this.si, r, c, true)); this.h.changed(this); return EMPTY;
      case 'clear': this.cells((r, c) => wb.clear(this.si, r, c, false)); this.h.changed(this); return EMPTY;
      case 'clearformats': this.cells((r, c) => { const o = wb.cellObj(this.sheet, r, c); if (o) o.s = null; }); wb.invalidate(); return EMPTY;
      case 'select': case 'activate': { h.tick('select', 1, rt); if (n === 'select' || !(h.sel.si === this.si && h.sel.book === this.book && inside(h.sel.g, g.r1, g.c1))) h.sel = { book: this.book, si: this.si, g: { ...g, r2: Math.min(g.r2, MAXR - 1) }, ar: g.r1, ac: g.c1 }; else { h.sel.ar = g.r1; h.sel.ac = g.c1; } this.book.activeSi = this.si; h.active = this.book; return EMPTY; }
      case 'copy': { const [dest] = bind(a, ['destination']); if (has(dest)) { if (!(dest instanceof Rng)) throw vbErr(424); pasteFrom(h, this.snap(), dest, 'all'); return EMPTY; } h.clip = this.snap(); return EMPTY; }
      case 'cut': { const [dest] = bind(a, ['destination']); const sn = this.snap(); if (has(dest)) { this.cells((r, c) => wb.clear(this.si, r, c, false)); pasteFrom(h, sn, dest, 'all'); return EMPTY; } h.clip = sn; return EMPTY; }
      case 'pastespecial': { const [paste] = bind(a, ['paste', 'operation', 'skipblanks', 'transpose']); if (!h.clip) throw vbErr(1004, 'PasteSpecial method of Range class failed'); const k = has(paste) ? toNum(paste) : -4104; pasteFrom(h, h.clip, this, k === -4163 || k === 12 ? 'values' : k === -4122 ? 'formats' : k === -4123 ? 'formulas' : 'all'); return EMPTY; }
      case 'delete': { const [sh] = bind(a, ['shift']); deleteCells(this, has(sh) ? toNum(sh) : null); return EMPTY; }
      case 'insert': { const [sh] = bind(a, ['shift', 'copyorigin']); insertCells(this, has(sh) ? toNum(sh) : null); return EMPTY; }
      case 'autofit': { for (let c = g.c1; c <= Math.min(g.c2, this.limitCols()); c++) { let w = 0; const r2 = this.limitRows(); for (let r = 0; r <= r2; r++) w = Math.max(w, String(wb.display(this.si, r, c)).length); this.sheet.widths[c] = Math.max(40, Math.min(400, Math.round(w * 7.5 + 12))); } return EMPTY; }
      case 'sort': { const [k1, o1, k2, o2, , , , hd] = bind(a, ['key1', 'order1', 'key2', 'type', 'order2', 'key3', 'order3', 'header']); sortRange(this, [[k1, o1], [k2, has(o2) ? o2 : MISSING]].filter(x => has(x[0])), has(hd) ? toNum(hd) === 1 : false); return EMPTY; }
      case 'autofilter': { const [field, c1, op, c2, vis] = bind(a, ['field', 'criteria1', 'operator', 'criteria2', 'visibledropdown']); autoFilter(this, field, c1, op, c2); return true; }
      case 'specialcells': { const [ty, vv] = bind(a, ['type', 'value']); return specialCells(this, toNum(ty), vv); }
      case 'find': { const [what, after, lookin, lookat, , , match] = bind(a, ['what', 'after', 'lookin', 'lookat', 'searchorder', 'searchdirection', 'matchcase']); const w = toStr(deref(what)); const whole = has(lookat) && toNum(lookat) === 1; const mc = has(match) && toBool(match); const eq = x => { let s = String(x === null || x === undefined ? '' : (typeof x === 'number' ? formatValue(x) : x)); let t = w; if (!mc) { s = s.toLowerCase(); t = t.toLowerCase(); } return whole ? s === t : s.includes(t); };
        let hit = null; this.cells((r, c) => { if (!hit && wb.value(this.si, r, c) !== null && eq(wb.value(this.si, r, c))) hit = { r, c }; }); return hit ? this.cell(hit.r, hit.c) : NOTHING; }
      case 'removeduplicates': { const [cols, hd] = bind(a, ['columns', 'header']); removeDuplicates(this, cols, has(hd) && toNum(hd) === 1); return EMPTY; }
      case 'merge': case 'unmerge': case 'calculate': case 'printout': return EMPTY;
      case 'areas': return new Areas(this); case 'name': return EMPTY;
      case 'listobject': { const tb = (this.sheet.tables || []).find(t => { const q = parseRange(t.ref); return inside(q, g.r1, g.c1); }); return tb ? new Table(h, new Sheet(h, this.book, this.si), tb) : NOTHING; }
      case 'exportasfixedformat': { const [ty, fn] = bind(a, ['type', 'filename']); return exportPdf(h, this.book, this.si, fn, { ...g, r2: this.limitRows(), c2: this.limitCols() }); }
      case 'formulaarray': return wb.input(this.si, g.r1, g.c1);
    } throw vbErr(438); }
  withRows(rows) { this.rows = rows ? rows.slice() : null; return this; }
  put(n, a, v, rt) { this.rt = rt; switch (n) {
      case 'value': case 'value2': this.setValue(v); return; case 'formula': case 'formulalocal': case 'formulaarray': this.setValue(v, true); return;
      case 'formular1c1': this.setValue(r1c1(toStr(v), this.g), true); return;
      case 'numberformat': case 'numberformatlocal': this.style({ fmt: toStr(v) }); return;
      case 'horizontalalignment': this.style({ align: ALIGN[toNum(v)] || null }); return; case 'verticalalignment': return; case 'wraptext': this.style({ wrap: toBool(v) || null }); return;
      case 'columnwidth': for (let c = this.g.c1; c <= Math.min(this.g.c2, this.g.c1 + 200); c++) this.sheet.widths[c] = Math.round(toNum(v) * 7.5 + 5); return; case 'rowheight': return;
      case 'hidden': { const s = this.sheet; s.hidden = s.hidden || new Set(); for (const r of this.rowList()) { if (toBool(v)) s.hidden.add(r); else s.hidden.delete(r); } return; }
      case 'name': { this.wb.defineName(toStr(v), `'${this.sheet.name}'!${this.address(true, true)}`); return; }
      case 'mergecells': case 'indentlevel': case 'orientation': case 'locked': return;
    } throw vbErr(438); }
  hasArgProp(n) { return n === 'value' || n === 'value2'; }
  dflt(a, rt) { this.rt = rt; if (!a.length) return this.value(false); return this.get('item', a, rt); }
  putDflt(a, v, rt) { this.rt = rt; if (!a.length) { this.setValue(v); return; } this.get('item', a, rt).setValue(v); }
  each() { if (this._mode === 'rows') return this.rowList().map(r => this.sub({ r1: r, c1: this.g.c1, r2: r, c2: this.g.c2 })); if (this._mode === 'cols') { const out = []; for (let c = this.g.c1; c <= this.limitCols(); c++) out.push(this.sub({ r1: this.g.r1, c1: c, r2: this.g.r2, c2: c })); return out; } const out = []; this.cells((r, c) => out.push(this.cell(r, c))); return out; }
  snap() { const g = this.g, wb = this.wb; const rows = this.rows || range(g.r1, this.limitRows()); const c2 = this.limitCols(); return { book: this.book, si: this.si, c1: g.c1, rows, r1: rows[0], cells: rows.map(r => { const out = []; for (let c = g.c1; c <= c2; c++) { const o = wb.cellObj(wb.sheets[this.si], r, c); out.push(o ? { v: o.v, f: o.f, s: o.s ? { ...o.s } : null, val: wb.value(this.si, r, c) } : null); } return out; }), widths: Array.from({ length: c2 - g.c1 + 1 }, (_, k) => wb.sheets[this.si].widths[g.c1 + k]) }; }
  end(d) { const wb = this.wb, si = this.si, s = this.sheet; let r = this.g.r1, c = this.g.c1; const filled = (rr, cc) => rr >= 0 && cc >= 0 && rr < MAXR && cc < MAXC && !wb.isEmpty(si, rr, cc);
    const [dr, dc] = d === -4162 ? [-1, 0] : d === -4121 ? [1, 0] : d === -4159 ? [0, -1] : d === -4161 ? [0, 1] : (() => { throw vbErr(1004); })();
    const u = wb.usedRange(s) || { r1: 0, c1: 0, r2: 0, c2: 0 };
    if (filled(r, c) && filled(r + dr, c + dc)) { while (filled(r + dr, c + dc)) { r += dr; c += dc; } return this.cell(r, c); }
    r += dr; c += dc; if (dr === -1 && r > u.r2) r = u.r2; if (dc === -1 && c > u.c2) c = u.c2;
    while (r >= 0 && c >= 0 && r < MAXR && c < MAXC && !filled(r, c)) { if ((dr === 1 && r > u.r2) || (dc === 1 && c > u.c2)) { r = dr ? MAXR - 1 : r; c = dc ? MAXC - 1 : c; return this.cell(r, c); } r += dr; c += dc; }
    r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c)); return this.cell(r, c); }
}
const range = (a, b) => { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; };
const inside = (g, r, c) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;
function groupRows(rows) { const out = []; for (const r of rows) { const last = out[out.length - 1]; if (last && r === last[1] + 1) last[1] = r; else out.push([r, r]); } return out; }
function r1c1(f, g) { return f.replace(/R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?/g, (m, R, C) => { const rel = x => x && x[0] === '['; const row = R ? (rel(R) ? g.r1 + +R.slice(1, -1) : +R - 1) : g.r1; const col = C ? (rel(C) ? g.c1 + +C.slice(1, -1) : +C - 1) : g.c1; return (R && !rel(R) ? '' : '') + colName(col) + (row + 1); }); }
export function currentRegion(wb, si, r, c) { let g = { r1: r, c1: c, r2: r, c2: c }; const filled = (rr, cc) => rr >= 0 && cc >= 0 && !wb.isEmpty(si, rr, cc); if (!filled(r, c)) { let any = false; for (let dr = -1; dr <= 1; dr++) for (let dc = -1; dc <= 1; dc++) if (filled(r + dr, c + dc)) any = true; if (!any) return g; }
  for (let k = 0; k < 2000; k++) { let ch = false; const rowHas = rr => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (filled(rr, cc)) return true; return false; }; const colHas = cc => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (filled(rr, cc)) return true; return false; };
    if (g.r1 > 0 && rowHas(g.r1 - 1)) { g.r1--; ch = true; } if (rowHas(g.r2 + 1)) { g.r2++; ch = true; } if (g.c1 > 0 && colHas(g.c1 - 1)) { g.c1--; ch = true; } if (colHas(g.c2 + 1)) { g.c2++; ch = true; } if (!ch) break; } return g; }
function pasteFrom(h, sn, dest, what) { const wb = dest.wb, si = dest.si; const r0 = dest.g.r1, c0 = dest.g.c1; grow(wb, si, { r2: r0 + sn.cells.length, c2: c0 + (sn.cells[0] || []).length });
  sn.cells.forEach((row, i) => row.forEach((cl, j) => { const r = r0 + i, c = c0 + j; const o = wb.cellObj(wb.sheets[si], r, c, true);
    if (what !== 'formats') { if (!cl) { o.v = null; o.f = null; } else if (what === 'values') { o.f = null; o.v = cl.val && cl.val.range ? null : cl.val; } else { o.f = cl.f ? shiftFormula(cl.f, r - sn.rows[i], c - (sn.c1 + j)) : null; o.v = cl.f ? null : cl.v; } }
    if (what !== 'values' && what !== 'formulas') o.s = cl && cl.s ? { ...cl.s } : (what === 'formats' ? null : o.s && what === 'all' ? null : o.s); }));
  if (what === 'all' && sn.widths) sn.widths.forEach((w, k) => { if (w && !wb.sheets[si].widths[c0 + k]) wb.sheets[si].widths[c0 + k] = w; });
  h.tick('bulk', 1, dest.rt); wb.invalidate(); }
function shiftRows(wb, si, from, n) { // move every cell at or below row `from` down by n (n < 0: up), rebuilding the cell map
  const s = wb.sheets[si]; const next = new Map(); for (const [k, cell] of s.cells) { const r = Math.floor(k / 16384), c = k % 16384; if (r < from) next.set(k, cell); else if (r + n >= from && n < 0 && r < from - n) continue; else next.set((r + n) * 16384 + c, cell); } s.cells = next;
  for (const t of s.tables || []) { const g = parseRange(t.ref); if (g.r2 >= from) { if (g.r1 >= from) { g.r1 += n; } g.r2 += n; t.ref = colName(g.c1) + (g.r1 + 1) + ':' + colName(g.c2) + (g.r2 + 1); } }
  if (s.hidden) s.hidden = new Set([...s.hidden].filter(r => r < from).concat([...s.hidden].filter(r => r >= from - (n < 0 ? n : 0)).map(r => r + n)).filter(r => r >= 0)); wb.invalidate(); }
function deleteCells(rng, shift) { const g = rng.g, wb = rng.wb, si = rng.si; if (rng.sheet.protectedBy) throw vbErr(1004, 'The cell or chart you’re trying to change is on a protected sheet.');
  const whole = g.c1 === 0 && g.c2 === MAXC - 1; const rows = rng.rowList(); if (whole || shift === -4162 || shift === null && g.c2 - g.c1 >= g.r2 - g.r1) {
    if (whole) { for (const [a, b] of groupRows(rows).reverse()) shiftRows(wb, si, b + 1, -(b - a + 1)); rng.h.tick('bulk', rows.length, rng.rt); return; }
    // shift cells up inside the columns
    const s = wb.sheets[si]; const u = wb.usedRange(s) || g; const n = g.r2 - g.r1 + 1; for (let c = g.c1; c <= g.c2; c++) { for (let r = g.r1; r <= u.r2; r++) { const src = wb.cellObj(s, r + n, c); const k = r * 16384 + c; if (src) s.cells.set(k, src); else s.cells.delete(k); } for (let r = u.r2 - n + 1; r <= u.r2; r++) if (r > g.r2 - n) s.cells.delete(r * 16384 + c); } wb.invalidate(); return; }
  const s = wb.sheets[si]; const n = g.c2 - g.c1 + 1; const u = wb.usedRange(s) || g; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= u.c2; c++) { const src = wb.cellObj(s, r, c + n); const k = r * 16384 + c; if (src) s.cells.set(k, src); else s.cells.delete(k); } wb.invalidate(); }
function insertCells(rng, shift) { const g = rng.g; const whole = g.c1 === 0 && g.c2 === MAXC - 1; if (whole || shift === -4121 || shift === null) { if (whole) { shiftRows(rng.wb, rng.si, g.r1, g.r2 - g.r1 + 1); return; } const s = rng.sheet, wb = rng.wb; const u = wb.usedRange(s) || g; const n = g.r2 - g.r1 + 1; for (let c = g.c1; c <= g.c2; c++) for (let r = u.r2; r >= g.r1; r--) { const src = wb.cellObj(s, r, c); const k = (r + n) * 16384 + c; if (src) s.cells.set(k, src); else s.cells.delete(k); s.cells.delete(r * 16384 + c); } wb.invalidate(); return; } throw vbErr(1004); }
function cmpCell(a, b) { const na = a === null || a === undefined || a === '', nb = b === null || b === undefined || b === ''; if (na || nb) return na && nb ? 0 : na ? 1 : -1; if (typeof a === 'number' && typeof b === 'number') return a - b; if (typeof a === 'number') return -1; if (typeof b === 'number') return 1; return String(a).localeCompare(String(b), 'en', { sensitivity: 'base' }); }
function sortRange(rng, keys, header) { const wb = rng.wb, si = rng.si, s = rng.sheet; const g = { ...rng.g, r2: rng.limitRows(), c2: rng.limitCols() }; const r0 = header ? g.r1 + 1 : g.r1;
  const kc = keys.map(([k, o]) => ({ c: k instanceof Rng ? k.g.c1 : colArg(k), desc: has(o) && toNum(o) === 2 }));
  const rows = []; for (let r = r0; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) { const o = wb.cellObj(s, r, c); cells.push(o ? { ...o } : null); } rows.push({ r, cells, key: kc.map(k => wb.value(si, r, k.c)) }); }
  rows.sort((a, b) => { for (let i = 0; i < kc.length; i++) { const d = cmpCell(a.key[i], b.key[i]); if (d) return kc[i].desc ? (a.key[i] === null || b.key[i] === null ? d : -d) : d; } return a.r - b.r; });
  rows.forEach((row, i) => row.cells.forEach((cl, j) => { const k = (r0 + i) * 16384 + g.c1 + j; if (cl) { if (cl.f) cl.f = shiftFormula(cl.f, r0 + i - row.r, 0); s.cells.set(k, cl); } else s.cells.delete(k); })); wb.invalidate(); }
class SortObj extends VObject { constructor(h, sh) { super(); this.h = h; this.sh = sh; this.fields = []; this.rng = null; this.header = 2; } get tn() { return 'Sort'; }
  get(n, a) { if (n === 'sortfields') return new SortFields(this); if (n === 'setrange') { this.rng = a[0].v; return EMPTY; } if (n === 'apply') { if (!this.rng) throw vbErr(1004, 'The sort reference is not valid.'); sortRange(this.rng, this.fields.map(f => [f.key, f.order]), this.header === 1); return EMPTY; } if (n === 'header') return this.header; if (['matchcase', 'orientation', 'sortmethod'].includes(n)) return EMPTY; throw vbErr(438); }
  put(n, a, v) { if (n === 'header') { this.header = toNum(v); return; } if (['matchcase', 'orientation', 'sortmethod'].includes(n)) return; throw vbErr(438); } }
class SortFields extends VObject { constructor(s) { super(); this.s = s; } get tn() { return 'SortFields'; } get(n, a) { if (n === 'clear') { this.s.fields = []; return EMPTY; } if (n === 'add' || n === 'add2') { const [key, on, order] = bind(a, ['key', 'sorton', 'order', 'customorder', 'dataoption']); this.s.fields.push({ key, order: has(order) ? order : 1 }); return EMPTY; } if (n === 'count') return this.s.fields.length; throw vbErr(438); } }
function critTest(c) { if (!has(c)) return () => true; if (c instanceof VArray) { const set = c.d.map(x => lc(toStr(x))); return v => set.includes(lc(fmtShown(v))); } const s = toStr(c); let m = /^(<>|>=|<=|=|>|<)?(.*)$/.exec(s); const op = m[1] || '='; const rhs = m[2]; const n = rhs.trim() !== '' && !isNaN(+rhs) ? +rhs : null;
  return v => { const blank = v === null || v === undefined || v === ''; if (op === '=' && rhs === '') return blank; if (op === '<>' && rhs === '') return !blank; if (n !== null && typeof v === 'number') return { '=': v === n, '<>': v !== n, '>': v > n, '<': v < n, '>=': v >= n, '<=': v <= n }[op]; const t = lc(fmtShown(v)); if (op === '=' || op === '<>') { const hit = /[*?]/.test(rhs) ? wild(rhs).test(t) : t === lc(rhs); return op === '=' ? hit : !hit; } return false; }; }
const fmtShown = v => v === null || v === undefined ? '' : typeof v === 'number' ? formatValue(v) : String(v);
function autoFilter(rng, field, c1, op, c2) { const s = rng.sheet, wb = rng.wb, si = rng.si; const reg = rng.g.r1 === rng.g.r2 && rng.g.c1 === rng.g.c2 ? currentRegion(wb, si, rng.g.r1, rng.g.c1) : { ...rng.g, r2: rng.limitRows(), c2: rng.limitCols() };
  if (!has(field)) { if (s.filter) { s.filter = null; s.hidden = new Set(); } else s.filter = { g: reg, fields: {} }; return; }
  if (!s.filter || s.filter.g.r1 !== reg.r1 || s.filter.g.c1 !== reg.c1) s.filter = { g: reg, fields: {} }; const f = Math.round(toNum(field)); if (f < 1 || f > reg.c2 - reg.c1 + 1) throw vbErr(1004, 'AutoFilter method of Range class failed');
  const t1 = critTest(c1), t2 = critTest(c2); const oper = has(op) ? toNum(op) : 1; s.filter.fields[f] = has(c1) ? (v => has(c2) ? (oper === 2 ? t1(v) || t2(v) : t1(v) && t2(v)) : t1(v)) : null;
  s.hidden = new Set(); for (let r = reg.r1 + 1; r <= reg.r2; r++) { for (const [k, test] of Object.entries(s.filter.fields)) if (test && !test(wb.value(si, r, reg.c1 + +k - 1))) { s.hidden.add(r); break; } } }
class AutoFilterObj extends VObject { constructor(h, sh) { super(); this.h = h; this.sh = sh; } get tn() { return 'AutoFilter'; } get(n) { const f = this.sh.s.filter; if (n === 'range') { if (!f) throw vbErr(91); return new Rng(this.h, this.sh.b, this.sh.idx(), f.g); } if (n === 'showalldata') { if (this.sh.s.hidden) this.sh.s.hidden.clear(); if (f) f.fields = {}; return EMPTY; } throw vbErr(438); } }
function specialCells(rng, ty, vv) { const wb = rng.wb, si = rng.si, s = rng.sheet; const g = { ...rng.g, r2: rng.limitRows(), c2: rng.limitCols() };
  if (ty === 12) { const rows = range(g.r1, g.r2).filter(r => !(s.hidden && s.hidden.has(r))); if (!rows.length) throw vbErr(1004, 'No cells were found.'); return new Rng(rng.h, rng.book, si, g, rows); }
  if (ty === 11) { const u = wb.usedRange(s) || g; return rng.cell(u.r2, u.c2); }
  const pick = (r, c) => ty === 4 ? wb.isEmpty(si, r, c) : ty === 2 ? !wb.isEmpty(si, r, c) && !wb.formula(si, r, c) : ty === -4123 ? !!wb.formula(si, r, c) : false;
  const rows = []; let c1 = Infinity, c2 = -1; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (pick(r, c)) { if (!rows.includes(r)) rows.push(r); c1 = Math.min(c1, c); c2 = Math.max(c2, c); }
  if (!rows.length) throw vbErr(1004, 'No cells were found.'); return new Rng(rng.h, rng.book, si, { r1: rows[0], c1, r2: rows[rows.length - 1], c2 }, rows); }
function removeDuplicates(rng, cols, header) { const wb = rng.wb, si = rng.si; const g = { ...rng.g, r2: rng.limitRows(), c2: rng.limitCols() }; const cs = cols instanceof VArray ? cols.d.map(x => g.c1 + toNum(x) - 1) : [g.c1 + (has(cols) ? toNum(cols) : 1) - 1]; const seen = new Set(); const drop = [];
  for (let r = header ? g.r1 + 1 : g.r1; r <= g.r2; r++) { const k = cs.map(c => lc(fmtShown(wb.value(si, r, c)))).join('\u0001'); if (seen.has(k)) drop.push(r); else seen.add(k); }
  const s = wb.sheets[si]; for (const r of drop.reverse()) { for (let rr = r; rr < g.r2; rr++) for (let c = g.c1; c <= g.c2; c++) { const src = wb.cellObj(s, rr + 1, c); const k = rr * 16384 + c; if (src) s.cells.set(k, src); else s.cells.delete(k); } for (let c = g.c1; c <= g.c2; c++) s.cells.delete(g.r2 * 16384 + c); g.r2--; } wb.invalidate(); }
class Areas extends VObject { constructor(r) { super(); this.r = r; } get tn() { return 'Areas'; } list() { if (!this.r.rows) return [this.r]; return groupRows(this.r.rows).map(([a, b]) => this.r.sub({ r1: a, c1: this.r.g.c1, r2: b, c2: this.r.g.c2 })); } get(n, a) { if (n === 'count') return this.list().length; if (n === 'item') return this.dflt(a); throw vbErr(438); } dflt(a) { const k = toNum(a[0].v); const l = this.list(); if (k < 1 || k > l.length) throw vbErr(9); return l[k - 1]; } each() { return this.list(); } }

// ---------------------------------------------------------------- Font, Interior, Borders
class Font extends VObject { constructor(r) { super(); this.r = r; } get tn() { return 'Font'; }
  get(n) { const st = this.r.first(); if (n === 'bold') return !!st.bold; if (n === 'italic') return !!st.italic; if (n === 'color') return st.color ? hexToBgr(st.color) : 0; if (n === 'size') return st.size || 11; if (n === 'name') return st.font || 'Calibri'; if (n === 'underline') return st.underline ? 2 : -4142; if (n === 'colorindex') return st.color ? 3 : -4105; throw vbErr(438); }
  put(n, a, v) { if (n === 'bold') this.r.style({ bold: toBool(v) || null }); else if (n === 'italic') this.r.style({ italic: toBool(v) || null }); else if (n === 'color') this.r.style({ color: toNum(v) === 0 ? null : bgrToHex(v) }); else if (n === 'colorindex') { const k = toNum(v); this.r.style({ color: k < 1 ? null : bgrToHex(COLORS56[k] || 0) }); } else if (n === 'size') this.r.style({ size: toNum(v) }); else if (n === 'name') this.r.style({ font: toStr(v) }); else if (n === 'underline') this.r.style({ underline: toNum(v) !== -4142 || null }); else if (n === 'strikethrough' || n === 'themecolor' || n === 'tintandshade') return; else throw vbErr(438); } }
class Interior extends VObject { constructor(r) { super(); this.r = r; } get tn() { return 'Interior'; }
  get(n) { const st = this.r.first(); if (n === 'color') return hexToBgr(st.fill); if (n === 'colorindex') return st.fill ? Math.max(1, COLORS56.indexOf(hexToBgr(st.fill))) : -4142; if (n === 'pattern') return st.fill ? 1 : -4142; throw vbErr(438); }
  put(n, a, v) { if (n === 'color') this.r.style({ fill: bgrToHex(v) }); else if (n === 'colorindex') { const k = toNum(v); this.r.style({ fill: k === -4142 || k === 0 ? null : bgrToHex(COLORS56[k] ?? 16777215) }); } else if (n === 'pattern') { if (toNum(v) === -4142) this.r.style({ fill: null }); } else if (['themecolor', 'tintandshade', 'patterncolorindex', 'patterncolor', 'patternthemecolor', 'patterntintandshade'].includes(n)) return; else throw vbErr(438); } }
class Borders extends VObject { constructor(r) { super(); this.r = r; } get tn() { return 'Borders'; } get(n, a) { if (n === 'linestyle') return this.r.first().border ? 1 : -4142; if (n === 'weight') return 2; if (n === 'item') return this; throw vbErr(438); } put(n, a, v) { if (n === 'linestyle') this.r.style({ border: toNum(v) !== -4142 || null }); else if (n === 'weight' || n === 'color' || n === 'colorindex' || n === 'themecolor' || n === 'tintandshade') { if (n === 'weight') this.r.style({ border: true }); } else throw vbErr(438); } dflt() { return this; } }

// ---------------------------------------------------------------- tables (ListObjects)
class Tables extends VObject { constructor(h, sh) { super(); this.h = h; this.sh = sh; } get tn() { return 'ListObjects'; }
  list() { return (this.sh.s.tables || []).map(t => new Table(this.h, this.sh, t)); }
  get(n, a) { if (n === 'count') return this.list().length; if (n === 'item') return this.dflt(a);
    if (n === 'add') { const [st, src, , hdr] = bind(a, ['sourcetype', 'source', 'linksource', 'xllistobjecthasheaders', 'destination', 'tablestylename']); const g = src instanceof Rng ? { ...src.g, r2: src.limitRows(), c2: src.limitCols() } : currentRegion(this.sh.b.wb, this.sh.idx(), this.h.sel.ar, this.h.sel.ac); const nm = this.sh.b.wb.addTable(this.sh.idx(), colName(g.c1) + (g.r1 + 1) + ':' + colName(g.c2) + (g.r2 + 1)); return new Table(this.h, this.sh, this.sh.s.tables.find(t => t.name === nm)); }
    throw vbErr(438); }
  dflt(a) { const k = deref(a[0].v); const l = this.list(); if (typeof k === 'number') { if (!l[k - 1]) throw vbErr(9); return l[k - 1]; } const t = l.find(x => lc(x.t.name) === lc(toStr(k))); if (!t) throw vbErr(9); return t; } each() { return this.list(); } }
class Table extends VObject { constructor(h, sh, t) { super(); this.h = h; this.sh = sh; this.t = t; } get tn() { return 'ListObject'; }
  get g() { return parseRange(this.t.ref); } rng(g) { return new Rng(this.h, this.sh.b, this.sh.idx(), g); }
  body() { const g = this.g; const end = this.t.totals ? g.r2 - 1 : g.r2; return end > g.r1 ? { ...g, r1: g.r1 + 1, r2: end } : null; }
  get(n, a) { const g = this.g;
    switch (n) { case 'name': return this.t.name; case 'range': return this.rng(g); case 'headerrowrange': return this.rng({ ...g, r2: g.r1 }); case 'databodyrange': { const b = this.body(); return b ? this.rng(b) : NOTHING; } case 'totalsrowrange': return this.t.totals ? this.rng({ ...g, r1: g.r2 }) : NOTHING; case 'showtotals': return !!this.t.totals;
      case 'listrows': return new ListRows(this); case 'listcolumns': { const c = new ListCols(this); return a.length ? c.dflt(a) : c; } case 'parent': return this.sh; case 'resize': { const r = a[0].v; this.t.ref = colName(r.g.c1) + (r.g.r1 + 1) + ':' + colName(r.g.c2) + (r.g.r2 + 1); this.sh.b.wb.invalidate(); return EMPTY; } case 'unlist': this.sh.s.tables = this.sh.s.tables.filter(x => x !== this.t); return EMPTY; case 'delete': return EMPTY; case 'tablestyle': return 'TableStyleMedium2'; case 'sort': return new SortObj(this.h, this.sh); case 'autofilter': return new AutoFilterObj(this.h, this.sh); }
    throw vbErr(438); }
  put(n, a, v) { if (n === 'name') { this.t.name = toStr(v); this.sh.b.wb.invalidate(); return; } if (n === 'showtotals' || n === 'tablestyle') return; throw vbErr(438); } }
class ListRows extends VObject { constructor(tb) { super(); this.tb = tb; } get tn() { return 'ListRows'; }
  get(n, a) { const tb = this.tb; const b = tb.body(); if (n === 'count') return b ? b.r2 - b.r1 + 1 : 0;
    if (n === 'add') { const g = tb.g; const wb = tb.sh.b.wb; const si = tb.sh.idx(); const [pos] = bind(a, ['position', 'alwaysinsert']); const empty1 = b && b.r1 === b.r2 && [...Array(g.c2 - g.c1 + 1).keys()].every(k => wb.isEmpty(si, b.r1, g.c1 + k)); if (empty1 && !has(pos)) return new ListRow(tb, b.r1);
      let r = has(pos) ? g.r1 + toNum(pos) : g.r2 + 1; if (has(pos)) insertCells(new Rng(tb.h, tb.sh.b, si, { r1: r, c1: g.c1, r2: r, c2: g.c2 }), -4121); tb.t.ref = colName(g.c1) + (g.r1 + 1) + ':' + colName(g.c2) + (g.r2 + 2); grow(wb, si, { r2: g.r2 + 1, c2: g.c2 });
      for (let c = g.c1; c <= g.c2; c++) { const above = wb.cellObj(wb.sheets[si], r - 1, c); if (above && above.f && r - 1 > g.r1) { wb.setRaw(wb.sheets[si], r, c, shiftFormula(above.f, 1, 0), true); } if (above && above.s && r - 1 > g.r1) wb.setStyle(si, r, c, { ...above.s }); } wb.invalidate(); return new ListRow(tb, r); }
    if (n === 'item') return this.dflt(a); throw vbErr(438); }
  dflt(a) { const b = this.tb.body(); const k = toNum(a[0].v); if (!b || k < 1 || k > b.r2 - b.r1 + 1) throw vbErr(9); return new ListRow(this.tb, b.r1 + k - 1); } each() { const b = this.tb.body(); return b ? range(b.r1, b.r2).map(r => new ListRow(this.tb, r)) : []; } }
class ListRow extends VObject { constructor(tb, r) { super(); this.tb = tb; this.r = r; } get tn() { return 'ListRow'; } get(n) { const g = this.tb.g; if (n === 'range') return this.tb.rng({ r1: this.r, c1: g.c1, r2: this.r, c2: g.c2 }); if (n === 'index') return this.r - g.r1; if (n === 'delete') { deleteCells(this.tb.rng({ r1: this.r, c1: g.c1, r2: this.r, c2: g.c2 }), -4162); this.tb.t.ref = colName(g.c1) + (g.r1 + 1) + ':' + colName(g.c2) + g.r2; return EMPTY; } throw vbErr(438); } }
class ListCols extends VObject { constructor(tb) { super(); this.tb = tb; } get tn() { return 'ListColumns'; } heads() { const g = this.tb.g; const wb = this.tb.sh.b.wb; const si = this.tb.sh.idx(); return range(g.c1, g.c2).map(c => toStr(fromCell(wb.value(si, g.r1, c)))); }
  get(n, a) { if (n === 'count') return this.heads().length; if (n === 'item') return this.dflt(a); if (n === 'add') { const g = this.tb.g; this.tb.t.ref = colName(g.c1) + (g.r1 + 1) + ':' + colName(g.c2 + 1) + (g.r2 + 1); const wb = this.tb.sh.b.wb; wb.setRaw(wb.sheets[this.tb.sh.idx()], g.r1, g.c2 + 1, 'Column' + (g.c2 - g.c1 + 2), true); wb.invalidate(); return new ListCol(this.tb, g.c2 + 1); } throw vbErr(438); }
  dflt(a) { const k = deref(a[0].v); const hs = this.heads(); let i; if (typeof k === 'number') i = k - 1; else i = hs.findIndex(x => lc(x) === lc(toStr(k))); if (i < 0 || i >= hs.length) throw vbErr(9); return new ListCol(this.tb, this.tb.g.c1 + i); } each() { return this.heads().map((_, i) => new ListCol(this.tb, this.tb.g.c1 + i)); } }
class ListCol extends VObject { constructor(tb, c) { super(); this.tb = tb; this.c = c; } get tn() { return 'ListColumn'; } get(n) { const g = this.tb.g; const b = this.tb.body(); if (n === 'name') { const wb = this.tb.sh.b.wb; return toStr(fromCell(wb.value(this.tb.sh.idx(), g.r1, this.c))); } if (n === 'index') return this.c - g.c1 + 1; if (n === 'range') return this.tb.rng({ r1: g.r1, c1: this.c, r2: g.r2, c2: this.c }); if (n === 'databodyrange') return b ? this.tb.rng({ r1: b.r1, c1: this.c, r2: b.r2, c2: this.c }) : NOTHING; throw vbErr(438); } put(n, a, v) { if (n === 'name') { const wb = this.tb.sh.b.wb; wb.setRaw(wb.sheets[this.tb.sh.idx()], this.tb.g.r1, this.c, toStr(v)); return; } throw vbErr(438); } }

// ---------------------------------------------------------------- WorksheetFunction: any of the engine's functions, by building a formula
// ---------------------------------------------------------------- PageSetup: kept on the sheet, used by ExportAsFixedFormat
const PAGE_DEFAULTS = { orientation: 1, zoom: 100, fittopageswide: 1, fittopagestall: 1, printtitlerows: '', printtitlecolumns: '', printarea: '', centerfooter: '', leftfooter: '', rightfooter: '', centerheader: '', leftheader: '', rightheader: '', papersize: 9, centerhorizontally: false, printgridlines: false };
class PageSetup extends VObject { constructor(s) { super(); this.s = s; } get tn() { return 'PageSetup'; }
  get(n) { const p = this.s.page || {}; if (n in p) return p[n]; if (n in PAGE_DEFAULTS) return PAGE_DEFAULTS[n]; if (/margin$/.test(n)) return 54; throw vbErr(438); }
  put(n, a, v) { if (!(n in PAGE_DEFAULTS) && !/margin$/.test(n) && !['draft', 'blackandwhite', 'order', 'firstpagenumber', 'printcomments', 'centervertically', 'printheadings', 'printquality'].includes(n)) throw vbErr(438);
    const p = this.s.page || (this.s.page = {}); let x = deref(v); if (n === 'zoom' && x === false) x = false; else if (typeof x === 'boolean' || typeof x === 'number') x = x; else x = toStr(x); p[n] = x; } }
// ---------------------------------------------------------------- Outlook, as CreateObject("Outlook.Application") sees it: emails go to the pretend Outbox
class OlApp extends VObject { constructor(h) { super(); this.h = h; } get tn() { return 'Application'; }
  get(n, a) { if (n === 'createitem') { const k = toNum(a[0] ? deref(a[0].v) : 0); if (k !== 0) throw vbErr(13); return new MailItem(this.h); } if (n === 'name') return 'Outlook'; if (n === 'session' || n === 'getnamespace') return this; if (n === 'quit') return EMPTY; throw vbErr(438); } }
class MailItem extends VObject { constructor(h) { super(); this.h = h; this.p = { to: '', cc: '', bcc: '', subject: '', body: '', htmlbody: '' }; this.att = []; } get tn() { return 'MailItem'; }
  snap(how) { return { to: this.p.to, cc: this.p.cc, bcc: this.p.bcc, subject: this.p.subject, body: this.p.body || this.p.htmlbody.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ''), html: !!this.p.htmlbody, attachments: this.att.map(x => x.split('\\').pop()), paths: this.att.slice(), how }; }
  record(how) { if (this.sent) throw vbErr(-2147221238, 'The item has been moved or deleted.'); const o = this.h.outbox; const e = this.snap(how); if (this.entry !== undefined) o[this.entry] = e; else { this.entry = o.length; o.push(e); } this.h.log.push({ kind: 'mail', how, subject: e.subject }); return e; }
  get(n, a) { if (n in this.p) return this.p[n]; if (n === 'attachments') return new MailAtt(this); if (n === 'send') { if (!String(this.p.to).trim() && !String(this.p.cc).trim()) throw vbErr(-2147467259, 'Outlook does not recognize one or more names.'); this.record('send'); this.sent = true; return EMPTY; }
    if (n === 'display') { const e = this.record('display'); return { ui: 'mail', mail: e, after: () => EMPTY }; } if (n === 'save') { this.record('draft'); return EMPTY; } if (n === 'close' || n === 'importance' || n === 'sentonbehalfofname') return n === 'importance' ? 1 : EMPTY; throw vbErr(438); }
  put(n, a, v) { if (n in this.p) { this.p[n] = toStr(v); return; } if (n === 'importance' || n === 'sentonbehalfofname' || n === 'bodyformat' || n === 'readreceiptrequested') return; throw vbErr(438); } }
class MailAtt extends VObject { constructor(m) { super(); this.m = m; } get tn() { return 'Attachments'; }
  get(n, a) { if (n === 'count') return this.m.att.length; if (n === 'add') { const p = toStr(deref(a[0].v)); const f = this.m.h.file(p); if (!f || f.kind === 'folder') throw vbErr(-2147024894, 'Cannot find this file. Verify the path and file name are correct.'); this.m.att.push(f.path); return EMPTY; } throw vbErr(438); } }
export class WsFn extends VObject {
  constructor(h, strict) { super(); this.h = h; this.strict = strict; } get tn() { return 'WorksheetFunction'; }
  static has(n) { return !!FN[n.toUpperCase()]; }
  get(n, a, rt) { const name = n.toUpperCase(); if (!FN[name]) throw vbErr(438); const h = this.h; let si = h.active.activeSi; let book = h.active;
    const lit = v => { v = v instanceof Rng ? v : deref(v); if (v === MISSING) return ''; if (v instanceof Rng) { book = v.book; return `'${v.sheet.name.replace(/'/g, "''")}'!${v.address(false, false)}`; } if (v instanceof VArray) { const two = v.b.length === 2; if (!two) return '{' + v.d.map(lit).join(',') + '}'; const rows = []; const nc = v.b[1][1] - v.b[1][0] + 1; for (let i = 0; i < v.d.length; i += nc) rows.push(v.d.slice(i, i + nc).map(lit).join(',')); return '{' + rows.join(';') + '}'; } if (typeof v === 'string') return '"' + v.replace(/"/g, '""') + '"'; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (v instanceof VDate) return String(v.n); if (v === EMPTY) return '0'; return String(toNum(v)); };
    const f = `=${name}(${a.map(x => lit(x.v)).join(',')})`; const sheetOf = a.find(x => x.v instanceof Rng); if (sheetOf) si = sheetOf.v.si; h.tick('read', 1, rt);
    let v; try { v = book.wb.evaluate(f, si); } catch (e) { v = { code: '#VALUE!' }; }
    if (v && v.range) { if (v.single !== undefined) v = v.single; else { const rows = v.vals; const out = new VArray([[1, rows.length], [1, rows[0].length]]); let k = 0; for (const r of rows) for (const x of r) out.d[k++] = fromCell(x); return out; } }
    if (isErr(v)) { if (this.strict) throw vbErr(1004, `Unable to get the ${n[0].toUpperCase() + n.slice(1)} property of the WorksheetFunction class`); return v; }
    return v === null || v === undefined ? 0 : v; }
}

// ---------------------------------------------------------------- the pretend folder: Dir, Kill, FileCopy, MkDir…
const FS_FNS = {
  dir(h, [pat, attr]) { if (pat === undefined || pat === MISSING) { const st = h.dirState; if (!st) throw vbErr(5); return st.list.length > st.i ? st.list[st.i++] : ''; }
    const p = toStr(pat); const isDir = has(attr) && (toNum(attr) & 16); const folder = /[*?]/.test(tail(p)) || /\.[a-z0-9]+$/i.test(tail(p)) || !isDir ? dirOf(p) : p; const name = /[*?]/.test(tail(p)) || /\.[a-z0-9]+$/i.test(tail(p)) ? tail(p) : (isDir ? '' : tail(p) || '*');
    if (isDir && !/[*?]/.test(p)) { const k = lc(p).replace(/\\$/, ''); if (h.folderExists(k)) { h.dirState = { list: [], i: 0 }; return tail(p.replace(/\\$/, '')); } const f = h.file(p); h.dirState = { list: [], i: 0 }; return f ? tail(p) : ''; }
    const re = wild(name || '*'); const list = [...h.fs.values()].filter(f => f.kind !== 'folder' && lc(dirOf(f.path)) === lc(folder.replace(/\\$/, '')) && re.test(tail(f.path))).map(f => tail(f.path)).sort((x, y) => x.localeCompare(y));
    h.dirState = { list, i: 0 }; return list.length ? list[h.dirState.i++] : ''; },
  kill(h, [p]) { const k = toStr(p); const re = wild(tail(k)); const d = lc(dirOf(k)); const hit = [...h.fs.keys()].filter(x => lc(dirOf(h.fs.get(x).path)) === d && re.test(tail(h.fs.get(x).path)) && h.fs.get(x).kind !== 'folder'); if (!hit.length) throw vbErr(53); for (const x of hit) h.fs.delete(x); h.log.push({ kind: 'kill', path: k }); return EMPTY; },
  filecopy(h, [a, b]) { const f = h.file(toStr(a)); if (!f) throw vbErr(53); h.writeFile(toStr(b), { ...f, path: toStr(b) }); return EMPTY; },
  mkdir(h, [p]) { const k = toStr(p).replace(/\\$/, ''); if (h.folderExists(k) && h.fs.has(lc(k))) throw vbErr(75); if (!h.folderExists(dirOf(k))) throw vbErr(76); h.fs.set(lc(k), { path: k, kind: 'folder', made: true }); h.log.push({ kind: 'mkdir', path: k }); return EMPTY; },
  rmdir(h, [p]) { h.fs.delete(lc(toStr(p))); return EMPTY; },
  filelen(h, [p]) { const f = h.file(toStr(p)); if (!f) throw vbErr(53); return (f.text || JSON.stringify(f.spec || f.pages || '')).length; },
  filedatetime(h, [p]) { if (!h.file(toStr(p))) throw vbErr(53); return new VDate(h.now()); },
  curdir(h) { return h.path; }, chdir() { return EMPTY; }, getattr(h, [p]) { const f = h.file(toStr(p)); if (f) return f.kind === 'folder' ? 16 : 32; if (h.folderExists(toStr(p))) return 16; throw vbErr(53); },
};
