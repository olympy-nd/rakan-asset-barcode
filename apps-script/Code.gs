/**
 * منظومة باركود الأصول — راكان لاستيراد المواد الغذائية والتوزيع
 * خادم Google Sheets (Apps Script) — الإصدار 2: مزامنة على مستوى الأصل الواحد.
 * يُلصق داخل الشيت: الإضافات ← Apps Script
 * ثم: نشر ← نشر جديد ← تطبيق ويب ← التنفيذ باسم: أنا ← من يمكنه الوصول: أي شخص
 * رمز الدخول يُحدَّد من المنظومة عند أول دخول ويُقفل هنا (لا يُكتب في الكود).
 * لتغيير رمز الدخول: شغّل الدالة resetKey من المحرر ثم أدخل الرمز الجديد في المنظومة.
 */
const A_SHEET = 'الأصول', CODE_H = 'رقم الباركود', VER_H = 'الإصدار', DEL_SHEET = '_محذوفات';
const SECTIONS = { odoo: 'أصول أودو', settings: 'الإعدادات', cats: 'الفئات', log: 'سجل الطباعة' };
const LEGACY = ['الأصول', 'أصول أودو', 'الفئات', 'سجل الطباعة', 'الإعدادات'];
const TEXT_COL = /رقم|باركود|الكود|الرمز|تاريخ|التاريخ|آخر طباعة|المقاس/;

function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  return handle_(p);
}

/* ---------- auth ---------- */
function handle_(p) {
  const props = PropertiesService.getScriptProperties(), stored = props.getProperty('key');
  if (p.action === 'init') {
    if (!p.key || String(p.key).length < 24) return out_({ ok: false, error: 'weak_key' });
    if (stored && stored !== p.key) return out_({ ok: false, error: 'unauthorized' });
    if (!stored) props.setProperty('key', p.key);
  } else {
    if (!stored) return out_({ ok: false, error: 'not_initialized' });
    if (!p.key || p.key !== stored) return out_({ ok: false, error: 'unauthorized' });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const a = p.action;
  if (a === 'init' || a === 'ping') return out_({ ok: true, name: ss.getName(), rev: rev_(), seq: seq_(), v: 2 });
  if (a === 'sync') { if (!seq_()) nextSeq_(); return out_(sync_(ss, +p.since || 0)); }
  if (a === 'load') return out_(legacyLoad_(ss));
  // ---- writes ----
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ ok: false, error: 'busy' });
  try {
    if (a === 'upsert') return out_(upsert_(ss, p));
    if (a === 'upsertMany') return out_(upsertMany_(ss, p));
    if (a === 'remove') return out_(remove_(ss, p));
    if (a === 'putSection') return out_(putSection_(ss, p));
    if (a === 'appendLog') return out_(appendLog_(ss, p));
    if (a === 'save') return out_(legacySave_(ss, p));
  } finally { lock.releaseLock(); }
  return out_({ ok: false, error: 'unknown_action' });
}

/* ---------- sequence / versions ---------- */
function P_() { return PropertiesService.getScriptProperties(); }
function seq_() { return +(P_().getProperty('seq') || 0); }
function nextSeq_() { const n = seq_() + 1; P_().setProperty('seq', String(n)); P_().setProperty('rev', 's' + n); return n; }
function secSeq_(k) { return +(P_().getProperty('sec_' + k) || 0); }
function rev_() { return P_().getProperty('rev') || ('s' + seq_()); }
/** تعديل يدوي داخل الشيت: يجبر الأجهزة على إعادة تحميل كاملة */
function onEdit() { const n = nextSeq_(); P_().setProperty('sec_full', String(n)); }
function resetKey() { P_().deleteProperty('key'); }

/* ---------- sheet helpers ---------- */
function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (head && head.length) { sh.getRange(1, 1, 1, head.length).setValues([head]); fmtHead_(sh, head.length); textCols_(sh, head); }
    if (name === DEL_SHEET) sh.hideSheet();
  }
  return sh;
}
function fmtHead_(sh, w) {
  sh.getRange(1, 1, 1, w).setBackground('#0C4CA2').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1); sh.setRightToLeft(true);
}
function textCols_(sh, head) {
  head.forEach((h, c) => { if (TEXT_COL.test(String(h))) sh.getRange(1, c + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
}
function cell_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v; }
function read_(sh) {
  if (!sh || sh.getLastRow() < 1) return [];
  return sh.getDataRange().getValues().map(r => r.map(cell_)).filter((r, i) => i === 0 || r.some(v => v !== '' && v !== null));
}
function header_(sh) { const w = sh.getLastColumn(); return w ? sh.getRange(1, 1, 1, w).getValues()[0].map(String) : []; }
/** يضمن وجود الأعمدة المطلوبة ويعيد خريطة العنوان ← رقم العمود */
function ensureCols_(sh, names) {
  let head = header_(sh); const add = names.filter(n => head.indexOf(n) < 0);
  if (add.length) {
    const start = head.length + 1;
    if (sh.getMaxColumns() < head.length + add.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length + add.length - sh.getMaxColumns());
    sh.getRange(1, start, 1, add.length).setValues([add]); fmtHead_(sh, head.length + add.length); head = head.concat(add);
    add.forEach((n, i) => { if (TEXT_COL.test(n)) sh.getRange(1, start + i, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  const m = {}; head.forEach((h, i) => m[h] = i); return { head, m };
}
function findRow_(sh, codeCol, code) {
  const n = sh.getLastRow(); if (n < 2) return -1;
  const vals = sh.getRange(2, codeCol + 1, n - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === code) return i + 2;
  return -1;
}
function codes_(sh, codeCol) {
  const n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, codeCol + 1, n - 1, 1).getValues().map(r => String(r[0]).trim());
}
/** رقم باركود بديل غير مكرر بنفس النمط */
function nextFreeCode_(code, used) {
  const set = new Set(used);
  if (/^29\d{11}$/.test(code)) {
    let max = 0; used.forEach(c => { if (/^29\d{11}$/.test(c)) max = Math.max(max, +c.slice(2, 12)); });
    for (let n = max + 1; ; n++) { const d12 = '29' + String(n).padStart(10, '0'); let s = 0; for (let i = 0; i < 12; i++) s += (+d12[i]) * (i % 2 ? 3 : 1); const c = d12 + ((10 - s % 10) % 10); if (!set.has(c)) return c; }
  }
  const m = code.match(/^(.*?)(\d+)$/); if (!m) { let i = 2; while (set.has(code + '-' + i)) i++; return code + '-' + i; }
  const head = m[1], w = m[2].length; let max = 0;
  used.forEach(c => { if (c.indexOf(head) === 0) { const t = c.slice(head.length); if (/^\d+$/.test(t)) max = Math.max(max, +t); } });
  for (let n = max + 1; ; n++) { const c = head + String(n).padStart(w, '0'); if (!set.has(c)) return c; }
}
function rowObj_(head, row) { const o = {}; head.forEach((h, i) => o[h] = cell_(row[i])); return o; }

/* ---------- sync ---------- */
function sync_(ss, since) {
  const seq = seq_();
  if (since && since === seq) return { ok: true, seq, rev: rev_(), same: true };
  const full = !since || since < secSeq_('full');
  const res = { ok: true, seq, rev: rev_(), full, name: ss.getName(), sections: {} };
  const sh = ss.getSheetByName(A_SHEET);
  const rows = read_(sh);
  res.head = rows[0] || [];
  const vi = res.head.indexOf(VER_H);
  res.assets = full ? rows.slice(1) : rows.slice(1).filter(r => vi >= 0 && (+r[vi] || 0) > since);
  if (!full) {
    const dsh = ss.getSheetByName(DEL_SHEET);
    res.deleted = dsh ? read_(dsh).slice(1).filter(r => +r[1] > since).map(r => String(r[0])) : [];
  } else res.deleted = [];
  Object.keys(SECTIONS).forEach(k => { if (full || secSeq_(k) > since) res.sections[k] = read_(ss.getSheetByName(SECTIONS[k])); });
  return res;
}

/* ---------- writes ---------- */
function upsert_(ss, p) {
  const obj = p.row || {}, isNew = !!p.isNew, oldCode = p.oldCode ? String(p.oldCode).trim() : '';
  let code = String(obj[CODE_H] || '').trim(); if (!code) return { ok: false, error: 'no_code' };
  const sh = sheet_(ss, A_SHEET, [CODE_H]);
  const { head, m } = ensureCols_(sh, Object.keys(obj).concat([VER_H]));
  const ci = m[CODE_H], vi = m[VER_H];
  const used = codes_(sh, ci);
  let target = -1, renamed = '';
  if (isNew) {
    if (used.indexOf(code) >= 0) { renamed = code; code = nextFreeCode_(code, used); }
  } else {
    const look = oldCode || code;
    target = findRow_(sh, ci, look);
    if (target < 0) return { ok: false, error: 'gone', code: look };
    const cur = sh.getRange(target, 1, 1, head.length).getValues()[0];
    if ((+cur[vi] || 0) !== (+p.baseVer || 0)) return { ok: false, error: 'conflict', code: look, row: rowObj_(head, cur), head };
    if (oldCode && code !== oldCode && used.indexOf(code) >= 0) return { ok: false, error: 'code_taken', code };
  }
  const n = nextSeq_();
  obj[CODE_H] = code; obj[VER_H] = n;
  const row = head.map(h => (h in obj) ? obj[h] : '');
  if (target < 0) { sh.appendRow(row); }
  else {
    const cur = sh.getRange(target, 1, 1, head.length).getValues()[0];
    head.forEach((h, i) => { if (!(h in obj)) row[i] = cur[i]; });
    sh.getRange(target, 1, 1, head.length).setValues([row]);
  }
  if (oldCode && code !== oldCode) logDel_(ss, oldCode, n);
  return { ok: true, seq: n, ver: n, code, renamed, row: rowObj_(head, row) };
}
/** دفعة واحدة: أصول جديدة و/أو تعديلات (مثل التسجيل الجماعي أو تحديث عدّاد الطباعة) */
function upsertMany_(ss, p) {
  const items = p.items || []; if (!items.length) return { ok: true, seq: seq_(), results: [] };
  const sh = sheet_(ss, A_SHEET, [CODE_H]);
  const keys = {}; items.forEach(it => Object.keys(it.row || {}).forEach(k => keys[k] = 1));
  const { head, m } = ensureCols_(sh, Object.keys(keys).concat([VER_H]));
  const ci = m[CODE_H], vi = m[VER_H], w = head.length;
  const last = sh.getLastRow(), data = last > 1 ? sh.getRange(2, 1, last - 1, w).getValues() : [];
  const idx = {}; data.forEach((r, i) => idx[String(r[ci]).trim()] = i);
  const used = Object.keys(idx);
  const n = nextSeq_(), results = [], adds = []; let touched = false;
  items.forEach(it => {
    const obj = Object.assign({}, it.row || {}); let code = String(obj[CODE_H] || '').trim();
    if (!code) return results.push({ code, error: 'no_code' });
    if (it.isNew) {
      let nc = code; if (idx[nc] !== undefined || used.indexOf(nc) >= 0) nc = nextFreeCode_(code, used);
      used.push(nc); obj[CODE_H] = nc; obj[VER_H] = n;
      adds.push(head.map(h => (h in obj) ? obj[h] : '')); idx[nc] = -1;
      results.push({ code, newCode: nc, ver: n });
    } else {
      const i = idx[code]; if (i === undefined || i < 0) return results.push({ code, error: 'gone' });
      const cur = data[i];
      if ((+cur[vi] || 0) !== (+it.baseVer || 0)) return results.push({ code, error: 'conflict', row: rowObj_(head, cur) });
      obj[VER_H] = n; data[i] = head.map((h, j) => (h in obj) ? obj[h] : cur[j]); touched = true;
      results.push({ code, newCode: code, ver: n });
    }
  });
  if (touched) sh.getRange(2, 1, data.length, w).setValues(data);
  if (adds.length) sh.getRange(sh.getLastRow() + 1, 1, adds.length, w).setValues(adds);
  return { ok: true, seq: n, results, head };
}
function remove_(ss, p) {
  const code = String(p.code || '').trim();
  const sh = ss.getSheetByName(A_SHEET); if (!sh) return { ok: true, seq: seq_() };
  const { head, m } = ensureCols_(sh, [VER_H]);
  const r = findRow_(sh, m[CODE_H], code);
  if (r < 0) return { ok: true, seq: seq_(), gone: true };
  const cur = sh.getRange(r, 1, 1, head.length).getValues()[0];
  if (!p.force && (+cur[m[VER_H]] || 0) !== (+p.baseVer || 0)) return { ok: false, error: 'conflict', code, row: rowObj_(head, cur), head };
  sh.deleteRow(r);
  const n = nextSeq_(); logDel_(ss, code, n);
  return { ok: true, seq: n };
}
function logDel_(ss, code, n) {
  const d = sheet_(ss, DEL_SHEET, ['رقم الباركود', 'التسلسل']);
  d.appendRow([code, n]);
  const last = d.getLastRow(); if (last > 1500) d.deleteRows(2, last - 1001);
}
function putSection_(ss, p) {
  const name = SECTIONS[p.section]; if (!name) return { ok: false, error: 'bad_section' };
  writeAll_(ss, name, p.rows || []);
  const n = nextSeq_(); P_().setProperty('sec_' + p.section, String(n));
  return { ok: true, seq: n };
}
function appendLog_(ss, p) {
  const rows = p.rows || []; if (!rows.length) return { ok: true, seq: seq_() };
  const head = p.head || ['التاريخ', 'رقم الباركود', 'اسم الأصل', 'المقاس', 'عدد النسخ'];
  const sh = sheet_(ss, SECTIONS.log, head);
  const w = Math.max(head.length, sh.getLastColumn());
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows.map(r => { const x = r.slice(0, head.length); while (x.length < head.length) x.push(''); return x; }));
  const n = nextSeq_(); P_().setProperty('sec_log', String(n));
  return { ok: true, seq: n, w };
}
function writeAll_(ss, name, rows) {
  let sh = ss.getSheetByName(name), fresh = false;
  if (!sh) { sh = ss.insertSheet(name); fresh = true; }
  sh.clearContents();
  if (!rows || !rows.length) return;
  const w = Math.max.apply(null, rows.map(r => r.length));
  const grid = rows.map(r => { const x = r.slice(); while (x.length < w) x.push(''); return x; });
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  if (sh.getMaxRows() < grid.length + 1) sh.insertRowsAfter(sh.getMaxRows(), grid.length + 1 - sh.getMaxRows());
  textCols_(sh, grid[0]);
  sh.getRange(1, 1, grid.length, w).setValues(grid);
  if (fresh || sh.getFrozenRows() !== 1) fmtHead_(sh, w);
}

/* ---------- legacy full load/save (للأجهزة التي لم تُحدَّث بعد) ---------- */
function legacyLoad_(ss) {
  const data = {}; LEGACY.forEach(n => { data[n] = read_(ss.getSheetByName(n)); });
  return { ok: true, name: ss.getName(), rev: rev_(), seq: seq_(), data };
}
function legacySave_(ss, p) {
  if (p.baseRev && p.baseRev !== rev_()) return { ok: false, error: 'conflict', rev: rev_() };
  const n0 = seq_() + 1;
  Object.keys(p.data || {}).forEach(name => {
    if (LEGACY.indexOf(name) < 0) return;
    let rows = p.data[name];
    if (name === A_SHEET && rows && rows.length) {
      let vi = rows[0].indexOf(VER_H);
      if (vi < 0) { rows = rows.map((r, i) => r.concat([i ? n0 : VER_H])); }
      else rows = rows.map((r, i) => { if (i) { r = r.slice(); r[vi] = n0; } return r; });
    }
    writeAll_(ss, name, rows);
  });
  const n = nextSeq_(); P_().setProperty('sec_full', String(n));
  return { ok: true, rev: rev_(), seq: n };
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
