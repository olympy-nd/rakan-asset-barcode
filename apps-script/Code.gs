/**
 * منظومة باركود الأصول — راكان لاستيراد المواد الغذائية والتوزيع
 * خادم Google Sheets (Apps Script). يُلصق داخل الشيت: الإضافات ← Apps Script
 * ثم: نشر ← نشر جديد ← تطبيق ويب ← التنفيذ باسم: أنا ← من يمكنه الوصول: أي شخص
 * المفتاح السري يُنشأ في المنظومة ويُقفل هنا عند أول اتصال (لا يُكتب في الكود).
 * لإعادة ضبط المفتاح: شغّل الدالة resetKey من محرر Apps Script ثم أعد الاتصال من المنظومة.
 */
const SHEETS = ['الأصول', 'أصول أودو', 'الفئات', 'سجل الطباعة', 'الإعدادات'];
const TEXT_COL = /رقم|باركود|الكود|الرمز|تاريخ|التاريخ|آخر طباعة|المقاس/;

function doGet(e) { return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  return handle_(p);
}

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
  if (p.action === 'init') return out_({ ok: true, name: ss.getName(), rev: rev_() });
  if (p.action === 'ping') return out_({ ok: true, name: ss.getName(), rev: rev_() });
  if (p.action === 'load') {
    const data = {};
    SHEETS.forEach(n => { const sh = ss.getSheetByName(n); data[n] = sh ? read_(sh) : []; });
    return out_({ ok: true, name: ss.getName(), rev: rev_(), data: data });
  }
  if (p.action === 'save') {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return out_({ ok: false, error: 'busy' });
    try {
      if (p.baseRev && p.baseRev !== rev_()) return out_({ ok: false, error: 'conflict', rev: rev_() });
      Object.keys(p.data || {}).forEach(n => { if (SHEETS.indexOf(n) >= 0) write_(ss, n, p.data[n]); });
      const r = bump_();
      return out_({ ok: true, rev: r });
    } finally { lock.releaseLock(); }
  }
  return out_({ ok: false, error: 'unknown_action' });
}

function read_(sh) {
  const rng = sh.getDataRange(), vals = rng.getValues(), tz = Session.getScriptTimeZone();
  return vals.map(r => r.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v))
             .filter((r, i) => i === 0 || r.some(v => v !== '' && v !== null));
}

function write_(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  if (!rows || !rows.length) return;
  const w = Math.max.apply(null, rows.map(r => r.length));
  const grid = rows.map(r => { const x = r.slice(); while (x.length < w) x.push(''); return x; });
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  if (sh.getMaxRows() < grid.length + 1) sh.insertRowsAfter(sh.getMaxRows(), grid.length + 1 - sh.getMaxRows());
  grid[0].forEach((h, c) => { if (TEXT_COL.test(String(h))) sh.getRange(1, c + 1, grid.length, 1).setNumberFormat('@'); });
  sh.getRange(1, 1, grid.length, w).setValues(grid);
  sh.getRange(1, 1, 1, w).setBackground('#0C4CA2').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
}

/** يمسح المفتاح المحفوظ؛ أول جهاز يتصل بعدها يحدد المفتاح الجديد */
function resetKey() { PropertiesService.getScriptProperties().deleteProperty('key'); }

function rev_() { return PropertiesService.getScriptProperties().getProperty('rev') || bump_(); }
function bump_() { const r = new Date().toISOString(); PropertiesService.getScriptProperties().setProperty('rev', r); return r; }
/** يُسجّل أي تعديل يدوي في الشيت حتى تعرف المنظومة أن البيانات تغيّرت */
function onEdit() { bump_(); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
