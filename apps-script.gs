// 書沛 IB 衝刺表：進度＋每週測驗 PDF（Google Apps Script）
// 貼到 Google 試算表的「擴充功能 → Apps Script」，部署為網頁應用程式（執行身分：我；存取權：所有人）。
// 第一次使用：在下方 PARENT_PIN 改成你自己的家長密碼（上傳 PDF、填分數時要輸入）。
const PARENT_PIN = '請改成家長密碼';
const FOLDER_NAME = '書沛 IB 週測';
const SHEET = 'progress';
const COLS = ['date', 'math', 'econ', 'chem', 'phys', 'eng', 'note', 'updated'];
const TSHEET = 'tests';
const TCOLS = ['fileId', 'name', 'subject', 'week', 'uploaded', 'score', 'comment'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, cols) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function checkPin_(pin) {
  if (PARENT_PIN === '請改成家長密碼') throw new Error('尚未設定家長密碼（apps-script.gs 的 PARENT_PIN）');
  if (String(pin || '') !== PARENT_PIN) throw new Error('家長密碼錯誤');
}
function ymd_(d) { return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd'); }

// ---------- 進度 ----------
function getProgress_() {
  const rows = sheet_(SHEET, COLS).getDataRange().getValues().slice(1);
  const data = {};
  rows.forEach(r => {
    const d = r[0] instanceof Date ? ymd_(r[0]) : String(r[0]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    data[d] = { math: r[1] === true, econ: r[2] === true, chem: r[3] === true, phys: r[4] === true, eng: r[5] === true, note: String(r[6] || '') };
  });
  return data;
}
function saveProgress_(rows) {
  const sh = sheet_(SHEET, COLS);
  const dates = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  Object.keys(rows || {}).forEach(d => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    const v = rows[d] || {};
    const vals = ['\'' + d, !!v.math, !!v.econ, !!v.chem, !!v.phys, !!v.eng, String(v.note || '').slice(0, 2000), new Date()];
    const i = dates.indexOf(d);
    if (i > 0) sh.getRange(i + 1, 1, 1, vals.length).setValues([vals]);
    else { sh.appendRow(vals); dates.push(d); }
  });
}

// ---------- 週測 ----------
function testMeta_() {
  const sh = sheet_(TSHEET, TCOLS);
  const vals = sh.getDataRange().getDisplayValues();
  const map = {};
  vals.slice(1).forEach((r, i) => { if (r[0]) map[r[0]] = { row: i + 2, name: r[1], subject: r[2], week: r[3], uploaded: r[4], score: r[5], comment: r[6] }; });
  return { sh: sh, map: map };
}
function listTests_() {
  const meta = testMeta_().map;
  const files = folder_().getFiles();
  const list = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== MimeType.PDF) continue;
    const m = meta[f.getId()] || {};
    const name = f.getName();
    const guess = (name.match(/\d{4}-\d{2}-\d{2}/) || [])[0];
    list.push({
      id: f.getId(), name: name, size: f.getSize(),
      subject: m.subject || '', week: m.week || guess || ymd_(f.getDateCreated()),
      uploaded: m.uploaded || ymd_(f.getDateCreated()), score: m.score || '', comment: m.comment || ''
    });
  }
  list.sort((a, b) => (b.week + b.name).localeCompare(a.week + a.name));
  return list;
}
function getFile_(id) {
  const f = DriveApp.getFileById(id);
  const parents = f.getParents();
  const folderId = folder_().getId();
  let ok = false;
  while (parents.hasNext()) if (parents.next().getId() === folderId) ok = true;
  if (!ok) throw new Error('找不到檔案');
  return { name: f.getName(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}
function upload_(b) {
  checkPin_(b.pin);
  const bytes = Utilities.base64Decode(b.data);
  const name = String(b.name || 'test.pdf').replace(/[\\/:*?"<>|]/g, '_');
  const file = folder_().createFile(Utilities.newBlob(bytes, MimeType.PDF, name));
  const t = testMeta_();
  t.sh.appendRow([file.getId(), name, String(b.subject || ''), '\'' + String(b.week || ''), '\'' + ymd_(new Date()), '', '']);
  return file.getId();
}
function score_(b) {
  checkPin_(b.pin);
  const t = testMeta_();
  const m = t.map[b.id];
  const vals = [String(b.score || ''), String(b.comment || '').slice(0, 1000)];
  if (m) t.sh.getRange(m.row, 6, 1, 2).setValues([vals]);
  else {
    const f = DriveApp.getFileById(b.id);
    t.sh.appendRow([b.id, f.getName(), String(b.subject || ''), '\'' + String(b.week || ''), '\'' + ymd_(f.getDateCreated())].concat(vals));
  }
}

// ---------- 入口 ----------
function doGet(e) {
  try {
    const a = (e && e.parameter && e.parameter.action) || 'progress';
    if (a === 'tests') return out_({ ok: true, tests: listTests_() });
    if (a === 'file') return out_(Object.assign({ ok: true }, getFile_(e.parameter.id)));
    return out_({ ok: true, data: getProgress_() });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'upload') return out_({ ok: true, id: upload_(b) });
    if (b.action === 'score') { score_(b); return out_({ ok: true }); }
    if (b.action === 'checkpin') { checkPin_(b.pin); return out_({ ok: true }); }
    saveProgress_(b.rows);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
