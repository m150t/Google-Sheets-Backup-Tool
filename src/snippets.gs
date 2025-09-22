/********** 共通スニペット **********/

function jstDateOnly(d = new Date()) {
  const tz = Session.getScriptTimeZone();
  const s = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
  return new Date(`${s} 00:00:00`);
}

function jstStamp(fmt = 'yyyyMMdd_HHmmss') {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), fmt);
}

function withScriptLock(fn, waitMs = 30000) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    lock.waitLock(waitMs);
    acquired = true;
  } catch (_) {
    throw userError('別のバックアップが進行中です。しばらく待ってから再実行してください。');
  }
  try {
    return fn();
  } finally {
    if (acquired) {
      try { lock.releaseLock(); } catch(_) {}
    }
  }
}

function userError(message) { const e = new Error(message); e.isUser = true; return e; }
function info(message) { try { SpreadsheetApp.getUi().alert(message); } catch(_) {} }
function fail(err, prefix = '') {
  const msg = (err && err.isUser) ? String(err.message)
            : (err && err.stack) ? `${err.message}\n\nStack:\n${err.stack}`
            : String(err);
  Logger.log(msg);
  info(prefix + msg);
}

function getProp(key, { required=false, trim=true } = {}) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  const val = trim ? (raw || '').trim() : (raw ?? '');
  if (required && !val) throw userError(`スクリプトプロパティ「${key}」が未設定です。`);
  return val;
}

function getFolderByIdSafe_(id) {
  try { return DriveApp.getFolderById(id); }
  catch(_) { throw userError(`フォルダIDにアクセスできません: ${id}`); }
}

function sanitizeName_(s) {
  return String(s || '').replace(/[\\/:*?"<>|#\[\]]/g, '').trim() || 'Untitled';
}

function exportSpreadsheetToPdf_(ss, { title = 'backup' } = {}) {
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`
    + `?format=pdf&size=A4&portrait=true&fitw=true`
    + `&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50`
    + `&gridlines=false&printtitle=false&sheetnames=true&pagenumbers=true&fzr=false`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw userError(`PDF出力に失敗（HTTP ${code}）: ${resp.getContentText().slice(0,500)}`);
  }
  return resp.getBlob().setName(`${title}.pdf`);
}

function upsertDailyTrigger(handlerName = 'runBackupNow', hour = 23) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(handlerName).timeBased().atHour(hour).everyDays(1).create();
}
