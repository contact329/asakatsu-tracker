/*****************************************************************
 * 朝活トラッカー — スプレッドシート連携バックエンド (Google Apps Script)
 * 個人アカウント・スタンドアロン版（SHEET_ID で対象シートを開く）
 *
 * 対象シート「出席表」の構造（自動検出）:
 *   - ヘッダー行は「高階」「はる」を含む行（1〜6行目から自動検出）
 *   - A列=日付 / D=高階 E=はる F=竹花 / L=担当 M=テーマ
 *   - 今日のテーマ・担当は「出席表の今日の行」から取得
 *****************************************************************/

var MEMBERS = ['高階', 'はる', '竹花']; // 記録対象
var SHEET_ID = '1Vl7CfmmjDAkTJNrtuLgG0Gm29NwAW9737uetvSwHV3Q';
function getSS() { return SpreadsheetApp.openById(SHEET_ID); }

function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || 'data';
  if (a === 'data') return json(getData());
  if (a === 'debug') return json(debugSheets());
  return json({ ok: false, error: 'unknown action: ' + a });
}
function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (b.action === 'checkin') return json(checkin(b.date, b.members));
    if (b.action === 'readBook') return json(readBook(b.date, b.book));
    if (b.action === 'saveTheme') return json(saveTheme(b.date, b.theme, b.owner));
    return json({ ok: false, error: 'unknown action' });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ヘルパ ---------- */
function stripTime(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function fmt(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy/MM/dd'); }
function colIndex(header, kw) {
  for (var i = 0; i < header.length; i++) { if (String(header[i]).indexOf(kw) >= 0) return i; }
  return -1;
}

// ヘッダー行（高階・はるを含む行）を1〜6行目から探す。1-indexedで返す（無ければ0）
function findHeaderRow(sh) {
  var maxScan = Math.min(6, sh.getLastRow());
  if (maxScan < 1 || sh.getLastColumn() < 1) return 0;
  var vals = sh.getRange(1, 1, maxScan, sh.getLastColumn()).getValues();
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i].map(String);
    if (row.indexOf('高階') >= 0 && row.indexOf('はる') >= 0) return i + 1;
  }
  return 0;
}

// 出席表シート（高階・はるのヘッダーを持ち、今日を含む日付範囲）を探す
function getRecordSheet() {
  var sheets = getSS().getSheets();
  var today = stripTime(new Date());
  var fallback = null;
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var hr = findHeaderRow(sh);
    if (hr === 0) continue;
    fallback = fallback || sh;
    var lastRow = sh.getLastRow();
    if (lastRow > hr) {
      var col = sh.getRange(hr + 1, 1, lastRow - hr, 1).getValues()
        .map(function (r) { return r[0]; }).filter(function (v) { return v instanceof Date; });
      if (col.length) {
        var min = stripTime(col[0]), max = stripTime(col[col.length - 1]);
        if (today >= min && today <= max) return sh;
      }
    }
  }
  return fallback;
}

function headerOf(sh, hr) { return sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }
function memberCols(header) {
  var map = {};
  MEMBERS.forEach(function (m) { var i = header.indexOf(m); if (i >= 0) map[m] = i; });
  return map;
}

/* ---------- データ読み出し（トラッカー用）---------- */
function getData() {
  var sh = getRecordSheet();
  if (!sh) return { ok: false, error: 'record sheet not found' };
  var hr = findHeaderRow(sh);
  var header = headerOf(sh, hr);
  var cols = memberCols(header);
  var cTheme = colIndex(header, 'テーマ');
  var cOwner = colIndex(header, '担当');
  var cBook = colIndex(header, '読書本');
  var lastRow = sh.getLastRow();
  var values = sh.getRange(hr + 1, 1, lastRow - hr, sh.getLastColumn()).getValues();
  var records = {}; MEMBERS.forEach(function (m) { records[m] = []; });
  var dates = [];
  var themes = [];
  var todayBook = '';
  var todayStr = fmt(new Date());
  values.forEach(function (row) {
    var d = row[0];
    if (!(d instanceof Date)) return;
    var ds = fmt(d);
    if (ds > todayStr) return; // 未来は出さない
    dates.push(ds);
    MEMBERS.forEach(function (m) {
      if (cols[m] != null && String(row[cols[m]]).trim() !== '') records[m].push(ds);
    });
    if (cTheme >= 0) { var th = String(row[cTheme]).trim(); if (th) themes.push({ date: ds, theme: th, owner: cOwner >= 0 ? String(row[cOwner]).trim() : '' }); }
    if (cBook >= 0 && ds === todayStr) { todayBook = String(row[cBook] || '').trim(); }
  });
  return { ok: true, dates: dates, records: records, theme: getTheme(sh, hr, header), themes: themes, todayBook: todayBook };
}

/* ---------- チェックイン書き込み ---------- */
function checkin(dateStr, members) {
  var sh = getRecordSheet();
  if (!sh) return { ok: false, error: 'record sheet not found' };
  var hr = findHeaderRow(sh);
  var header = headerOf(sh, hr);
  var cols = memberCols(header);
  var lastRow = sh.getLastRow();
  var colA = sh.getRange(hr + 1, 1, lastRow - hr, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < colA.length; i++) {
    var d = colA[i][0];
    if (d instanceof Date && fmt(d) === dateStr) { rowIdx = hr + 1 + i; break; }
  }
  if (rowIdx < 0) return { ok: false, error: 'date not found: ' + dateStr };
  MEMBERS.forEach(function (m) {
    if (cols[m] != null) sh.getRange(rowIdx, cols[m] + 1).setValue(members.indexOf(m) >= 0 ? 1 : '');
  });
  return { ok: true, date: dateStr, members: members };
}

/* ---------- 今日読んだ本を記録（「読書本」列、無ければ自動追加）---------- */
function readBook(dateStr, book) {
  var sh = getRecordSheet();
  if (!sh) return { ok: false, error: 'record sheet not found' };
  var hr = findHeaderRow(sh);
  var header = headerOf(sh, hr);
  var cBook = colIndex(header, '読書本');
  if (cBook < 0) { // 列が無ければヘッダー行の末尾に追加
    cBook = sh.getLastColumn();
    sh.getRange(hr, cBook + 1).setValue('読書本');
  }
  var lastRow = sh.getLastRow();
  var colA = sh.getRange(hr + 1, 1, lastRow - hr, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < colA.length; i++) {
    var d = colA[i][0];
    if (d instanceof Date && fmt(d) === dateStr) { rowIdx = hr + 1 + i; break; }
  }
  if (rowIdx < 0) return { ok: false, error: 'date not found: ' + dateStr };
  sh.getRange(rowIdx, cBook + 1).setValue(book);
  return { ok: true, date: dateStr, book: book };
}

/* ---------- 今日のテーマ・担当を書き込み（出席表の今日の行）---------- */
function saveTheme(dateStr, theme, owner) {
  var sh = getRecordSheet();
  if (!sh) return { ok: false, error: 'record sheet not found' };
  var hr = findHeaderRow(sh);
  var header = headerOf(sh, hr);
  var cTheme = colIndex(header, 'テーマ');
  var cOwner = colIndex(header, '担当');
  var lastRow = sh.getLastRow();
  var colA = sh.getRange(hr + 1, 1, lastRow - hr, 1).getValues();
  var rowIdx = -1;
  for (var i = 0; i < colA.length; i++) {
    var d = colA[i][0];
    if (d instanceof Date && fmt(d) === dateStr) { rowIdx = hr + 1 + i; break; }
  }
  if (rowIdx < 0) return { ok: false, error: 'date not found: ' + dateStr };
  if (cTheme >= 0) sh.getRange(rowIdx, cTheme + 1).setValue(theme);
  if (owner && cOwner >= 0) sh.getRange(rowIdx, cOwner + 1).setValue(owner);
  return { ok: true, date: dateStr, theme: theme, owner: owner };
}

/* ---------- 今日のテーマ＋担当（出席表の今日の行から）---------- */
function getTheme(sh, hr, header) {
  if (!sh) { sh = getRecordSheet(); if (!sh) return null; hr = findHeaderRow(sh); header = headerOf(sh, hr); }
  var cTheme = colIndex(header, 'テーマ');
  var cOwner = colIndex(header, '担当');
  if (cTheme < 0) return null;
  var todayStr = fmt(new Date());
  var lastRow = sh.getLastRow();
  var values = sh.getRange(hr + 1, 1, lastRow - hr, sh.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    var d = values[i][0];
    if (d instanceof Date && fmt(d) === todayStr) {
      var t = String(values[i][cTheme]).trim();
      if (!t) return null;
      return { text: t, owner: cOwner >= 0 ? String(values[i][cOwner]).trim() : '', sub: '今日のテーマ' };
    }
  }
  return null;
}

/* ---------- デバッグ ---------- */
function debugSheets() {
  var sheets = getSS().getSheets();
  return { ok: true, sheets: sheets.map(function (sh) {
    var n = Math.min(3, sh.getLastRow());
    return {
      name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn(),
      preview: n > 0 ? sh.getRange(1, 1, n, Math.min(sh.getLastColumn(), 14)).getValues() : []
    };
  })};
}
