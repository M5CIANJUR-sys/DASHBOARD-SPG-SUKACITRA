/**
 * Dashboard SPG - API CSV untuk sheet FEED_DASH dan INS_DASH.
 * Pasang di Google Sheets: Extensions > Apps Script, paste file ini.
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 *
 * URL hasil:
 *   <WEB_APP_URL>?sheet=FEED_DASH
 *   <WEB_APP_URL>?sheet=INS_DASH
 */
var ALLOWED = ['FEED_DASH', 'INS_DASH'];

function doGet(e) {
  var name = (e && e.parameter && e.parameter.sheet) || '';
  if (ALLOWED.indexOf(name) === -1) {
    return ContentService.createTextOutput('Parameter sheet harus salah satu dari: ' + ALLOWED.join(', '))
      .setMimeType(ContentService.MimeType.TEXT);
  }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) {
    return ContentService.createTextOutput('Sheet tidak ditemukan: ' + name)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  var rows = sh.getDataRange().getValues();
  var csv = rows
    .filter(function (r) { return r.some(function (c) { return c !== '' && c !== null; }); })
    .map(function (r) { return r.map(toCell).join(','); })
    .join('\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function toCell(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  v = String(v);
  return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}
