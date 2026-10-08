var SPREADSHEET_ID = 'ISI_ID_SPREADSHEET_DI_SINI';
var ALLOWED = ['FEED_DASH', 'INS_DASH'];

function doGet(e) {
  var name = (e && e.parameter && e.parameter.sheet) || '';
  if (ALLOWED.indexOf(name) === -1) {
    return ContentService.createTextOutput('Parameter sheet harus salah satu dari: ' + ALLOWED.join(', '))
      .setMimeType(ContentService.MimeType.TEXT);
  }
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
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
