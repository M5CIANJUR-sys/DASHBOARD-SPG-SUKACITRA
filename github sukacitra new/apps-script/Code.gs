var SPREADSHEET_IDS = [
  '1M_bMF0OQg2SqYQJlMzrpzy3wpCgKDc5OsZqGiYEzoNI',
  '1TN90W3XfiruQVvrp6ZCZw0mnTanLMq-9apJUGZdsDuY'
];
var ALLOWED = ['FEED_DASH', 'INS_DASH'];

function doGet(e) {
  var name = (e && e.parameter && e.parameter.sheet) || '';
  if (ALLOWED.indexOf(name) === -1) {
    return ContentService.createTextOutput('Parameter sheet harus salah satu dari: ' + ALLOWED.join(', '))
      .setMimeType(ContentService.MimeType.TEXT);
  }
  var sh = null;
  for (var i = 0; i < SPREADSHEET_IDS.length && !sh; i++) {
    try {
      sh = SpreadsheetApp.openById(SPREADSHEET_IDS[i]).getSheetByName(name);
    } catch (err) {}
  }
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
