// Google Apps Script - スプレッドシートから「ch」シートのデータを返す
// デプロイ: Webアプリ / 実行:自分 / アクセス:全員

var SHEET_ID = '1VswrYvsGmY7hKLXwIjh9cptVHqhwthCiy5hatLvn-A4';

function doGet(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var result = [];

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf('ch') === -1) return;

    var data = sheet.getDataRange().getValues();
    var sheetGid = sheet.getSheetId();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // O列(14): 編集D = 伏見 のみ
      if (String(row[14]).trim() !== '伏見') continue;

      // A列(0): 公開予定日が今日以降のみ
      var pubRaw = row[0];
      if (!pubRaw) continue;
      var pubDate = pubRaw instanceof Date ? pubRaw : new Date(pubRaw);
      pubDate.setHours(0, 0, 0, 0);
      if (pubDate < today) continue;

      var rowNum = i + 1;
      var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID
        + '/edit#gid=' + sheetGid + '&range=A' + rowNum;

      result.push({
        id: name + '_' + i,
        pubDate:      fmt(row[0]),
        pubDateMs:    pubDate.getTime(),
        title:        String(row[1] || ''),
        shootDate:    fmt(row[4]),
        handoverDate: fmt(row[5]),
        handoverDone: row[6] === true ? '済み' : '',
        editor:       String(row[15] || ''),
        draftDate:    fmt(row[16]),
        sheetUrl:     sheetUrl,
        sheetName:    name
      });
    }
  });

  result.sort(function(a, b) { return a.pubDateMs - b.pubDateMs; });

  var json = JSON.stringify(result);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function fmt(v) {
  if (!v) return '';
  var d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'M/d');
}
