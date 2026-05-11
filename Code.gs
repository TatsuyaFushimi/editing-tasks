// Google Apps Script - スプレッドシート＋カレンダーデータを返す
// デプロイ: Webアプリ / 実行:自分 / アクセス:全員

var SHEET_ID = '1VswrYvsGmY7hKLXwIjh9cptVHqhwthCiy5hatLvn-A4';

// カレンダー準備ルール
var CAL_RULES = [
  { pattern: 'Ch',      daysAhead: 1, type: '定例MTG',   prepLabel: 'FBの準備' },
  { pattern: 'ナレッジ', daysAhead: 3, type: 'ナレッジ',  prepLabel: '準備'    },
  { pattern: '1on1',    daysAhead: 1, type: '1on1',      prepLabel: '準備'    },
];

function doGet(e) {
  var type = e && e.parameter && e.parameter.type;
  if (type === 'calendar') return doGetCalendar(e);
  return doGetSheets(e);
}

function doGetSheets(e) {
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

      if (String(row[14]).trim() !== '伏見') continue;

      var pubRaw = row[0];
      if (!pubRaw) continue;
      var pubDate = pubRaw instanceof Date ? pubRaw : new Date(pubRaw);
      pubDate.setHours(0, 0, 0, 0);
      if (pubDate < today) continue;

      var rowNum = i + 1;
      var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID
        + '/edit#gid=' + sheetGid + '&range=A' + rowNum;

      result.push({
        id:           name + '_' + i,
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

  return respond(e, result);
}

function doGetCalendar(e) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var lookahead = new Date(today);
  lookahead.setDate(lookahead.getDate() + 14);

  var cal = CalendarApp.getDefaultCalendar();
  var events = cal.getEvents(today, lookahead);

  var result = [];

  events.forEach(function(ev) {
    var title = ev.getTitle();
    var start = ev.getStartTime();
    var startDay = new Date(start);
    startDay.setHours(0, 0, 0, 0);
    var daysUntil = Math.round((startDay.getTime() - today.getTime()) / 86400000);
    var timeStr = Utilities.formatDate(start, 'Asia/Tokyo', 'HH:mm');
    var isAllDay = ev.isAllDayEvent();

    var matched = false;
    for (var i = 0; i < CAL_RULES.length; i++) {
      var rule = CAL_RULES[i];
      if (title.indexOf(rule.pattern) !== -1 && daysUntil <= rule.daysAhead) {
        result.push({
          title:      title,
          eventDate:  fmt(start),
          startTime:  isAllDay ? '終日' : timeStr,
          daysUntil:  daysUntil,
          type:       rule.type,
          prepLabel:  rule.prepLabel,
          unknown:    false
        });
        matched = true;
        break;
      }
    }

    // 未分類：7日以内のイベントを返す（都度確認用）
    if (!matched && daysUntil <= 7) {
      result.push({
        title:      title,
        eventDate:  fmt(start),
        startTime:  isAllDay ? '終日' : timeStr,
        daysUntil:  daysUntil,
        type:       'unknown',
        prepLabel:  '',
        unknown:    true
      });
    }
  });

  result.sort(function(a, b) { return a.daysUntil - b.daysUntil; });

  return respond(e, result);
}

function respond(e, data) {
  var json = JSON.stringify(data);
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
