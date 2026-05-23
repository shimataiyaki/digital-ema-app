// ★★★ スプレッドシートID ★★★
var SPREADSHEET_ID = 'XXXXXXXXXXXXXXXXX';

// ウェブアプリのエントリポイント（JSONP対応APIサーバー）
function doGet(e) {
  var callback = e.parameter.callback || 'callback';
  var type = e.parameter.type;
  var data;

  if (type === 'notification') {
    data = getRecentNicknames();
  } else {
    data = getRandomWishes();
  }

  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  output.setContent(callback + '(' + JSON.stringify(data) + ');');
  return output;
}

// タイムスタンプが今日かどうかを判定
function isToday(timestamp) {
  if (!timestamp) return false;
  var date = new Date(timestamp);
  var today = new Date();
  return date.getFullYear() === today.getFullYear() &&
         date.getMonth() === today.getMonth() &&
         date.getDate() === today.getDate();
}

// ランダムに最大15件の想いを取得（今日のデータのみ）
function getRandomWishes() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var range = sheet.getRange(2, 1, lastRow - 1, 3);
  var values = range.getValues();

  var wishes = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var timestamp = row[0];
    var wish = row[1];
    var nickname = row[2];

    if (wish && wish.toString().trim() !== '' && isToday(timestamp)) {
      wishes.push({
        wish: wish.toString().trim(),
        nickname: nickname ? nickname.toString().trim() : '匿名'
      });
    }
  }

  // シャッフル
  for (var i = wishes.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = wishes[i];
    wishes[i] = wishes[j];
    wishes[j] = temp;
  }

  return wishes.slice(0, 15);
}

// 直近30秒以内に奉納されたニックネームを最大5件返す（今日のデータのみ）
function getRecentNicknames() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var now = new Date();
  var recentNicknames = [];

  for (var i = lastRow; i >= 2; i--) {
    var timestamp = sheet.getRange(i, 1).getValue();
    if (!timestamp) continue;

    if (!isToday(timestamp)) continue;

    var secondsDiff = (now - timestamp) / 1000;
    if (secondsDiff > 30) continue;

    var nickname = sheet.getRange(i, 3).getValue();
    if (nickname && nickname.toString().trim() !== '') {
      var name = nickname.toString().trim();
      if (!recentNicknames.includes(name)) {
        recentNicknames.push(name);
      }
    }

    if (recentNicknames.length >= 5) break;
  }

  return recentNicknames;
}
