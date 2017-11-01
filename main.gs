var REPORT_DIR_ID = '<YOUR DIR ID>'
var TEMPLATE_SS_ID = '<YOUR TEMPLATE SS ID>'

function doPost(e) {
  Logger.log("GET POST REQUEST");
  Logger.log(e);

  main(e.parameter.state)
}
function doGet(e) {
  Logger.log("GET GET REQUEST");
  Logger.log(e);

  main(e.parameter.state)
}

function main(state) {

  // Dateオブジェクトを取得
  var date = new Date();

  // 編集するシートを取得
  var ss = getSS(date);
  var sheet = ss.getActiveSheet();

  // 編集
  insert(sheet, date, state);
}


function getSS(date) {

  // 日付 yyyy-MMを取得
  dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');

  var folder = DriveApp.getFolderById(REPORT_DIR_ID);
  var ssFileItr = folder.getFilesByName(dateStr);
  var isExist = ssFileItr.hasNext();

  var ss = 0;
  // 今月分の記録シートの存在確認
  if (isExist) {
    ss = SpreadsheetApp.open(ssFileItr.next())
  } else {
    // スプレッドシートがない場合はテンプレートをコピー
    var ssTemplate = DriveApp.getFileById(TEMPLATE_SS_ID);
    var ssCopied = ssTemplate.makeCopy(dateStr, folder);
    Logger.log('Created SpreadSheet: ' + dateStr);

    ss = SpreadsheetApp.open(ssCopied);
  }

  return ss;
}

function insert(sheet, date, state) {

  // 日時を取得
  var day = Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd');
  var time = Utilities.formatDate(date, 'Asia/Tokyo', 'HH:mm');

  // 最下行を取得
  var lastRow = sheet.getLastRow();

  // 入社処理
  if (state == 'in') {
    // 既に本日の入社時刻があるなら終了
    var insertedDays = sheet.getRange(2, 1, lastRow, 1).getValues();
    if ((lastRow != 1) && (new Date(date.getFullYear(), date.getMonth() + 1, date.getDate() in insertedDays))) {
      Logger.log('Error: in-time@' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd') +' already exists')
      return;
    }

    // 挿入
    sheet.getRange(lastRow + 1, 1).setValue(day);
    sheet.getRange(lastRow + 1, 2).setValue(time);
    sheet.getRange(lastRow + 1, 4).setValue("=IF(AND(ISBLANK(B2), ISBLANK(C2)), , IF(TIME(10,0,0)-B2 > TIME(0,20,0), TIME(10,0,0)-B2, 0) + IF(C2-TIME(18,0,0) > TIME(0,10,0), C2-TIME(18,0,0), 0))");
  }
  // 退社処理
  else if (state == 'out') {
    sheet.getRange(lastRow, 3).setValue(time);
  }
  else {
    Logger.log("Error: invalid state");
  }
}
