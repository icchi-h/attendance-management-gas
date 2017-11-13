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
    //Logger.log(insertedDays);
    //Logger.log(new Date(date.getFullYear(), date.getMonth(), date.getDate()));
    //Logger.log(findDateRow(sheet, new Date(date.getFullYear(), date.getMonth(), date.getDate()), 1));
    if ((lastRow != 1) && findDateRow(sheet, new Date(date.getFullYear(), date.getMonth(), date.getDate()), 1)) {
      Logger.log('Error: in-time@' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd') +' already exists');
      return;
    }
    
    // 挿入
    sheet.getRange(lastRow + 1, 1).setValue(day);
    sheet.getRange(lastRow + 1, 2).setValue(time);
    sheet.getRange(lastRow + 1, 4).setValue("=IF(AND(ISBLANK(B"+(lastRow + 1)+"), ISBLANK(C"+(lastRow + 1)+")), , IF(TIME(10,0,0)-B"+(lastRow + 1)+" > TIME(0,20,0), TIME(10,0,0)-B"+(lastRow + 1)+", 0) + IF(C"+(lastRow + 1)+"-TIME(18,0,0) > TIME(0,10,0), C"+(lastRow + 1)+"-TIME(18,0,0), 0))");
  }
  // 退社処理
  else if (state == 'out') {
    var targetRow = findDateRow(sheet, date, 1);
    if (targetRow != 0) {
      sheet.getRange(targetRow, 3).setValue(time);
    } else {
      Logger.log('Error: in-time@' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd') +' Does not found target date row');
    }
  }
  else {
    Logger.log("Error: invalid state");
  }
}

// 2つの日時オブジェクトが等しいがどうか
function isSameDate(date1,date2){

  if (date1.getFullYear() === date2.getFullYear() && date1.getMonth() === date2.getMonth() && date1.getDate() === date2.getDate()){
    return true;
  } else {
    return false;
  }

}

// ある列から指定の日時がある場合はその行数を返す関数
function findDateRow(sheet,date,col){

  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(var i=1; i<dat.length; i++){
    if(isSameDate(dat[i][col-1], date)){
      return i+1;
    }
  }
  return 0;
}
