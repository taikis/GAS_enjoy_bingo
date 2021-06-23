function getBingoNum(sellNum) {//ビンゴの番号を作成する　min以上max以下、重複のないsellNum個の数字
  var randoms = []
  var min = 1
  var max = 60
  while (randoms.length < sellNum) {
    var tmp = Math.floor(Math.random() * (max - min + 1)) + min;
    if (!randoms.includes(tmp)) {
      randoms.push(tmp)
    }
  }
  return randoms
}

function sliceArray(array, part) {//配列を二次元配列にする
  var tmp = [];
  for (var i = 0; i < array.length; i += part) {
    tmp.push(array.slice(i, i + part));
  }
  return tmp;
}

function createBingo(sheetNum) {//bingoのシートを作る関数
  /*新しいシートを作る*/
  var objSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newBingoSheet = objSpreadsheet.insertSheet("Bingo" + sheetNum);
  /*幅を揃えて枠線を書く*/
  newBingoSheet.setRowHeights(1, 3, 20);
  newBingoSheet.setColumnWidths(1, 3, 20);
  var bingoSell = newBingoSheet.getRange("A1:C3");
  bingoSell.setBorder(true, true, true, true, true, true);
  /*番号を入れる*/
  var bingoNum = getBingoNum(9);
  var bingoNum2x = sliceArray(bingoNum, 3);
  bingoSell.setValues(bingoNum2x);
  /*真ん中を空白に*/
  newBingoSheet.getRange("B2").clearContent();
  /*整理番号を入れる*/
  newBingoSheet.getRange("A5").setValue(sheetNum);
  /*一旦保存*/
  SpreadsheetApp.flush();
}

function getPdfBingo(bingoID) {//PDFを作る関数。戻り値はPDFの実体
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var key = spreadsheet.getId();
  var bingoSheet = spreadsheet.getSheetByName("Bingo" + bingoID);
  var bingoKey = bingoSheet.getSheetId();
  var token = ScriptApp.getOAuthToken();
  var pdfUrl = "https://docs.google.com/spreadsheets/d/" + key + "/export?gid=" + bingoKey + "&format=pdf&portrait=false&size=A4&gridlines=false&fitw=true";
  var bingoPDF = UrlFetchApp.fetch(pdfUrl, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob().setName("bingo" + bingoID + ".pdf");
  return bingoPDF;
}

function sendMail(mailAdress, bingoPDF, bingoID) {
  var massageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("message");
  var subjectSell = massageSheet.getRange("B1");
  var subject = subjectSell.getValue();
  var bodySell = massageSheet.getRange("B2");
  var body = bingoID;
  GmailApp.sendEmail(mailAdress,
    subject,
    body,
    { attachments: bingoPDF });
}

function mainFun() {
  var adressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("adress");
  var adressSell = adressSheet.getRange("A2");
  var bingoID = 1;
  while (adressSell.getValue() != "") {
    createBingo(bingoID); //bingoをスプレッドシートに作成
    var bingoPDF = getPdfBingo(bingoID); //bingoをPDFに出力
    //sendMail(adressSell.getValue(), bingoPDF, bingoID);
    adressSell = adressSell.offset(1, 0);
    ++bingoID;
  }
}