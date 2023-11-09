// 表示中のシート名だけを取得する
function getShownSheetNames() {
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetNames = [];
  if (allSheets.length >= 1) {
    for (var i = 0; i < allSheets.length; i++) {
      // 表示中のシートなら配列に追加する
      if (!allSheets[i].isSheetHidden()) {
        sheetNames.push(allSheets[i].getName());
      }
    }
  }
  return sheetNames;
}