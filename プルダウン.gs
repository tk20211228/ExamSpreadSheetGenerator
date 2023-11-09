//--------------------------------------------------------------------【基本プルダウン】--------------------------------------------------------------------
//試験対象
function testSubjectList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //プルダウンの選択肢を配列で指定
  const setvalues = ['検', 'ー', '再','他'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(setvalues).build();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(rule);
  }
  
}

//結果
function resultSimpleList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //プルダウンの選択肢を配列で指定
  const setvalues = ['OK', 'NG', 'NA','保留','試験対象外'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(setvalues).build();
  

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(rule);
  }
  
}

//redmine
function redmineList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'W11:W';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//試験時間
function testtimeList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //プルダウンの選択肢を配列で指定
  const setvalues = ['Light','Middle','Heavy','他(放置など)'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(setvalues).build();
  

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(rule);
  }
  
}

//試験環境
function testplaceList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //プルダウンの選択肢を配列で指定
  const setvalues = ['β','β2','Pβ','Az','本番'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(setvalues).build();
  

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(rule);
  }
  
}

//PC:OS Ver.(Console)
function testPCosverList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AH11:AH';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}
function testAllOSverList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AE11:AI';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

function testAndroidverList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AE11:AE';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

function iOSverList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AF11:AF';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

function iPadOSverList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AG11:AG';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

function windowsVerList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AH11:AH';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}
function macOSVerList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'AI11:AI';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//シート一覧
function sheetList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('集計シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'Z11:Z';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//項目修正シート（修正箇所）
function modificationPlaceList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('項目修正シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'N4:N';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//項目修正シート（修正箇所）
function correspondenceSituationList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('項目修正シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'O4:O';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//項目修正シート（対応状況）//Record//
function recordmenList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('項目修正シート');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'P4:P';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//--------------------------------------------------------------------【プルダウン　連動型】--------------------------------------------------------------------
function createLinkListAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');

  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();
  
  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();
  // console.log(values);
  // console.log(values.length);

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'C11:C' ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }

    for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'LE'+ no +':OZ'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+1);
    cel.setDataValidation(setrule);
  }
    for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'PA'+ no +':SV'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+2);
    cel.setDataValidation(setrule);
  }
    for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'SW'+ no +':WR'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+3);
    cel.setDataValidation(setrule);
  }
}

//機能カテゴリ
function createLinkList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');//プルダウンリスト管理
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'C11:C' ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//大項目
function createLinkList1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'LE'+ no +':OZ'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//中項目
function createLinkList2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'PA'+ no +':SV'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//小項目
function createLinkList3() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let no = selectedRow + i ;
    let rangevalues = 'SW'+ no +':WR'+ no ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}


//--------------------------------------------------------------------【プルダウン　簡易型】--------------------------------------------------------------------
function createSimpleListAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');

  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();
  
  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  //機能カテゴリ
  for(let i=0; i<values.length ;i++){
    let rangevalues = 'C11:C' ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }

  //大項目
  for(let i=0; i<values.length ;i++){
    let rangevalues = 'CZ11:CZ';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+1);
    cel.setDataValidation(setrule);
  }

  //中項目
  for(let i=0; i<values.length ;i++){
    let rangevalues = 'GW11:GW';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+2);
    cel.setDataValidation(setrule);
  }

  //小項目
  for(let i=0; i<values.length ;i++){
    let rangevalues = 'KT11:KT';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn+3);
    cel.setDataValidation(setrule);
  }

}

//機能カテゴリ
function createSimpleList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'C11:C' ;
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//大項目
function createSimpleList1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();


  for(let i=0; i<values.length ;i++){
    let rangevalues = 'CZ11:CZ';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//中項目
function createSimpleList2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'GW11:GW';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}

//小項目
function createSimpleList3() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet0 = sheet.getSheetByName('プルダウンリスト管理');
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation();

  //リストをセットするセル範囲を取得
  let mySheet = sheet.getActiveSheet();
  let cel = mySheet.getActiveRange();

  // そのセル範囲にある値の多次元配列を取得
  let values = cel.getValues();

  //アクティブな範囲から最初のRow:行、Column:列を取得する
  let selectedRow = cel.getRow();
  let selectedColumn = cel.getColumn();

  for(let i=0; i<values.length ;i++){
    let rangevalues = 'KT11:KT';
    let setrange =  sheet0.getRange(rangevalues);
    let setrule = rule.requireValueInRange(setrange).build();
    let cel = mySheet.getRange(selectedRow+i,selectedColumn);
    cel.setDataValidation(setrule);
  }
  
}


//--------------------------------------------------------------------【プルダウン　削除】--------------------------------------------------------------------
//選択箇所のプルダウンリスト管理削除
function selectListClear() {
  var range = SpreadsheetApp.getActive().getActiveRange();
  // そのセル範囲に設定されているデータの入力規則のみクリア
  range.clearDataValidations();
}
//結果箇所のプルダウンリスト管理削除
function resultListClear() {
  var range = SpreadsheetApp.getActive().getRange('Q:Q');
  // そのセル範囲に設定されているデータの入力規則のみクリア
  range.clearDataValidations();
}

//カテゴリ箇所のプルダウンリスト管理削除
function categoryListClear() {
  var range = SpreadsheetApp.getActive().getRange('D:G');
  // そのセル範囲に設定されているデータの入力規則のみクリア
  range.clearDataValidations();
}

//redmine箇所のプルダウンリスト管理削除
function redmineListClear() {
  var range = SpreadsheetApp.getActive().getRange('R:R');
  // そのセル範囲に設定されているデータの入力規則のみクリア
  range.clearDataValidations();
}
//シート全体のプルダウンリスト管理削除
function sheetListClear() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var sheet0 = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getA1Notation();
  var range = spreadsheet.getRange(sheet0);
  // そのセル範囲に設定されているデータの入力規則のみクリア
  range.clearDataValidations();
  };
  
//スプレッドシート全体のプルダウンリスト管理削除
function spreadsheetListClear() {
  //シート名一覧を取得
  var sheets = SpreadsheetApp.getActive().getSheets();
  for(var i=0; i<sheets.length; i++) {
      // シートのIDと名前
      var sheetName = sheets[i].getSheetName();
      console.log(sheetName);
      var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var sheet0 = sheet1.getRange(1, 1, sheet1.getMaxRows(), sheet1.getMaxColumns()).getA1Notation();
      var range = sheet1.getRange(sheet0);
      // そのセル範囲に設定されているデータの入力規則のみクリア
      range.clearDataValidations();
  
   }

  };


//--------------------------------------------------------------------【その他　プルダウン】--------------------------------------------------------------------



































