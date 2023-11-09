//https://caymezon.com/gas-filter/
//フィルター作成
function createFilter(resetsheet) {
  if(resetsheet){
    var sheet = resetsheet
  }else{
    var sheet = SpreadsheetApp.getActiveSheet();
  }

  // A1:B7のセル範囲にフィルター作成
  sheet.getRange(10,1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter();
};

//フィルター取得し、削除
function removeFilter(resetsheet) {
  if(resetsheet){
    var sheet = resetsheet
  }else{
    var sheet = SpreadsheetApp.getActiveSheet();
  }
  // フィルターを取得して削除
  if(sheet.getFilter()){
  sheet.getFilter().remove();

  }
};
function resetFilter(resetsheet){
  removeFilter(resetsheet);
  createFilter(resetsheet);
}

//データフィルタの一括削除
//参考URL:https://www.automation-technology.info/index.php/2021/05/04/post-268/
function removeDetaFiler(){
  // アクティブなスプレッドシートのIDを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = ss.getId();

  //データフィルタを全て取得
  var filters = Sheets.Spreadsheets.getByDataFilter({
      dataFilters: [],
      includeGridData: false
    }, sheetId);
  
  //削除対象のフィルターと削除APIに投げるリクエスト情報を格納
  var requests = []; 
  for (var i = 0; i < filters.sheets.length; i++) {
    var sheets = filters.sheets[i];
    Logger.log(sheets);
    if(!sheets.filterViews){
      continue;
    }
    for (var j = 0; j < sheets.filterViews.length; j++) {
      // 一括削除に投げるフィルタIDをリクエストデータに追加
      requests.push({
        'deleteFilterView': {
          'filterId': sheets.filterViews[j].filterViewId
        }
      });
    }
  }
  //フィルター削除
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({'requests': requests}, sheetId);
    Browser.msgBox('削除処理が完了しました。\\n※表示されている場合は、画面更新をお試しください。')
  }else {
      Browser.msgBox('フィルタが1件もありません。')
      }
}

//特定のデータフィルタのみ削除
//参考URL:https://www.automation-technology.info/index.php/2021/05/04/post-268/
function deleteidentificationDeteFilter(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = ss.getId();

  //フィルターの一覧を取得
  var filters = Sheets.Spreadsheets.getByDataFilter({
      dataFilters: [],
      includeGridData: false
    }, sheetId);
  
  //削除対象のフィルターと削除APIに投げるリクエスト情報を格納
  var requests = []; 
  for (var i = 0; i < filters.sheets.length; i++) {
    var sheets = filters.sheets[i];
    if(!sheets.filterViews){
      continue;
    }
    for (var j = 0; j < sheets.filterViews.length; j++) {
      //フィルタ名にフィルタという名前がつくもののみ削除
      if(sheets.filterViews[j].title.indexOf('フィルタ') > -1){
        requests.push({
          'deleteFilterView': {
            'filterId': sheets.filterViews[j].filterViewId
          }
        });
      }
    }
  }
  //フィルター削除
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({'requests': requests}, sheetId);
    Browser.msgBox('削除処理が完了しました。\\n※表示されている場合は、画面更新をお試しください。')
  }
}


//フィルター基準作成
function newFilterCriteriaSample() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // 'test1','test3'を非表示にするフィルター基準作成
  var criteria = SpreadsheetApp.newFilterCriteria()
  //以下は非表示にする。
  // .setHiddenValues(['再','ー','他',''])
  .setHiddenValues([''])

  .build();

  sheet.getFilter().setColumnFilterCriteria(62, criteria);
}
//
// function newFilterCriteriaSample1() {
//   var sheet = SpreadsheetApp.getActiveSheet();
  
//   // 'test1','test3'を非表示にするフィルター基準作成
//   var criteria = SpreadsheetApp.newFilterCriteria()
//   //以下は非表示にする。
//   .setVisibleValues(['検'])
//   .build();

//   sheet.getFilter().setColumnFilterCriteria(62, criteria);
// }

//指定した列のフィルタ条件を取得
//非表示のフェルタ値を取得
function checkFilter(){
   var sheet = SpreadsheetApp.getActiveSheet().getFilter().getColumnFilterCriteria(58).getHiddenValues();
   console.log(sheet);
}
//表示のフェルタ値を取得//失敗
// function checkFilter1(){
//    var sheet = SpreadsheetApp.getActiveSheet().getFilter().getColumnFilterCriteria(58).getVisibleValues();
//    console.log(sheet);
// }



//指定した列のフィルタ削除
function checkFilterRemve(){
   var sheet = SpreadsheetApp.getActiveSheet().getFilter().removeColumnFilterCriteria(47);
   console.log(sheet);
}
//指定した列のフェルタ条件有無を確認し、設定されていたら削除
function getColumnFilterCriteriaSample() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var filter = sheet.getFilter();
  var criteria = filter.getColumnFilterCriteria(47);
  // 1列目のフィルタ条件を取得
  if(criteria != null){
    // フィルタ条件が設定されている場合は解除
    filter.removeColumnFilterCriteria(47)
  }
}
//最初のフェルタ列位置を取得
function getRangeSample() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var filter = sheet.getFilter();
  if(filter != null){
    var col = filter.getRange().getColumn();
    Logger.log(col);
  }
}
//指定したフィルタ列で昇順、降順にする
function checkFiltersort(){
  //true:昇順、false:降順
   var sheet = SpreadsheetApp.getActiveSheet().getFilter().sort(3, true);
   console.log(sheet);
  //  var sheet = SpreadsheetApp.getActiveSheet().getFilter().sort(49, true);
  //  console.log(sheet);
}

function getVisibleValues() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange('A10:10');
//get hidden values
// var filter = range.getFilter();
// var hiddenValues = filter.getColumnFilterCriteria(47).getHiddenValues();
// console.log(hiddenValues);
//get all values, filter blank
// var all = range.getValues().filter(String);
// console.log(all[0]);
// console.log(all[0].indexOf('メンテナンス'));
var all0 = range.getValues();
console.log(all0[0]);
console.log(all0[0].indexOf('Teams\n/定期検証'));
  var rule = SpreadsheetApp.newFilterCriteria() 
    // //トリガー　セルが空白ではない場合表示する
    // .whenCellNotEmpty()

    //条件　東京と書かれたテキストがある場合表示
    // .whenTextEqualTo('検')

    //以下は非表示にする。
    //.setHiddenValues(['再','ー','他',''])
    .setHiddenValues([''])
    .build();

    
sheet.getFilter().setColumnFilterCriteria(all0[0].indexOf('Teams\n/定期検証')+1, rule);


}
function checklist(){
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('BJ11:BJ');
  // const farstRow =range.getRow();
  // const lastRow = range.getLastRow();
  const selectValues =range.getValues();
  // console.log(range.getRow());
  // console.log(range.getLastRow());
  // console.log(range.getLastRow() - range.getRow());
  console.log(selectValues);



  // let test =sheet.isRowHiddenByUser(1);
  // console.log(test);
  let sercthlist = [[]];

  for(n = 0; n<selectValues.length; ++n){
    if(selectValues[n][0]){
      sercthlist[n] = ['検'];
    }
    else{
      sercthlist[n] = [''];
    }

  }
  console.log(sercthlist);
  console.log(sercthlist.length);
  sheet.getRange('A11:A').setValues(sercthlist);

}




























