function allPutin() {
  let viewpoinValue = '【条件】\n～である場合、\n～すると、\n\n【結果】\n～であること。\nまた、～であること。';
  let advancePreparationValue = ' [SDMC/CL] ～～であること。\n [SDMC] ～～であること。\n [CL] ～～であること。\n [PC] ～～であること。\n----------------------------------------------\n(※)\n [SDMC/CL]   ：　console/端末操作\n [SDMC]         ：　console操作\n [CL]               ：　端末操作\n [PC]               ：　PCのエクスプローラー操作など\n [ブラウザ]      ：　ブラウザ操作\n----------------------------------------------';
  let stepFormatValue = '　. [SDMC/CL] ～する。\n　. [CL] ～する。\n　. [CL]期待動作1を確認する。\n　. [SDMC] ～する。\n　. [SDMC] 期待動作2を確認する。\n　. [PC] ～する。\n　. [PC] 期待動作3を確認する。\n----------------------------------------------\n(※)\n [SDMC/CL]   ：　console/端末操作\n [SDMC]         ：　console操作\n [CL]               ：　端末操作\n [PC]               ：　PCのエクスプローラー操作など\n [ブラウザ]      ：　ブラウザ操作\n----------------------------------------------';
  let expectedBehaviorValue = ' [期待動作1] ～であること。\r\n [期待動作2] ～であること。\r\n [期待動作3] ～であること。' ;

  let setValue1 = [viewpoinValue];
  let setValue2 = [advancePreparationValue];
  let setValue3 = [stepFormatValue];
  let setValue4 = [expectedBehaviorValue];

  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);

  let range1 = sheet1.getRange(range.getRow() , range.getColumn() +1,range.getLastRow()-range.getRow()+1,1);
  range1.setValue(setValue2);

  let range2 = sheet1.getRange(range.getRow() , range.getColumn() +2,range.getLastRow()-range.getRow()+1,1);
  range2.setValue(setValue3);

  let range3 = sheet1.getRange(range.getRow() , range.getColumn() +3,range.getLastRow()-range.getRow()+1,1);
  range3.setValue(setValue4);

}

function viewpointPutin() {
  let setValue = '【条件】\n～である場合、\n～すると、\n\n【結果】\n～であること。\nまた、～であること。';
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function advancePreparationPutin() {
  let setValue = ' [SDMC/CL] ～～であること。\n [SDMC] ～～であること。\n [CL] ～～であること。\n [PC] ～～であること。\n----------------------------------------------\n(※)\n [SDMC/CL]   ：　console/端末操作\n [SDMC]         ：　console操作\n [CL]               ：　端末操作\n [PC]               ：　PCのエクスプローラー操作など\n [ブラウザ]      ：　ブラウザ操作\n----------------------------------------------';
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function stepFormatPutin() {
  let setValue = '　. [SDMC/CL] ～する。\n　. [CL] ～する。\n　. [CL]期待動作1を確認する。\n　. [SDMC] ～する。\n　. [SDMC] 期待動作2を確認する。\n　. [PC] ～する。\n　. [PC] 期待動作3を確認する。\n----------------------------------------------\n(※)\n [SDMC/CL]   ：　console/端末操作\n [SDMC]         ：　console操作\n [CL]               ：　端末操作\n [PC]               ：　PCのエクスプローラー操作など\n [ブラウザ]      ：　ブラウザ操作\n----------------------------------------------';
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function expectedBehaviorPutin() {
  let setValue = ' [期待動作1] ～であること。\r\n [期待動作2] ～であること。\r\n [期待動作3] ～であること。' ;
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function testMemo() {
  let setValue = '※以下、必ずお客様に提出する際には\n備考欄のこの文章を消して下さい\n----' ;
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function testMemo1() {
  let setValue = '※新規開発機能に伴い、本項目を削除しました。\n' ;
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function testMemo2() {
  let setValue = '※以下項目に、統合しました。\nID：\n' ;
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function testMemo3() {
  let setValue = '※以下チケット番号に、注意。\n・#\n\n' ;
  let setValue1 = [setValue];
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet1.getActiveRange();
  range.setValue(setValue1);
}

function wordGet() {
  //スプレッドシートのシートオブジェクトを取得
  let  sheet=SpreadsheetApp.getActiveSheet();
  //選択されているアクティブなセルを取得
  let  sel=sheet.getActiveRange();
　// そのセル範囲にある値の多次元配列を取得
  let values = sel.getValues();
  console.log(values);
}





















