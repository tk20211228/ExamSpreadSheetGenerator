function GET_ACTIVE_SPREADSHEETS_NAME()
{
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}
function GET_ACTIVE_SPREADSHEETS_URL()
{
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}


//参考URL：https://www.carefree-it.com/2021/05/gas_099367726.html
//https://qiita.com/kazuph/items/a1eab3054b644edddb49
//https://teratail.com/questions/196315
//https://ytakeuchi.jp/?p=130
function createRecord(e) {

  // var sheet = SpreadsheetApp.getActiveSheet();
  // var sheet1 = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = e.source.getActiveSheet();
  var sheet1 = e.source;

  if (sheet.getName() === "マスター") {
    sheet1.toast("項目修正の編集履歴を確認中…", "実行中", 0);
    var changedCell = sheet.getActiveRange();
    var cellRow = changedCell.getRow();
    // console.log(sheet.getRange(cellRow,2).getValue());
    const masterId = sheet.getRange(cellRow,2).getValue();
    if(masterId =='') {
     sheet1.toast("「マスターID」が空欄であるため、履歴記録なし", "実行完了", 3);
     return; 
    }
    // else if(masterId =='マスターの通番'){
    //  sheet1.toast("履歴記録なし", "実行完了", 2);
    //  return; 

    // }
    var cellCol = changedCell.getColumn();
    var cellLastRow = changedCell.getLastRow()
    var cellLastCol = changedCell.getLastColumn();
    // console.log(cellRow,cellCol,cellLastRow,cellLastCol);
    var sheetBackup = sheet1.getSheetByName('マスター(項目修正記録用バックアップ)');
    if(sheetBackup){
      //変更箇所のマスターID、バックアップIDを取得
      var backupId = sheetBackup.getRange(cellRow,2).getValue();
      //変更箇所のマスター列番号、バックアップ番号取得を取得
      var masterColNo = sheet.getRange(10,cellCol).getValue();
      var backupColNo = sheetBackup.getRange(10,cellCol).getValue();
    }

    if(masterId != backupId || masterColNo != backupColNo || !sheetBackup) {
      if(masterId != backupId){
        var question = Browser.msgBox('変更箇所の「マスターID」がバックアップと一致しません。\\nバックアップを更新してもよろしいでしょうか？\\n※未更新前の編集は記録されません。\\n※マスターを大幅に変更する場合は、シート名を「マスター」以外に設定すると、\\n　このダイヤログは表示されません。', Browser.Buttons.OK_CANCEL);

      }else if(masterColNo != backupColNo){
        var question = Browser.msgBox('変更箇所の「行10」値がバックアップと一致しません。\\nバックアップを更新してもよろしいでしょうか？\\n※未更新前の編集は記録されません。\\n※マスターを大幅に変更する場合は、シート名を「マスター」以外に設定すると、\\n　このダイヤログは表示されません。', Browser.Buttons.OK_CANCEL);

      }else{
        var question = Browser.msgBox('バックアップと一致しません。(原因不明)\\nバックアップを更新してもよろしいでしょうか？\\n※未更新前の編集は記録されません。\\n※マスターを大幅に変更する場合は、シート名を「マスター」以外に設定すると、\\n　このダイヤログは表示されません。', Browser.Buttons.OK_CANCEL);

      }
      if(!sheetBackup && question == 'ok'){
        sheet1.toast("バックアップ更新中", "実行中", 0);

        const newSheet = sheet.copyTo(sheet1);
        newSheet.setName('マスター(項目修正記録用バックアップ)');
        newSheet.hideSheet();

        sheet1.toast("バックアップ更新完了", "実行終了", 3);
        return;

      }
      if(question == 'ok'){
        sheet1.toast("バックアップ更新中", "実行中", 0);
        sheet1.deleteSheet(sheetBackup);
        const newSheet = sheet.copyTo(sheet1);
        newSheet.setName('マスター(項目修正記録用バックアップ)');
        newSheet.hideSheet();

        sheet1.toast("バックアップ更新完了", "実行終了", 3);
      }else{
        sheet1.toast("履歴記録なし", "実行完了", 3);
      }
      return;
    }
    
    // const sheetRecord = e.source.getSheetByName('項目修正シート');
    // const sheetBackup = e.source.getSheetByName('マスター(項目修正記録用バックアップ)');
    var sheetRecord = sheet1.getSheetByName('項目修正シート');
    // var alterDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var alterDate = new Date();

    let user = Session.getActiveUser();
    let contact = ContactsApp.getContact(user);
    if(!contact){
        var fullName = e.user.getEmail();
      }else{
        var fullName = contact.getFamilyName();
      }


    var cellNewwords = changedCell.getValues();//e.value
    var cellOldwords = sheetBackup.getRange(cellRow,cellCol,cellLastRow-cellRow+1,cellLastCol-cellCol+1).getValues()

    // console.log(cellNewword);
    // console.log(cellOldword);

    //マスターシートのフィルター値をリストで取得
    let sheetFilterValues =sheet.getRange('A10:10').getValues();
    //マスターシートに記録する列番号を検索
    let colFilterNo = sheetFilterValues[0].indexOf('修正箇所');
    if( colFilterNo <= cellCol-1) {
     sheet1.toast("行10の[修正箇所]	[修正日時]	[修正者の編集]の編集は、記録されません。", "実行完了", 3);
     return; 
    }
    

    for(n=0;n<cellNewwords.length;n++){
      // console.log(cellID);
      var cellID = sheet.getRange(cellRow+n,2).getValue();
      // if(!cellID) continue;
      var cellNewword = '';
      var cellOldword = '';
      var cellChangedplaces = '';
      for(k=0;k<cellNewwords[n].length;k++){
        // console.log(cellNewword[n][k]);
        // console.log(cellOldword[n][k]);
        

        if(cellNewwords[n][k] != cellOldwords[n][k]){

          var cellChangedplace = sheet.getRange(10,cellCol+k).getValue();
          // var cellNewword = cellNewword + '【'+cellChangedplace+'】'+ '\n'+ cellNewwords[n][k] + '\n' +colFilterNo+ '\n' +Number(cellCol+1);
          var cellNewword = cellNewword + '【'+cellChangedplace+'】'+ '\n'+ cellNewwords[n][k] ;

          var cellOldword = cellOldword + '【'+cellChangedplace+'】'+ '\n'+ cellOldwords[n][k] + '\n' ;
          // if(k!=0)  var cellChangedplaces =  cellChangedplaces + '\n' ;
          var cellChangedplaces = cellChangedplaces + '【'+cellChangedplace+'】';
          sheetBackup.getRange(cellRow+n,cellCol+k).setValue(cellNewwords[n][k]);
        }
      }
      if(cellNewword != ''){
        var value = ['',cellID,cellChangedplaces,cellOldword,cellNewword,'',alterDate,fullName];
        sheetRecord.appendRow(value);
        sheetRecord.getFilter().sort(7 , false);
        var valueMasterRecord =[[cellChangedplaces,alterDate,fullName]];
        
        sheet.getRange(cellRow+n,colFilterNo+1,1,3).setValues(valueMasterRecord).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setHorizontalAlignment('left'); //切り詰める
      }
    }
    sheet1.toast("項目修正の編集履歴を記録完了", "実行終了", 3);
    // var cellID = sheet.getRange(cellRow,2).getValue();
    // var cellChangedplace = sheet.getRange(10,cellCol).getValue();
    // var cellNewword = sheet.getRange(cellRow,cellCol).getValue();//e.value
    // var cellOldword = sheetBackup.getRange(cellRow,cellCol).getValue();//e.oldValue

    // var value = ['',cellID,cellChangedplace,cellOldword,cellNewword,'',alterDate,fullName];
    // sheetRecord.appendRow(value);
    // var valueMasterRecord =[[alterDate,fullName]];
    // sheet.getRange(cellRow,81,1,2).setValues(valueMasterRecord);
    // sheetBackup.getRange(cellRow,cellCol).setValue(cellNewword);

  }
}








































