function resultreport() {
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totallsheet = mainsheet.getSheetByName('集計シート');
  let sheetURL = mainsheet.getUrl();
  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です。
  const value = totallsheet.getRange(44,6,1,8).getValues();

  const NGcheck = totallsheet.getRange(11,16).getValue();
  const NAcheck = totallsheet.getRange(11,22).getValue();
  const Keepcheck = totallsheet.getRange(11,30).getValue();
  const Outcheck = totallsheet.getRange(11,26).getValue();

    var NGbody = '';
    var NewNGcount = 0;

    var NewNGbody = '◆ 新規バグ\n';
    var ExitNGbody = '◆ 既存バグ\n';
  if(NGcheck != ''){
  console.log('NGcheck');
    const NGvaluelastRow = totallsheet.getRange(10,16).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NGvalue = totallsheet.getRange(11,16,NGvaluelastRow-10,5).getDisplayValues();//11,16,14,5

    for(i=0; i<NGvalue.length; ++i){
      if(NGvalue[i][4] == '新規バグ'){
        NewNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
        NewNGcount ++;
      }else if(NGvalue[i][4] == '既存バグ'){
        ExitNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }else{
        NGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }
    }
  }

  if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = '・なし';  
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = ExitNGbody + '\n'+NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody ;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+NGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody + '\n'+NGbody;
  }
  var NGreport = '　　新規バグはありませんでした。';
  if(NewNGbody != '◆ 新規バグ\n'){
    var NGreport = '　　新規バグが '+NewNGcount+ '件ありました。';
  }

  if(NAcheck != ''){
    const NAvaluelastRow = totallsheet.getRange(10,22).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NAvalue = totallsheet.getRange(11,22,NAvaluelastRow-10,3).getDisplayValues();
    var NAbody = '';
    for(i=0;i<NAvalue.length;++i){
      NAbody += '　・' + NAvalue[i][0] + NAvalue[i][2] + '\n'+ '　　→'+ NAvalue[i][1] + '\n';
    }
  }
  var NAbreakdown = '・なし';
  if(NAcheck !=''){
    var NAbreakdown = NAbody;
  }

  if(Keepcheck != ''){
    const KeepvaluelastRow = totallsheet.getRange(10,30).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Keepvalue = totallsheet.getRange(11,30,KeepvaluelastRow-10,3).getDisplayValues();
    var Keepbody = '';
    for(i=0;i<Keepvalue.length;++i){
      Keepbody += '　・' + Keepvalue[i][0] + Keepvalue[i][2] + '\n'+ '　　→'+ Keepvalue[i][1] + '\n';
      }     
  }
  var Keepbreakdown = '・なし';
  if(Keepcheck !=''){
    var Keepbreakdown = Keepbody;
  }

  if(Outcheck != ''){
    const OutvaluelastRow = totallsheet.getRange(10,26).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Outvalue = totallsheet.getRange(11,26,OutvaluelastRow-10,3).getDisplayValues();
    var Outbody = '';
    for(i=0;i<Outvalue.length;++i){
      Outbody += '　・' + Outvalue[i][0] + Outvalue[i][2] + '\n'+ '　　→'+ Outvalue[i][1] + '\n';
    }
  }
  var Outbreakdown = '・なし';
  if(Outcheck !=''){
    var Outbreakdown = Outbody;
  }


  const values1 =[mainsheet.getName(),'上記、検証が終了しました。','ご確認をよろしくお願いいたします。','■　検証結果','　------------------------------','　　総項目数：'+value[0][7],'　　OK：'+value[0][1],'　　NG：'+value[0][2] ,'　　N/A：'+value[0][3] ,'　　保留：'+value[0][5] ,'　　試験対象外：'+value[0][4] ,'　------------------------------','■　特記事項',NGreport,'','【NG内訳】',NGbreakdown,'【N/A内訳】',NAbreakdown,'【保留内訳】',Keepbreakdown,'【試験対象外内訳】',Outbreakdown,'【検証書】',sheetURL]; 
  let body = '';
  for(i=0; i<values1.length; i++){
      body += values1[i] + '\n';
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1600)
             .setHeight(700);
  ui.showModelessDialog(html,'検証結果');
}


function Slackresultreport() {
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totallsheet = mainsheet.getSheetByName('集計シート');
  let sheetURL = mainsheet.getUrl();
  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です。
  const value = totallsheet.getRange(44,6,1,8).getValues();

  const NGcheck = totallsheet.getRange(11,16).getValue();
  const NAcheck = totallsheet.getRange(11,22).getValue();
  const Keepcheck = totallsheet.getRange(11,30).getValue();
  const Outcheck = totallsheet.getRange(11,26).getValue();

    var NGbody = '';
    var NewNGcount = 0;

    var NewNGbody = '◆ 新規バグ\n';
    var ExitNGbody = '◆ 既存バグ\n';
  if(NGcheck != ''){
  console.log('NGcheck');
    const NGvaluelastRow = totallsheet.getRange(10,16).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NGvalue = totallsheet.getRange(11,16,NGvaluelastRow-10,5).getDisplayValues();//11,16,14,5

    for(i=0; i<NGvalue.length; ++i){
      if(NGvalue[i][4] == '新規バグ'){
        NewNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
        NewNGcount ++;
      }else if(NGvalue[i][4] == '既存バグ'){
        ExitNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }else{
        NGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }
    }
  }

  if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = '・なし';  
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = ExitNGbody + '\n'+NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody ;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+NGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody + '\n'+NGbody;
  }
  var NGreport = '　　新規バグはありませんでした。';
  if(NewNGbody != '◆ 新規バグ\n'){
    var NGreport = '　　新規バグが '+NewNGcount+ '件ありました。';
  }

  if(NAcheck != ''){
    const NAvaluelastRow = totallsheet.getRange(10,22).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NAvalue = totallsheet.getRange(11,22,NAvaluelastRow-10,3).getDisplayValues();
    var NAbody = '';
    for(i=0;i<NAvalue.length;++i){
      NAbody += '　・' + NAvalue[i][0] + NAvalue[i][2] + '\n'+ '　　→'+ NAvalue[i][1] + '\n';
    }
  }
  var NAbreakdown = '・なし';
  if(NAcheck !=''){
    var NAbreakdown = NAbody;
  }

  if(Keepcheck != ''){
    const KeepvaluelastRow = totallsheet.getRange(10,30).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Keepvalue = totallsheet.getRange(11,30,KeepvaluelastRow-10,3).getDisplayValues();
    var Keepbody = '';
    for(i=0;i<Keepvalue.length;++i){
      Keepbody += '　・' + Keepvalue[i][0] + Keepvalue[i][2] + '\n'+ '　　→'+ Keepvalue[i][1] + '\n';
      }     
  }
  var Keepbreakdown = '・なし';
  if(Keepcheck !=''){
    var Keepbreakdown = Keepbody;
  }

  if(Outcheck != ''){
    const OutvaluelastRow = totallsheet.getRange(10,26).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Outvalue = totallsheet.getRange(11,26,OutvaluelastRow-10,3).getDisplayValues();
    var Outbody = '';
    for(i=0;i<Outvalue.length;++i){
      Outbody += '　・' + Outvalue[i][0] + Outvalue[i][2] + '\n'+ '　　→'+ Outvalue[i][1] + '\n';
    }
  }
  var Outbreakdown = '・なし';
  if(Outcheck !=''){
    var Outbreakdown = Outbody;
  }


  const values1 =[mainsheet.getName(),'上記、検証が終了しました。','ご確認をよろしくお願いいたします。','■　検証結果','　------------------------------','　　総項目数：'+value[0][7],'　　OK：'+value[0][1],'　　NG：'+value[0][2] ,'　　N/A：'+value[0][3] ,'　　保留：'+value[0][5] ,'　　試験対象外：'+value[0][4] ,'　------------------------------','■　特記事項',NGreport,'','【NG内訳】','```',NGbreakdown,'```','【N/A内訳】','```',NAbreakdown,'```','【保留内訳】','```',Keepbreakdown,'```','【試験対象外内訳】','```',Outbreakdown,'```','【検証書】','```',sheetURL,'```']; 
  let body = '';
  for(i=0; i<values1.length; i++){
      body += values1[i] + '\n';
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1600)
             .setHeight(700);
  ui.showModelessDialog(html,'検証結果');
}


function redmienresultreport() {
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totallsheet = mainsheet.getSheetByName('集計シート');
  let sheetURL = mainsheet.getUrl();
  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です。
  const value = totallsheet.getRange(44,6,1,8).getValues();

  const NGcheck = totallsheet.getRange(11,16).getValue();
  const NAcheck = totallsheet.getRange(11,22).getValue();
  const Keepcheck = totallsheet.getRange(11,30).getValue();
  const Outcheck = totallsheet.getRange(11,26).getValue();

    var NGbody = '';
    var NewNGcount = 0;

    var NewNGbody = '◆ 新規バグ\n';
    var ExitNGbody = '◆ 既存バグ\n';
  if(NGcheck != ''){
  console.log('NGcheck');
    const NGvaluelastRow = totallsheet.getRange(10,16).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NGvalue = totallsheet.getRange(11,16,NGvaluelastRow-10,5).getDisplayValues();//11,16,14,5

    for(i=0; i<NGvalue.length; ++i){
      if(NGvalue[i][4] == '新規バグ'){
        NewNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
        NewNGcount ++;
      }else if(NGvalue[i][4] == '既存バグ'){
        ExitNGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }else{
        NGbody += '　・' + NGvalue[i][0] + NGvalue[i][1] + NGvalue[i][2] + '\n' + '　　→'+ NGvalue[i][3] + '\n';
      }
    }
  }

  if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = '・なし';  
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = ExitNGbody + '\n'+NGbody;
  }else if(NewNGbody == '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody ;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody == ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody == '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+NGbody;
  }else if(NewNGbody != '◆ 新規バグ\n' && ExitNGbody != '◆ 既存バグ\n' && NGbody != ''){
    var NGbreakdown = NewNGbody + '\n'+ExitNGbody + '\n'+NGbody;
  }
  var NGreport = '　　新規バグはありませんでした。';
  if(NewNGbody != '◆ 新規バグ\n'){
    var NGreport = '　　新規バグが '+NewNGcount+ '件ありました。';
  }

  if(NAcheck != ''){
    const NAvaluelastRow = totallsheet.getRange(10,22).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const NAvalue = totallsheet.getRange(11,22,NAvaluelastRow-10,3).getDisplayValues();
    var NAbody = '';
    for(i=0;i<NAvalue.length;++i){
      NAbody += '　・' + NAvalue[i][0] + NAvalue[i][2] + '\n'+ '　　→'+ NAvalue[i][1] + '\n';
    }
  }
  var NAbreakdown = '・なし';
  if(NAcheck !=''){
    var NAbreakdown = NAbody;
  }

  if(Keepcheck != ''){
    const KeepvaluelastRow = totallsheet.getRange(10,30).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Keepvalue = totallsheet.getRange(11,30,KeepvaluelastRow-10,3).getDisplayValues();
    var Keepbody = '';
    for(i=0;i<Keepvalue.length;++i){
      Keepbody += '　・' + Keepvalue[i][0] + Keepvalue[i][2] + '\n'+ '　　→'+ Keepvalue[i][1] + '\n';
      }     
  }
  var Keepbreakdown = '・なし';
  if(Keepcheck !=''){
    var Keepbreakdown = Keepbody;
  }

  if(Outcheck != ''){
    const OutvaluelastRow = totallsheet.getRange(10,26).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
    const Outvalue = totallsheet.getRange(11,26,OutvaluelastRow-10,3).getDisplayValues();
    var Outbody = '';
    for(i=0;i<Outvalue.length;++i){
      Outbody += '　・' + Outvalue[i][0] + Outvalue[i][2] + '\n'+ '　　→'+ Outvalue[i][1] + '\n';
    }
  }
  var Outbreakdown = '・なし';
  if(Outcheck !=''){
    var Outbreakdown = Outbody;
  }

  const values1 =[mainsheet.getName(),'上記、検証が終了しました。','ご確認をよろしくお願いいたします。','■　検証結果','　------------------------------','　　総項目数：'+value[0][7],'　　OK：'+value[0][1],'　　NG：'+value[0][2] ,'　　N/A：'+value[0][3] ,'　　保留：'+value[0][5] ,'　　試験対象外：'+value[0][4] ,'　------------------------------','■　特記事項',NGreport,'','【NG内訳】','<pre>',NGbreakdown,'</pre>','【N/A内訳】','<pre>',NAbreakdown,'</pre>','【保留内訳】','<pre>',Keepbreakdown,'</pre>','【試験対象外内訳】','<pre>',Outbreakdown,'</pre>','【検証書】',sheetURL]; 
  let body = '';
  for(i=0; i<values1.length; i++){
      body += values1[i] + '\n';
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1600)
             .setHeight(700);
  ui.showModelessDialog(html,'検証結果');
}

function SlackFormat() {
  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です。

  const values1 =['@shimizu_satoshi  ','@(担当開発者名)','@natori tomose','(CC: @toyoshima @Yukiko Fukasawa @Okawa @Rei Takagi @kuboki )','【検証の対応名】　報告スレッド','このスレッドに、【検証の対応名】の検証結果を、報告をさせていただきます。']; 
  let body = '';
  for(i=0; i<values1.length; i++){
      body += values1[i] + '\n';
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1600)
             .setHeight(700);
  ui.showModelessDialog(html,'検証結果');
}


function TerminalInspectionReport(){
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = mainsheet.getSheetByName('検証結果');
  const testSheetlastRow = testSheet.getLastRow();
  console.log(testSheetlastRow);
  const reportSheet = mainsheet.getSheetByName('ＳＤＭ端末検証報告書');

  const result = testSheet.getRange(11,14,testSheetlastRow-10,2).getValues();
  const funcTitle = testSheet.getRange(11,5,testSheetlastRow-10,1).getValues();

  // console.log(result);

  var date = Utilities.formatDate( new Date(), 'Asia/Tokyo', 'MM/dd');

  
  var terminalBug = '';
  var terminalBuglist = [];


  var osBig = '';
  var osBiglist = [];

  var others = '';
  var otherslist = [];

  var noTest = '';
  var noTestlist = [];

  for(a=0;a<result.length;++a){//《端末固有の注意事項》,《Android OSバージョン依存の注意事項》,《その他の注意事項》,《試験未実施項目》
    // var word = '　　　・';
    // console.log(result[a][0]);
    // console.log(result[a][1]);
    // console.log(result[a][1].replace(/\n/g,'\n　　　'));


    switch(result[a][0]){
      case '《端末固有の注意事項》' :
        if(!terminalBuglist.includes(result[a][1])){
          terminalBuglist.push(result[a][1]);
          terminalBug += '　　　'　+　result[a][1].replace(/\n/g,'\n　　　')+ '\n';
        }
         break;

      case '《Android OSバージョン依存の注意事項》' :
        if(!osBiglist.includes(result[a][1])){
          osBiglist.push(result[a][1]);
          osBig += '　　　'　+　result[a][1].replace(/\n/g,'\n　　　')+ '\n';
        }
         break;

      case '《その他の注意事項》' :
        if(!otherslist.includes(result[a][1])){
          otherslist.push(result[a][1]);
          others += '　　　'　+　result[a][1].replace(/\n/g,'\n　　　')+ '\n';
        }
         break;

      case '《試験未実施項目》' :
        if(!noTestlist.includes(funcTitle[a][0])){
          noTestlist.push(funcTitle[a][0]);
          noTest += '　　　・' + funcTitle[a][0] +'の制御関連'+ '\n';
          }
         break;

      default :
         break;   
    }   
  }//<<for(a=0;)>>
  //  console.log(noTestlist);

  
  const boby = ['検証結果　'+ date +'\n'+'　本端末では、一部正常に動作しない機能がございますので、ご注意願います。'+'\n'+'\n'+ '　《端末固有の注意事項》'+'\n'+terminalBug +'\n'+'　《Android OSバージョン依存の注意事項》'+'\n'+ osBig +'\n'+  '　《その他の注意事項》'+'\n'+ others +'\n'+ '　《試験未実施項目》以下に機能ついては未実施となります。'+'\n'+ noTest];

  console.log(boby);

  reportSheet.getRange(19,2).setValue(boby);

}

function motionConfirmationlist(){
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = mainsheet.getSheetByName('検証結果');
  const testSheetlastRow = testSheet.getLastRow();

  const result = testSheet.getRange(11,14,testSheetlastRow-10,2).getValues();
  const funcTitle = testSheet.getRange(11,5,testSheetlastRow-10,1).getValues();
  
  var terminalBug = '';
  var terminalBuglist = [];

  for(a=0;a<result.length;++a){//《端末固有の注意事項》,《Android OSバージョン依存の注意事項》,《その他の注意事項》,《試験未実施項目》
   if(result[a][0] == '《端末固有の注意事項》'){
     if(!terminalBuglist.includes(result[a][1])){
          terminalBuglist.push(result[a][1]);
          terminalBug += result[a][1]+ '\n';
        }

   }
  
  }//<<for(a=0;)>>
  //  console.log(noTestlist);

  var ui = SpreadsheetApp.getUi();
  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = terminalBug;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1600)
             .setHeight(700);
  ui.showModelessDialog(html,'動作確認済み一覧（※端末の注意事項のみ）');

}

function createExlsxfile() {
  //作業シート取得
  const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // const mainSheet = SpreadsheetApp.getActiveSheet();
  const mainfailId = mainSpreadsheet.getId();
  //参考URL　https://teratail.com/questions/172850
  const mainfolderId = DriveApp.getFileById(mainfailId).getParents().next().getId();
  // console.log(mainfailId);
  // console.log(mainfolderId);

  var mainfile = DriveApp.getFileById(mainfailId);
  var mainfolder = DriveApp.getFolderById(mainfolderId);
  var copyFile = mainfile.makeCopy(mainfolder);
  // console.log(copyFile.getId());

  var spreadsheet_id = copyFile.getId();
  var copymainSheet = SpreadsheetApp.openById(spreadsheet_id);
  var copymainSheets = copymainSheet.getSheets();
  copymainSheet.rename('端末検証依頼書Ver1.0_【XXXX】_XXXX様');

  for(i=0;i<copymainSheets.length;++i){
    var sheetName = copymainSheets[i].getSheetName();
    let sht = copymainSheet.getSheetByName(sheetName);
    if(sheetName =='ＳＤＭ端末検証報告書' ||sheetName =='検証端末情報' ||sheetName =='Activity' ||sheetName =='検証結果' ||sheetName =='独自機能調査' ){
      continue;

    }
      copymainSheet.deleteSheet(sht);

  }

  // let sht = copymainSheet.getSheetByName('集計シート');
  // if(sht) copymainSheet.deleteSheet(sht);

  // let sht1 = copymainSheet.getSheetByName('試験表コピー管理');
  // if(sht1) copymainSheet.deleteSheet(sht1);

  let sht2 = copymainSheet.getSheetByName('検証結果');
  if(sht2){
    // sht2.getRange('C:C').shiftColumnGroupDepth(-1);
    sht2.getRange('1:9').shiftRowGroupDepth(-9);
    sht2.getRange('A11:M').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    sht2.deleteColumn(3);
    sht2.deleteColumns(13,2);
  }
  let sht3 = copymainSheet.getSheetByName('ＳＤＭ端末検証報告書');
  if(sht3){
    sht3.deleteColumn(8);
  }
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
  var contents = "<script>window.open('" + url + "', '_blank').focus();google.script.host.close();</script>";
  var htmlOutput = HtmlService
      .createHtmlOutput(contents)//変数contentsで作成したhtmlを入力
      .setWidth(500) //幅250のサイズを設定
      .setHeight(50); //高さ300のサイズを設定
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ダウンロードを実行中');

  DriveApp.getFileById(spreadsheet_id).setTrashed(true);

}




function outputReport(){
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const reportValues = mainsheet.getRange("A1:A").getValues();

  let body = '';
  for(i=0; i<reportValues.length; i++){
    if(reportValues[i] != "<pre>"　&& reportValues[i] != "</pre>"　){
      body += reportValues[i] + '\n';
      
    }
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1000)
             .setHeight(700);
  ui.showModelessDialog(html,'出力結果');
}

function outputReport_toSlack(){
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const reportValues = mainsheet.getRange("A1:A").getValues();

  let body = '';
  for(i=0; i<reportValues.length; i++){
    if(reportValues[i] == '<pre>' || reportValues[i] == '</pre>' ){
      console.log(reportValues[i]);

      body += reportValues[i].toString().replace(/<pre>|<\/pre>/,"```") + '\n';
      
    }else{
      body += reportValues[i] + '\n';
    }
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1000)
             .setHeight(700);
  ui.showModelessDialog(html,'出力結果');
}

function outputReport_toRedmine(){
  let mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const reportValues = mainsheet.getRange("A1:A").getValues();

  let body = '';
  for(i=0; i<reportValues.length; i++){
      body += reportValues[i] + '\n';
      
  }

  var output = HtmlService.createTemplateFromFile('index-報告書')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(1000)
             .setHeight(700);
  ui.showModelessDialog(html,'出力結果');
}

































