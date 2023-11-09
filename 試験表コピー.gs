function copysheet() {
  //getActiveSpreadsheet()現在アクティブなスプレッドシートを返すかnull、スプレッドシートがない場合は返します。
  var sh = SpreadsheetApp.getActiveSpreadsheet();

  //コピーの確認
  var result = Browser.msgBox('シートをコピーする','コピーを実行しますか？',Browser.Buttons.OK_CANCEL);
  if(result=='cancel'){
    // Browser.msgBox('Cancel','コピーをキャンセルしました。',Browser.Buttons.OK);
   }else{
       do {
         do{
           var shtname = Browser.inputBox('シート名の設定','シート名を設定しコピーできます。' + '\\n' + '以下に、入力してください。' + '\\n' + '※シート名を設定しない場合は、「キャンセル」を押下してください。' +'\\n'+ '※シート名を設定しない場合は、『試験書(リネームしてください)』に設定されます。',Browser.Buttons.OK_CANCEL);
           switch(shtname){
             case "":
                 Browser.msgBox('シート名を空白で作成することはできません。');
                 break;
             case "悪魔":
                 Browser.msgBox('そんな名前を付けてはいけません');
                 break;
             case "天使":
                 Browser.msgBox('すごい名前ですね！');
                 break;
             default:  
             }
           }while(shtname=="") 

         if(shtname=='cancel'){
           //新シートをshtnameに定義
           var shtname = "試験書(リネームしてください)"
           var sht = sh.getSheetByName(shtname);
           }else{}

           var result2 = Browser.msgBox('シート名の設定','シート名は『' + shtname + '』でよろしいですか？ \\n' + '※「はい」を押下すると、コピーが開始されます。' + ' \\n' + '※「いいえ」を押下すると、再設定できます。' + '\\n' + '※「キャンセル」を押下すると、コピーをキャンセルします。'+ '\\n' + '※同名のシート名がある場合は、自動的にコピーがキャンセルされます。',Browser.Buttons.YES_NO_CANCEL);
         } while(result2=='no')
        var sht = sh.getSheetByName(shtname);
        if(result2=='cancel'){
          // Browser.msgBox('Cancel','コピーをキャンセルしました。',Browser.Buttons.OK);
         }else{
           //既に同じシート名（sht）があったら削除//
           if(!sht){       
             //複製したシートの名前を変更する
             sh.duplicateActiveSheet().setName(shtname);　
             //文字を維持する
             var currentCell = sh.getCurrentCell();
             sh.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
             currentCell.activateAsCurrentCell();
             sh.getRange('10:10').activate();
             sh.getRange('10:2390').copyTo(sh.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

             sh.getRange('A1').activate();
              }else{
                 //そうでなかったら同じシートを削除する
                 Browser.msgBox('Error','同名のシートが、すでに作成されています。\\n' + 'コピーをキャンセルしました。' ,Browser.Buttons.OK);
                 }
              }
   }
};

//参考URL:https://ujimasayuruyuru.blogspot.com/2020/12/gasgoogledrive.html
function copyMastersheet() {
  //コピーの確認
  var result = Browser.msgBox('Googleドライブから試験表をコピーする','\
  完了までに3分程度かかる場合があります。\
  \\n処理を開始してよろしいですか？\
  \\n\
  \\n\※Googleドライブ>「MDMプロジェクト」＞「検証Gr」＞「マスターシート」フォルダ直下にある\
  \\n\　ファイルを一覧で取得し、任意のファイルから試験表マスターをコピーできます。\
  \\n※取得の条件は以下\
  \\n　・スプレッドシートファイル\
  \\n　・スターがついているファイル',Browser.Buttons.OK_CANCEL);
  
  if(result=='cancel'){
    // Browser.msgBox('Cancel','コピーをキャンセルしました。',Browser.Buttons.OK);
   }else{
     /* 基準フォルダから下のフォルダ階層をすべてリストアップする */
      var folderId = "11eYkhvrmYZy2hR9y7-i3Lt2CoAM7yA6Z";

      let key = folderId;
      let stt = DriveApp.getFolderById(key);
      let name = "";
      let i = 0;
      let j = 0;
      let folderlist = [];
      let result = [];

      // 開始フォルダ
      folderlist.push([stt, key]);
      name = stt + '/';


    do {
      //フォルダ一覧取得
      let folders = DriveApp.searchFolders("'"+key+"' in parents");
      //フォルダ一覧からフォルダ名称とIdを取得
      while(folders.hasNext()){
        i++;
        let folder = folders.next();
        let tmparray = new Array();
        tmparray.push(name + folder.getName());
        tmparray.push(folder.getId());
        folderlist.push(tmparray);
      }
      //フォルダ名称とIdを取り出す
      j++;
      if(j <= i){
        name = folderlist[j][0] + '/';
        key = folderlist[j][1];
        }
      } while (j <= i);
    
    let fileArray = [];
    //ファイル情報取得
    for (i=0; i<=folderlist.length-1; i++) {
      let key = DriveApp.getFolderById(folderlist[i][1]).getId();
      let files = DriveApp.searchFiles("starred = true and '"+key+"' in parents");
      while(files.hasNext()){
        fileArray = [];
        let file = files.next();
        fileArray.push(folderlist[i][0]);
        fileArray.push(file.getName());
        let materSheetID = file.getId();
        fileArray.push(materSheetID);
        result.push(fileArray);
        j++;
        }
      }
    console.log(folderlist);
    console.log(result);
    let masterlist =['以下マスターシートの該当番号を入力してください。'+'\\n' + '※マスターのコピーを中止する場合は、「キャンセル」を押下してください。' ];
    let masterNolist = [[],[],[]];

    for( i=1; i<result.length+1; i++){
      masterlist.push([i]+' . '+result[i-1][1]);
      masterNolist[0].push(i);
      masterNolist[1].push(result[i-1][1]);
      masterNolist[2].push(result[i-1][2]);
      }
    
    console.log(masterNolist);
    
           var test = masterlist.join("\\n");
           do{
            var shtname = Browser.inputBox('マスターシートの設定',test,Browser.Buttons.OK_CANCEL);
            console.log(shtname);
            /**
             * 参考URL:https://webllica.com/javascript-number-check-function/
             * 数値チェック関数
             * 入力値が数値 (符号あり小数 (- のみ許容)) であることをチェックする
             * [引数]   numVal: 入力値
             * [返却値] true:  数値
             *          false: 数値以外
             * 数値パターン	正規表現（ゼロ埋めなし）	正規表現 (ゼロ埋めあり)
             * 0以上の整数のみ	/^([1-9]\d*|0)$/	/^\d*$/
             * 整数 (- のみ許容)	/^[-]?([1-9]\d*|0)$/	/^[-]?\d*$/
             * 整数 (+, – 許容)	/^[+,-]?([1-9]\d*|0)$/	/^[+,-]?\d$/
             * 符号なし小数	/^([1-9]\d*|0)(\.\d+)?$/	/^\d(\.\d+)?$/
             * 符号あり小数 (- のみ許容)	/^[-]?([1-9]\d*|0)(\.\d+)?$/	/^[-]?\d*(\.\d+)?$/
             * 符号あり小数 (+, – 許容)	/^[+,-]?([1-9]\d*|0)(\.\d+)?$/	/^[+,-]?\d(\.\d+)?$/
             */
             if(shtname > result.length|shtname == 0){
                var patternTest = false;
               Browser.msgBox('Cancel','該当番号がありません。',Browser.Buttons.OK);

              }else if(shtname != 'cancel'){
                // チェック条件パターン
                var pattern = /^([1-9]\d*|0)(\.\d+)?$/;
                // 数値チェック
                var patternTest = pattern.test(shtname);
                console.log(patternTest);

                if(patternTest == false ){
                  Browser.msgBox('Cancel','番号以外は入力しないでください。',Browser.Buttons.OK);

                }
                console.log(!patternTest);
              }else if(shtname == 'cancel'){
                var patternTest = true;
              }
              console.log(patternTest); 

            }while(!patternTest);
  
            if(shtname=='cancel'){
              // Browser.msgBox('Cancel','コピーをキャンセルしました。',Browser.Buttons.OK);
              }else if(shtname <=result.length){
                var sh = SpreadsheetApp.openById(masterNolist[2][shtname-1]).getSheetByName('マスター');
                  if(sh){
                  var meinsheetID =SpreadsheetApp.getActiveSpreadsheet().getId(); 
                  var newsheet = sh.copyTo(SpreadsheetApp.openById(meinsheetID)).getSheetName();
                  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newsheet).activate();
                  }else{
                   Browser.msgBox('Cancel','該当のスプレッドシートには、シート名「マスター」がありません。\\nシート名が「マスター」でなければコピーすることはできません。',Browser.Buttons.OK);
                  }  
               }
   }
};

function URLcopysheet() {
  //getActiveSpreadsheet()現在アクティブなスプレッドシートを返すかnull、スプレッドシートがない場合は返します。
  var sh = SpreadsheetApp.getActiveSpreadsheet();
           var  url = Browser.inputBox('URL設定','スプレッドシートのURLを以下に入力してください。' + '\\n' + '　' ,Browser.Buttons.OK_CANCEL);

        if( url=='cancel'){
          // Browser.msgBox('Cancel','コピーをキャンセルしました。',Browser.Buttons.OK);
         }else{   
             //複製したシートの名前を変更する
             
             　//参考URL：https://takeda-san.hatenablog.com/entry/2021/06/12/233351
               // URLの3階層目からスプレッドシートID取得
              const regExpSpreadsheetId = new RegExp("https?://.*?/.*?/.*?/(.*?)(?=/)")
              const spreadsheetId = url.match(regExpSpreadsheetId)[1]
              console.log(spreadsheetId);

              // gidパラメータからシートID取得
              const regExpGid = new RegExp("gid=(.*?)(&|$)")
              const gid = url.match(regExpGid)[1]
              console.log(gid);
              var sheets = SpreadsheetApp.openByUrl(url).getSheets();
              for(i=0;i<sheets.length; i++){
              // シートのIDと名前
              var sheetName = sheets[i].getSheetName();
              var sht = sh.getSheetByName(sheetName)
              if(!sht){
                sheets[i].copyTo(sh).setName(sheetName);
              }else{
                sheets[i].copyTo(sh);
              }          
              }


              // var sheetId = sheets[0].getSheetId();

              // // var minesheetID = sh.getId()
              // // var newsheet = sheets[0].copyTo(SpreadsheetApp.openById(minesheetID)).getSheetName();
              // // sh.getSheetByName(newsheet).activate();

             }

   
};

function createMastarSheet(){
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainsheet1 = SpreadsheetApp.getActiveSheet();
  const mainsheet2 = SpreadsheetApp.getActive();
  const mainsheetID =mainsheet.getId(); 
  
  //main sheet名　リスト取得
  const sheets = SpreadsheetApp.getActive().getSheets();
  // sheet名　リスト取得の配列
  let sheetNameList = [];
  for(k=0; k<sheets.length; k++) {
    // シートのIDと名前
    let sheetName = sheets[k].getSheetName();
    // ハイパーリンク文字列を配列に格納
    sheetNameList[k] = sheetName;
  }
  
  // console.log(sheetNameList)
  // console.log(sheetNameList.includes('集計シート'))
  // console.log(sheetNameList.includes('集計シート12'))

  // const getDeta = mainsheet.getRange('B11:P').getDisplayValues();
  const getRaelDeta = mainsheet.getRange('B11:P').getValues();
  // console.log(getDeta);
  console.log(getRaelDeta);

  //spredsheetURLからIDを取得する識別アイテム
  // const regExpFileURL = new RegExp("https?://.*?/.*?/.*?/(.*?)(?=/)")

  //オプション設定取得
  const sumsheet = mainsheet.getSheetByName('集計シート');
  const getoptionset = mainsheet1.getRange(3,8,4,3).getValues();
  console.log(getoptionset);
  console.log(getoptionset[2][1]);
  if(getoptionset[2][1] == true){
      if(!sumsheet) {
       Browser.msgBox('エラー','シート名「集計シート」がないため、集計シートに集計シートの既存情報は削除できません。',Browser.Buttons.OK);
      }else{
        let inputlastRow = sumsheet.getRange(9,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
        console.log(inputlastRow);
        if(inputlastRow > 10) {
          sumsheet.getRange(11,2,inputlastRow-10,1).clearContent();
        }

      }
  }

　//ループ処理開始（i=0）
  for(i=0; i<getRaelDeta.length; ++i){
    //B列にチェック有の場合実行
    // if(getRaelDeta[i][0] == true){
    if(getRaelDeta[i][0] == true ){
    
      // var fileId = getRaelDeta[i][2].match(regExpFileURL);      
      // console.log(fileId[1]);
      // console.log(sheetNameList.includes(getRaelDeta[i][4]));

      // var copySheet = SpreadsheetApp.openById(fileId[1]).getSheetByName(getRaelDeta[i][3]);
      console.log(getRaelDeta[i]);
      var copySheet = SpreadsheetApp.openByUrl(getRaelDeta[i][2]).getSheetByName(getRaelDeta[i][3]);
      var mainaSheetName = mainsheet.getSheetByName(getRaelDeta[i][4]);

      // if(copySheet && sheetNameList.includes(getRaelDeta[i][4]) == false ){ 
      if(copySheet){

        //作業中のスプレッドシートに設定したシート名がなければ、設定したシート名に改変
        if(!mainaSheetName && getRaelDeta[i][4] != ""){
          // var newsheet = copySheet.copyTo(SpreadsheetApp.openById(mainsheetID));
          var newsheet = copySheet.copyTo(mainsheet);
          newsheet.setName(getRaelDeta[i][4])
        }else{
          var newsheet = copySheet.copyTo(mainsheet);
        }

        //実行後オプション　集計シートにシート名を自動入力
        if(getoptionset[1][1] == true) {
          
          if(!sumsheet) Browser.msgBox('エラー','シート名「集計シート」がないため、集計シートに作成したシート名を入力できません。',Browser.Buttons.OK);
          if(!sumsheet) break;
          let inputlastRow = sumsheet.getRange(9,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getLastRow();
          console.log("確認："+　inputlastRow);
          sumsheet.getRange(inputlastRow+1,2).setValue(newsheet.getSheetName());
        }
        
        //別で作成したフィルタリセット関数を実行
        const resetsheet = newsheet;
        resetFilter(resetsheet);

        mainsheet1.getRange(i+11, 2).setValue(`FALSE`);
        let newsheetValues = newsheet.getRange('A10:10').getValues();

        //ソート処理
        // if(getRaelDeta[i][10] || getRaelDeta[i][11] ){
        if(getRaelDeta[i][10]){
          //設定したフィルタ名がある列を検索
          let columnsortFilter = newsheetValues[0].indexOf(getRaelDeta[i][10]);
          //検索結果、ない場合は「-1」となり、if処理終了。次の処理に進む
          if(columnsortFilter == -1 ) break;
          //フィルタの「昇順/降順」を識別し、処理、想定外の文字列の場合はなにも処理しない。
          if(getRaelDeta[i][11] == '昇順')newsheet.getFilter().sort(columnsortFilter +1 , true);
          if(getRaelDeta[i][11] == '降順')newsheet.getFilter().sort(columnsortFilter +1 , false);
        }

        //フィルタ処理識別、未設定の場合は、次のループ処理に移行。
        //※フィルタ処理をしないければ、A列に「検」を入れる処理はしない
        if(!getRaelDeta[i][5]) continue;
        // if(!getRaelDeta[i][6]) continue;

        // console.log(newsheetValues[0]);
        // console.log(getRaelDeta[i][5]);

        if(getRaelDeta[i][6] == '設定した文字のみ'){
          //設定したフィルタ名がある列を検索
          var columnFilter = newsheetValues[0].indexOf(getRaelDeta[i][5]);
          // console.log(columnFilter);
          //検索結果、ない場合は「-1」となり、次のループ処理に移行。
          if(columnFilter == -1 ) continue;
          //フィルタ処理
          let rule = SpreadsheetApp.newFilterCriteria() 
                    //セルが空白ではない場合表示する
                    // .whenCellNotEmpty()

                    //条件　東京と書かれたテキストがある場合表示
                    // .whenTextEqualTo('検')
                    .whenTextEqualTo(getRaelDeta[i][8].split(","))

                    //以下は非表示にする。
                    //.setHiddenValues(['再','ー','他',''])

                    .build();

          newsheet.getFilter().setColumnFilterCriteria(columnFilter + 1 , rule);

        }else if(getRaelDeta[i][6] == '設定した文字を含む'){
          //設定したフィルタ名がある列を検索
          var columnFilter = newsheetValues[0].indexOf(getRaelDeta[i][5]);
          // console.log(columnFilter);
          //検索結果、ない場合は「-1」となり、次のループ処理に移行。
          if(columnFilter == -1 ) continue;
          //フィルタ処理
          let rule = SpreadsheetApp.newFilterCriteria() 
                    //セルが空白ではない場合表示する
                    // .whenCellNotEmpty()

                    //条件　東京と書かれたテキストがある場合表示
                    // .whenTextEqualTo('検')
                    .whenTextContains(getRaelDeta[i][8].split(","))

                    //以下は非表示にする。
                    //.setHiddenValues(['再','ー','他',''])

                    .build();

          newsheet.getFilter().setColumnFilterCriteria(columnFilter + 1 , rule);

        }else if(getRaelDeta[i][6] == '設定した文字以外'){
          //設定したフィルタ名がある列を検索
          var columnFilter = newsheetValues[0].indexOf(getRaelDeta[i][5]);
          console.log(columnFilter);
          //検索結果、ない場合は「-1」となり、次のループ処理に移行。
          if(columnFilter == -1 ) continue;
          //フィルタ処理
          let rule = SpreadsheetApp.newFilterCriteria() 
                    //セルが空白ではない場合表示する
                    // .whenCellNotEmpty()

                    //条件　東京と書かれたテキストがある場合表示
                    // .whenTextEqualTo('検')

                    //以下は非表示にする。
                    //.setHiddenValues(['再','ー','他',''])
                    .setHiddenValues(getRaelDeta[i][7].split(","))

                    .build();

          newsheet.getFilter().setColumnFilterCriteria(columnFilter + 1 , rule);

        }else{
          continue;
        }



        //A列に「検」を入れる処理
        // if(getRaelDeta[i][9] == false) continue;
        //falseの場合、次のループ処理に移行。
        if(getRaelDeta[i][9] == false ) continue;

        //コピーしたシート＞フィルタ処理した列のA11:A内になる文字列を取得
        let selectValues = newsheet.getRange(11,columnFilter + 1,newsheet.getMaxRows()-10,1).getValues();
        let sercthlist = [[]];
        var valueCountlist = [[]];
        var Count = 0;
        var regValue = new RegExp(getRaelDeta[i][8]);
        var regValuegroup = new RegExp('('+getRaelDeta[i][8]+')');

        // console.log(regValue);
        for(n = 0; n<selectValues.length; ++n){
          if(getRaelDeta[i][6] == '設定した文字のみ' ){
            //フィルタ設定した文字列が含まていれば、'検'を格納、含まれいなければ、’’を格納
            // if(getRaelDeta[i][8].split(",").includes(selectValues[n][0])){
            if(selectValues[n][0].toString().match(regValue)){

              sercthlist[n] = ['検'];
              valueCountlist[Count] = [ Count + 1 ];
              ++Count;
            }
            else{
              sercthlist[n] = [''];
            }
          }else if(getRaelDeta[i][6] == '設定した文字を含む' ){
            //フィルタ設定した文字列が含まていれば、'検'を格納、含まれいなければ、’’を格納
            // if(getRaelDeta[i][8].split(",").includes(selectValues[n][0])){
            if(selectValues[n][0].toString().match(regValuegroup)){

              sercthlist[n] = ['検'];
              valueCountlist[Count] = [ Count + 1 ];
              ++Count;
            }
            else{
              sercthlist[n] = [''];
            }
          }else if(getRaelDeta[i][6] == '設定した文字以外'){
            //フィルタ設定した文字列が含まなければ、'検'を格納、含まれていれば、’’を格納
            if(!getRaelDeta[i][7].split(",").includes(selectValues[n][0])){
              sercthlist[n] = ['検'];
              valueCountlist[Count] = [ Count + 1 ];
              ++Count;
            }
            else{
              sercthlist[n] = [''];
            }

          }else{
          continue;
          }




          

        }//<<--for(n = 0)-->>

        // console.log(sercthlist);
        // console.log(sercthlist.length);

        //フィルタ設定を元に、複製したシートのA列に検をつける
        newsheet.getRange('A11:A').setValues(sercthlist);
        //連番を入力
        // console.log(valueCountlist);
        // console.log(valueCountlist.length);
        

        //端末検証用、アウトプット作成
        if(getRaelDeta[i][12]){
          var outPutsheet = mainsheet.getSheetByName('検証結果');
          const resetsheet = outPutsheet;
          resetFilter(resetsheet);

          var outPut = outPutsheet.getRange('A10:A');
          var newSheetinPut = newsheet.getRange('A10:A');
          

          var outPut1 = outPutsheet.getRange('C10:C');
          var newSheetinPut1 = newsheet.getRange('B10:B');

          var outPut2 = outPutsheet.getRange('F10:H');
          var newSheetinPut2 = newsheet.getRange('M10:O');

          var outPut3 = outPutsheet.getRange('L10:L');
          var newSheetinPut3 = newsheet.getRange('P10:P');

          var outPut3 = outPutsheet.getRange('L10:L');
          var newSheetinPut3 = newsheet.getRange('P10:P');



          //設定したフィルタ名がある列を検索
          let reportTcFilterCol1 = newsheetValues[0].indexOf('機能');
          let reportTcFilterCol2 = newsheetValues[0].indexOf('実施対象');
          let reportTcFilterCol3 = newsheetValues[0].indexOf('端末検証報告');
          let reportTcFilterCol4 = newsheetValues[0].indexOf('記載内容');

          //参考コードを使用
          let col1alf = ExcelUtils.convertNumericColumnToAlphabet(reportTcFilterCol1+1);
          let col2alf = ExcelUtils.convertNumericColumnToAlphabet(reportTcFilterCol2+1);
          let col3alf = ExcelUtils.convertNumericColumnToAlphabet(reportTcFilterCol3+1);
          let col4alf = ExcelUtils.convertNumericColumnToAlphabet(reportTcFilterCol4+1);

          var outPut4 = outPutsheet.getRange('D10:D');
          console.log(col1alf+'10'+':'+col1alf)
          var newSheetinPut4 = newsheet.getRange(col1alf+'10'+':'+col1alf);
          

          var outPut5 = outPutsheet.getRange('E10:E');
          var newSheetinPut5 = newsheet.getRange(col2alf+'10'+':'+col2alf);

          var outPut6 = outPutsheet.getRange('N10:N');
          var newSheetinPut6 = newsheet.getRange(col3alf+'10'+':'+col3alf);

          var outPut7 = outPutsheet.getRange('O10:O');
          var newSheetinPut7 = newsheet.getRange(col4alf+'10'+':'+col4alf);

          outPut.clearContent();
          outPut1.clearContent();
          outPut2.clearContent();
          outPut3.clearContent();
          outPut4.clearContent();
          outPut5.clearContent();
          outPut6.clearContent();
          outPut7.clearContent();


          let putinWordcont = valueCountlist.length;

          newSheetinPut.copyTo(outPut , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          outPutsheet.getRange(11,2,putinWordcont,1).setValues(valueCountlist);
          newSheetinPut1.copyTo(outPut1 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut2.copyTo(outPut2 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut3.copyTo(outPut3 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut4.copyTo(outPut4 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut5.copyTo(outPut5 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut6.copyTo(outPut6 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
          newSheetinPut7.copyTo(outPut7 , SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

          const removeRowCount = outPutsheet.getLastRow()-10-putinWordcont;
          if(removeRowCount >= 1){
            outPutsheet.deleteRows(10+ putinWordcont +1,removeRowCount);
          }

          outPutsheet.getRange('A10:O').setBorder(true, true, true, true, true, true);

          if(getRaelDeta[i][13] == true){
            mainsheet.deleteSheet(newsheet);
          }

        }

        }//<<--if(copySheet){ -->>

      else{
        Browser.msgBox('Cancel','該当のスプレッドシート「'+　getRaelDeta[i][1] +'」には、シート名「'+ getRaelDeta[i][3] +'」がありません。\\n設定したシート名がなければコピーすることはできません。',Browser.Buttons.OK);
        }
    }
  }//ループ処理終了（i=0）

//実行後オプション　試験表コピーシートを削除する
  if(getoptionset[0][1] == true) mainsheet2.deleteSheet(mainsheet1);

}

function testSheetClear(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('検証結果');
  const resetsheet = sheet;
  resetFilter(resetsheet);
  const clerRange = sheet.getRange('A11:O');
  clerRange.clearContent();
  let inputVelu = [[]];
  for(i=0;i<200;++i){
    inputVelu[i] = [i+1]
  }
  sheet.getRange(11,2,inputVelu.length,1).setValues(inputVelu);
  const sheetlastRow = sheet.getMaxRows();
  let removeCount = sheetlastRow-10-inputVelu.length;
  console.log(10+inputVelu.length+1);
  console.log(sheetlastRow);
  console.log(removeCount);
  if(removeCount >= 1){
    sheet.deleteRows(10+inputVelu.length+1,removeCount);
  }
  const clerRange1 = sheet.getRange('A11:O');
  clerRange1.setBorder(true, true, true, true, true, true);

}


//マスター試験表の情報を取得
function getMastersheetlist() {

     /* 基準フォルダから下のフォルダ階層をすべてリストアップする */
      var folderId = "11eYkhvrmYZy2hR9y7-i3Lt2CoAM7yA6Z";

      let key = folderId;
      let stt = DriveApp.getFolderById(key);
      let name = "";
      let i = 0;
      let j = 0;
      let folderlist = [];
      let result = [];

      // 開始フォルダ
      folderlist.push([stt, key]);
      name = stt + '/';


    do {
      //フォルダ一覧取得
      let folders = DriveApp.searchFolders("'"+key+"' in parents");
      //フォルダ一覧からフォルダ名称とIdを取得
      while(folders.hasNext()){
        i++;
        let folder = folders.next();
        let tmparray = new Array();
        tmparray.push(name + folder.getName());
        tmparray.push(folder.getId());
        folderlist.push(tmparray);
      }
      //フォルダ名称とIdを取り出す
      j++;
      if(j <= i){
        name = folderlist[j][0] + '/';
        key = folderlist[j][1];
        }
      } while (j <= i);
    
    let fileArray = [];
    //ファイル情報取得
    for (i=0; i<=folderlist.length-1; i++) {
      let key = DriveApp.getFolderById(folderlist[i][1]).getId();
      let files = DriveApp.searchFiles("starred = true and '"+key+"' in parents");
      while(files.hasNext()){
        fileArray = [];
        let file = files.next();
        // fileArray.push(folderlist[i][0]);
        fileArray.push(file.getName());
        let materSheetUrl = file.getUrl();
        fileArray.push(materSheetUrl);
        result.push(fileArray);
        j++;
        }
      }
    // console.log(folderlist);
    console.log(result);

    const inputAia = SpreadsheetApp.getActiveSheet().getRange(11,20,result.length,2);

    inputAia.setValues(result);
}

















