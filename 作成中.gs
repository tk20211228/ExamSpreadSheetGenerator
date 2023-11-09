//別のスプレッドシートにコピーできます。
function copSp() {
  //スクリプトに紐付いたアクティブなシートをコピー対象のシートとして読み込む
  let copySheet = SpreadsheetApp.getActiveSheet();
     // ID からスプレッドシートを取得
     // var destination = SpreadsheetApp.openById('XXXX');
  // URL からスプレッドシートを開く
  let destSpreadsheet = SpreadsheetApp.openByUrl('https://script.google.com/home/projects/1W2vRfOa09pVN1lRHgBAXnvJFgyhMXFWeHswecQtsPZ1YiRoF8pxUB-ys/edit');
  //コピー対象シートを同一のスプレッドシートにコピー
  //シート名を設定しないとコピー先で「～～のコピー」と表示されます。
  let newCopySheet = copySheet.copyTo(destSpreadsheet);
  //コピーしたシート名を変更する。
  newCopySheet.setName("******")
};


//フィルター操作
function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('P8').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['音楽'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(16, criteria);
};

  //--------------------------------------------------------------------------------------【メッセージボックス】-----------------------------------------------------------------------------------
//Google Apps Scriptでダイアログボックスやサイドバーを作る
function normalmsg(){

    //スプレッドシートのみで使える
  Browser.msgBox("これは、スプレッドシートでしか使えない。 \\n 改行テスト");

  //通常のメッセージボックス
  //メッセージを表示
  //今回はスプレッドシートのケースですが、フォームならば、FormApp.getUi()で始めます。
  //getUiでUIを取得させて、ui.alertで表示する仕組みです。
  var ui = SpreadsheetApp.getUi();
  ui.alert("これは、どんなアプリでも使えるメッセージ。Browser.msgboxと見た目は変わらない。 \\n しかし改行が使えないかも");

  //インプットボックス
  //ユーザに値を入力させる為のボックス
  var ret = ui.prompt("テキストボックスに何かいれて下さい \\n 改行テスト", 
      ui.ButtonSet.OK_CANCEL);
   //押されたボタンによって処理を分岐
   switch(ret.getSelectedButton()){
    //OKボタンを押した時の処理
    case ui.Button.OK:
      var str = ret.getResponseText();
      ui.alert(str);
      break;

    //キャンセルを押した時の処理
    case ui.Button.CANCEL:
      ui.alert("何もせずに閉じました。");
      break;

    case ui.Button.CLOSE:
      break;

  }
 
  //問合せボックス
  //処理開始前の確認などで使います。
  var re = ui.alert("自爆スイッチ", "自爆してみますか？", ui.ButtonSet.YES_NO_CANCEL);
  switch(re){
    case ui.Button.YES:
      ui.alert("命は大切にしましょう。");
      return 0;
      break;
    case ui.Button.NO:
      ui.alert("自爆を中止しました。");
      return 0;
      break;
    case ui.Button.CANCEL:
      ui.alert("自爆をキャンセルしました。");
      return 0;
     break;

  } 
}

//--------------------------------------------------------------------------------------【ダイアログボックス】-----------------------------------------------------------------------------------
//HTML Serviceを使ったダイアログボックスです。HTMLで作成しますので、ありとあらゆるタイプの豪華なダイアログボックスを作成する事が可能


function htmlbox() {
  /*
  //GAS側コード(スプレッドシートのみ）
  //今回はcreateTemplateFromFileを使ってますが、createHtmlOutputFromFileを使っても問題ありません。
  　var output = HtmlService.createTemplateFromFile('index');
  　var ss = SpreadsheetApp.getActiveSpreadsheet();
  　var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             .setWidth(500)
             .setHeight(600);
  　ss.show(html);    //メッセージボックスとしてを表示する
  */


  //GAS側コード（全アプリ共通）
  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(500).setHeight(600);
  ui.showModalDialog(html, "タイトル");
  　
}

//--------------------------------------------------------------------------------------【サイドバー】-----------------------------------------------------------------------------------
//ダイアログボックスと異なりユーザの画面上ではなく右サイドに出てきます。作業の邪魔をしないので、設定関係であったりスプレッドシートのデータに対して作業するようなツールボックス的な利用が主流
//入力用コンソールをsidebarで表示する
function openSidebar() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile("index").evaluate()
                    .setTitle('データ作成').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  ui.showSidebar(html);
}





















































