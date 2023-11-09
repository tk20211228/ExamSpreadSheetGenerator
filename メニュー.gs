
//参考URLhttps://www.gesource.jp/weblog/?p=8140
function onOpen(e) {
  var meinUI = SpreadsheetApp.getUi();
    meinUI
        .createMenu('項目作成・修正ツール')
        .addSubMenu(meinUI.createMenu('メンテナンスツール')
            .addItem('単数シート削除', 'deleteSheet')
            .addItem('複数シート削除', 'deleteSheets')
            .addItem('選択したセルにシート一覧を作成(リンク有)', 'PutLinks')
            .addSeparator()
            .addItem('選択範囲の位置を取得', 'mygetRowcolumnActiveRange')
            .addItem('getRangeで使用できる選択範囲の位置','mygetRowcolumnActiveRange0530'))
        .addSeparator()
        .addItem('redmineのハイパーリンク作成', 'PutLink_redmine')
        .addSeparator()
        .addSubMenu(meinUI.createMenu('フィルタ')//resetFilter
            .addItem('フィルタリセット', 'resetFilter')
            .addItem('フィルタ生成', 'createFilter')
            .addItem('フィルタ削除', 'removeFilter')
            .addSeparator()
            .addItem('データフィルタ一括削除', 'removeDetaFiler')
            .addItem('データフィルタ名に「フィルタ」の含むのみ削除', 'deleteidentificationDeteFilter'))
        .addSeparator()
        .addSubMenu(meinUI.createMenu('項目修正ツール')
            .addItem('手順修正', 'deleteSteptNo1')
            .addItem('期待動作修正', 'expectedBehaviorNo')
            .addSeparator()
            .addItem('手順番号のみ入力', 'stepNoPutin')
            .addItem('手順番号のみ削除', 'deleteSteptNo'))
        .addSeparator()
        .addSubMenu(meinUI.createMenu('項目作成フォーマット')
            .addItem('評価観点、事前準備、手順、期待動作', 'allPutin')
            .addItem('評価観点', 'viewpointPutin')
            .addItem('事前準備', 'advancePreparationPutin')//
            .addItem('手順', 'stepFormatPutin')
            .addItem('期待動作', 'expectedBehaviorPutin'))
        .addSubMenu(meinUI.createMenu('検証memoフォーマット')
            .addItem('端末検証の注意書き', 'testMemo')
            .addItem('削除理由（新規開発関連）', 'testMemo1')
            .addItem('削除理由（観点重複）', 'testMemo2')
            .addItem('既存バグの注意書き', 'testMemo3'))
        .addSeparator()
        .addSubMenu(meinUI.createMenu('プルダウン作成')
            .addItem('シート一覧', 'sheetList')
            .addItem('試験対象、フェルター', 'testSubjectList')
            .addItem('結果（OK,NG,NA…）', 'resultSimpleList')
            .addItem('redmine(#XXXX,#ZZZZ…)', 'redmineList')
            .addItem('試験時間', 'testtimeList')
            .addItem('試験環境', 'testplaceList')
            .addItem('OS Ver. (Console)', 'testPCosverList')
            .addSeparator()
            .addSubMenu(meinUI.createMenu('OS Ver.(検証端末)')
                .addItem('全て', 'testAllOSverList')
                .addItem('Android', 'testAndroidverList')
                .addItem('iOS', 'iOSverList')
                .addItem('iPadOS', 'iPadOSverList')
                .addItem('windows', 'windowsVerList')
                .addItem('macOS', 'macOSVerList'))          
            .addSeparator()
            .addSubMenu(meinUI.createMenu('カテゴリ（簡易型）')
                .addItem('機能カテゴリ', 'createSimpleList')
                .addItem('大項目', 'createSimpleList1')
                .addItem('中項目', 'createSimpleList2')
                .addItem('小項目', 'createSimpleList3')
                .addItem('機能カテゴリ、大項目、中項目、小項目', 'createSimpleListAll')) 
            .addSubMenu(meinUI.createMenu('カテゴリ（連動型）')
                .addItem('機能カテゴリ', 'createLinkList')
                .addItem('大項目', 'createLinkList1')
                .addItem('中項目', 'createLinkList2')
                .addItem('小項目', 'createLinkList3')
                .addItem('機能カテゴリ、大項目、中項目、小項目', 'createLinkListAll'))
            .addSeparator()
            .addSubMenu(meinUI.createMenu('項目修正シート')
                .addItem('修正箇所', 'modificationPlaceList')
                .addItem('マスタ―（対応状況）', 'correspondenceSituationList')
                .addItem('修正者', 'recordmenList'))              
            )
        .addSubMenu(meinUI.createMenu('プルダウン削除')
            .addItem('選択部分', 'selectListClear')
            .addItem('結果部分', 'resultListClear')
            .addItem('カテゴリ部分', 'categoryListClear')
            .addItem('redmine部分', 'redmineListClear')
            .addSeparator()
            .addSubMenu(meinUI.createMenu('全体削除')
                .addItem('シート', 'sheetListClear')
                .addItem('スプレッドシート全体', 'spreadsheetListClear'))     
            )
        .addToUi();
    meinUI
        .createMenu('URL一覧')
        .addItem('検証参考_URL一覧', 'displayUrls')
        .addItem('サポートサイト', 'testUrl')
        .addItem('サポートサイト ビジプラ', 'testUrl0')
        .addItem('テストサイト', 'testUrl1')
        .addItem('ISB', 'testUrl2')
        .addItem('アカウント一覧', 'testUrl3')
        .addItem('Slack', 'testUrl4')
        .addItem('redmien', 'testUrl5')
        .addSeparator()
        .addSubMenu(meinUI.createMenu('console　DL')
            .addItem('本番', 'testUrl6')
            .addItem('PB', 'testUrl7')
            .addItem('B2', 'testUrl8')
            .addItem('B', 'testUrl9')
            .addItem('Az', 'testUrl10'))
        .addSubMenu(meinUI.createMenu('DES　DL')
            .addItem('本番', 'testUrl11')
            .addItem('PB', 'testUrl12')
            .addItem('B2', 'testUrl13')//
            .addItem('B', 'testUrl14')
            .addItem('Az', 'testUrl15'))
        .addSubMenu(meinUI.createMenu('Web console')
            .addItem('本番', 'testUrl36')
            .addItem('PB', 'testUrl37')
            .addItem('B2', 'testUrl38')//
            .addItem('B', 'testUrl39')
            .addItem('Az', 'testUrl40'))
        .addSeparator()
        .addSubMenu(meinUI.createMenu('Agent')
            .addSubMenu(meinUI.createMenu('Android')
                .addItem('本番', 'testUrl16')
                .addItem('PB', 'testUrl17')
                .addItem('B2', 'testUrl18')
                .addItem('B', 'testUrl19')
                .addItem('Az', 'testUrl20'))
            .addSubMenu(meinUI.createMenu('iOS(iDEP)')
                .addItem('本番', 'testUrl21')
                .addItem('PB', 'testUrl22')
                .addItem('B2', 'testUrl23')
                .addItem('B', 'testUrl24')
                .addItem('Az', 'testUrl25'))
            .addSubMenu(meinUI.createMenu('windows')
                .addItem('本番', 'testUrl26')
                .addItem('PB', 'testUrl27')
                .addItem('B2', 'testUrl28')
                .addItem('B', 'testUrl29')
                .addItem('Az', 'testUrl30'))
            .addSubMenu(meinUI.createMenu('macOS')
                .addItem('本番', 'testUr31')
                .addItem('PB', 'testUrl32')
                .addItem('B2', 'testUrl33')
                .addItem('B', 'testUrl34')
                .addItem('Az', 'testUrl35'))
            )
        .addToUi();
    meinUI
        .createMenu('試験表取得・コピー')
        .addItem('開いているシートをコピーする', 'copysheet')
        .addItem('Googleドライブからマスターを取得する', 'copyMastersheet')
        .addItem('URLからシートを全て取得する', 'URLcopysheet')

        
        .addToUi();
    meinUI
        .createMenu('報告書作成ツール')
        .addItem('NG内訳、NA内訳など出力', 'result')
        .addItem('通常用　作成', 'resultreport')
        .addItem('Slack用　作成', 'Slackresultreport')
        .addItem('redmine用　作成', 'redmienresultreport')
        .addItem('Slack用　報告用スレッド', 'SlackFormat')
        .addToUi();
    meinUI
        .createMenu('redmine')
        .addItem('ウィンドウ（表出力）', 'showWindow')
        .addItem('サイドバー（表出力）', 'showSidebar')
        .addItem('セル結合[ 水平配置（中央）、垂直配置（中央）]', 'cellmerge')
        .addItem('セル結合解除[ 水平配置（左）、垂直配置（中央）]', 'cellBreakApart')
        .addToUi();
    meinUI
        .createMenu('端末検証')
        .addItem('報告書に試験結果入力', 'TerminalInspectionReport')
        .addItem('動作確認済み一覧用の試験結果出力', 'motionConfirmationlist')
        .addItem('Excelファイル出力', 'createExlsxfile')//testSheetClear
        .addItem('「検証結果」シートをリセット', 'testSheetClear')//

        .addToUi();
  //自動シート名一覧を取得
  var sheets = SpreadsheetApp.getActive().getSheets();
  var ssId = SpreadsheetApp.getActive().getId();
  // ハイパーリンク文字列の配列
  var linkList = [[]];
  for(var i=0; i<sheets.length; i++) {
      var sheetId = sheets[i].getSheetId();
      // シートのIDと名前
      var sheetName = sheets[i].getSheetName();

     // シートのURLからハイパーリンク文字列を組み立て
     var url = "https://docs.google.com/spreadsheets/d/" + ssId + "/edit#gid=" + sheetId;
     var link = [ '=HYPERLINK("' + url + '","' + sheetName + '")' ];
   // ハイパーリンク文字列を配列に格納
    linkList[i] = link;
   }

  // セルにハイパーリンク文字列を入れる
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("集計シート");
  sheet1.getRange('AM11:AM').clearContent();
  var range = sheet1.getRange(11,39,linkList.length,1);
  range.setValues(linkList);
}

function deleteSheet(){
  // var mainsheet0 = SpreadsheetApp.getActiveSheet();
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const getActiveRange = mainsheet.getActiveRange();
  const selectValue = getActiveRange.getValues();
  let sht = mainsheet.getSheetByName(selectValue[0]);
  if(sht){
  mainsheet.deleteSheet(sht)
  // getActiveRange.setValue(selectValue[0]);
  getActiveRange.clearContent();

    }else{
    Browser.msgBox('削除するシートがありません');
    }
}

function deleteSheets(){
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const getActiveRange = mainsheet.getActiveRange();
  const selectValue = getActiveRange.getValues();

    for(i=0 ;i<selectValue.length ;i++){
      var sht1 = mainsheet.getSheetByName(selectValue[i]);
      if(sht1){
        mainsheet.deleteSheet(sht1)
        }
      }
    //選択したセルに、選択したセルの値（表示されている文字列）を上書き
    //ハイパーリンクを削除しています。
    // getActiveRange.setValues(selectValue);
    getActiveRange.clearContent();

}

function displayUrls() {
  var html = HtmlService.createHtmlOutput('<ul>');
  var sheet = SpreadsheetApp.getActive().getSheetByName('setting');
  var lastRow = sheet.getLastRow();
  var names = sheet.getRange(2, 1, lastRow, 1).getValues();
  var urls = sheet.getRange(2, 2, lastRow, 1).getValues();
  for (var i = 0; i < lastRow-1; i++) {
    var name = names[i][0];
    var url = urls[i][0];
    html.append('<li><a href="' + url + '" target="_blank">' + name + '</a></li>');
  }
  html.append('</ul>');
  SpreadsheetApp.getUi().showSidebar(html);
}

