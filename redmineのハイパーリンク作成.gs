///選択範囲のシート内のデータを配列に格納、リンクから特定の文字を削除し,選択中のセルにハイパーリンク文字列を入れる
function PutLink_redmine() {
  //スプレッドシートのシートオブジェクトを取得
  let  sheet=SpreadsheetApp.getActiveSheet();
  //選択されているアクティブなセルを取得
  let  sel=sheet.getActiveRange();
　// そのセル範囲にある値の多次元配列を取得
  let values = sel.getValues();
  // ハイパーリンク文字列の配列
  let linkList = [[]];

  for(let i=0; i<values.length; i++) {
    // シート内のデータ
    let value = values[i];  
    // シートのURLからハイパーリンク文字列を組み立て
    let url = "https://172.16.126.9/redmine/issues/" + value;
    //リンクから特定の文字を置換
    let url_replace1 = url.replace("＃","#"); 
    //リンクから特定の文字を削除
    let url_replace2 = url_replace1.replace("#",""); 
    let link1 = '=HYPERLINK("' + url_replace2 + '","' + value + '")';
    //特定の文字を置換
    let link2 = link1.replace("＃","#"); 
    let link3 = [ link2 ];
    // ハイパーリンク文字列を配列に格納
    linkList[i] = link3;
  }

  // 選択中のセルにハイパーリンク文字列を入れる
  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cell1 = sheet1.getActiveCell();
  let range = sheet1.getRange(cell1.getRow() , cell1.getColumn() ,  linkList.length , 1);
  range.setValues(linkList);
}
