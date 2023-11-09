function cellmerge(){
  const selectRange = SpreadsheetApp.getActiveRange();
  const selectValues = selectRange.getDisplayValues();
  var regRow = /\\[1-9]\. *|\\[1-9][0-9]\. */
  var regCol = /\/[1-9]\. *|\/[1-9][0-9]\. */

  if(selectValues.length > 1 && selectValues[0].length > 1){
    Browser.msgBox('縦2以上、かつ、横2以上を結合することはできません。\\n redmein の表は、行のみ、または、列のみを結合することができます。');
    return;
  }
  console.log(selectValues);
  console.log(selectValues.length);
  console.log(selectValues[0][0]);
  console.log(selectValues[0].length);
  selectRange.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
  const reselectValues = selectRange.getDisplayValues();
  if(regRow.test(reselectValues[0][0])){
    if(reselectValues[0][0].match(/\\([1-9])\. */) != null ){
      selectRange.setValue('\\'　+　selectValues[0].length + reselectValues[0][0].slice(2));
    }
    else if(reselectValues[0][0].match(/\\([1-9][0-9])\. */) != null ){
      selectRange.setValue('\\'　+　selectValues[0].length + reselectValues[0][0].slice(3));
    }
  }
  else if(regCol.test(reselectValues[0][0])){
    if(reselectValues[0][0].match(/\/([1-9])\. */) != null ){
      selectRange.setValue('/'　+　selectValues.length + reselectValues[0][0].slice(2));
    }
    else if(reselectValues[0][0].match(/\/([1-9][0-9])\. */) != null ){
      selectRange.setValue('/'　+　selectValues.length + reselectValues[0][0].slice(3));
    }
  }
  else if(selectValues[0].length > 1 ){
    selectRange.setValue('\\'　+　selectValues[0].length + '. '+reselectValues[0][0])
  }
  else{
    selectRange.setValue('/'　+　selectValues.length + '. '+reselectValues[0][0])
  }
}


function cellBreakApart(){
  const selectRange = SpreadsheetApp.getActiveRange();
  // console.log(selectRange);
  const selectValues = selectRange.getDisplayValues();
  var reg = /\/[1-9]\. *|\/[1-9][<>=]\. *|\\[1-9]\. *|\\[1-9][<>=]\. *|\/[1-9][0-9]\. *|\/[1-9][0-9][<>=]\. *|\\[1-9][0-9]\. *|\\[1-9][0-9][<>=]\. */
  // if(selectValues.length > 1){
  //   Browser.msgBox('複数行を結合することはできません。\\n redmein の表は一行のみを結合することができます。');
  //   return;
  // }
  var reselectValues = selectRange.getDisplayValues();
  var newWords = [];

  if(selectValues.length > 1 || selectValues[0].length > 1){
    selectRange.breakApart().setVerticalAlignment("middle").setHorizontalAlignment("left");
    var reselectValues = selectRange.getDisplayValues();
    console.log(reselectValues)

    for(i=0;i<reselectValues.length;++i){
      var newWord = [];
      for(k=0;k<reselectValues[i].length; ++k){
        console.log('reselectValues[i][k] :'+reselectValues[i][k]);
        console.log('newWords :'+newWords);

        if(reselectValues[i][k].match(/\\[1-9]\. */) != null | reselectValues[i][k].match(/\/[1-9]\. */) != null ){
          
          newWord.push(reselectValues[i][k].slice(4));
        }
        else if(reselectValues[i][k].match(/\\[1-9][0-9]\. */) != null | reselectValues[i][k].match(/\/[1-9][0-9]\. */) != null ){
          
          newWord.push(reselectValues[i][k].slice(5));
        }
        else{
          newWord.push(reselectValues[i][k]);
        }
     }
     newWords.push(newWord);
    }
    selectRange.setValues(newWords);
    return;
  }

  if(reg.test(reselectValues[0][0])){
    if(reselectValues[0][0].match(/\\([1-9])\. */) != null | reselectValues[0][0].match(/\/([1-9])\. */) != null ){
      // const selectOneRange = SpreadsheetApp.getActiveRange().getRan
      selectRange.setValue(reselectValues[0][0].slice(4));
    }
    else if(reselectValues[0][0].match(/\\([1-9][0-9])\. */) != null | reselectValues[0][0].match(/\/([1-9][0-9])\. */) != null ){
      // const selectOneRange = SpreadsheetApp.getActive().getRange();
      selectRange.setValue(reselectValues[0][0].slice(5));
    }
  }
  selectRange.breakApart().setVerticalAlignment("middle").setHorizontalAlignment("left");

}




function createExit() {
  　var selectRange = SpreadsheetApp.getActive().getActiveRange();
    var values = selectRange.getDisplayValues();
    var results = selectRange.getHorizontalAlignments();
    var resultStyles = selectRange.getTextStyles();

    // const a = values[0][0].slice(0, 3)
    // const b = '='
    // const c = values[0][0].slice(3)
    // console.log(values);
    console.log(values[0][0]);
    console.log(selectRange.getTextStyles());
    // console.log(selectRange.getTextStyles()[0][0].isBold());
    // console.log(selectRange.getTextStyles()[0][1].isBold());
    // console.log(resultStyles[0][0].isItalic());
    // console.log(resultStyles[0][1].isItalic());
    console.log(selectRange.getTextStyle().isBold());
    // console.log(a + b + c);
    // console.log(values[0][0].slice(0, 3) + '=' + values[0][0].slice(3));
    // console.log(results);

  console.log(values);
  // var reg = /\/[1-9]\. *|\\[1-9]\. *|\/[1-9][0-9]\. *|\\[1-9][0-9]\. */
  var reg = /[\/\\][1-9]\. *|[\/\\][1-9][<>_=]\. *|[\/\\][1-9][0-9]\. *|[\/\\][1-9][0-9][<>_=]\. */

  var newValues = [];

  console.log(values.length);
  console.log(values[0].length);
    for(k=0;k<values.length;k++){
      for(i=0; i<values[k].length; i++){
        // console.log('k:'+k+' '+'i:'+i);
        // console.log('k:'+k+' '+'i:'+i+' '+resultStyles[k][i].isItalic());
        console.log(values);
        // if(resultStyles[k][i].isItalic() == true){
        //   if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
        //     values[k].splice( i, 1, values[k][i].slice(0, 4) + '_' + values[k][i].slice(4) + '_');
        //     }
        //   else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
        //     values[k].splice( i, 1, values[k][i].slice(0, 5) + '_' + values[k][i].slice(5) + '_');
        //     }
        //   else{
        //     values[k].splice( i, 1, '_' + values[k][i] + '_');
        //     }
        //   }
        if(resultStyles[k][i].isUnderline() == true){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '+' + values[k][i].slice(4) + '+');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '+' + values[k][i].slice(5) + '+');
            }
          else{
            values[k].splice( i, 1, '+' + values[k][i] + '+');
            }
          }
        if(resultStyles[k][i].getForegroundColor() == '#ff0000'){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '%{color:#ff0000}' + values[k][i].slice(4) + '%');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '%{color:#ff0000}' + values[k][i].slice(5) + '%');
            }
          else{
            values[k].splice( i, 1, '%{color:#ff0000}' + values[k][i] + '%');
            }
          }
        if(resultStyles[k][i].getForegroundColor() == '#0000ff'){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '%{color:#0000ff}' + values[k][i].slice(4) + '%');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '%{color:#0000ff}' + values[k][i].slice(5) + '%');
            }
          else{
            values[k].splice( i, 1, '%{color:#0000ff}' + values[k][i] + '%');
            }
          }
        if(resultStyles[k][i].getForegroundColor() == '#008000'){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '%{color:#008000}' + values[k][i].slice(4) + '%');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '%{color:#008000}' + values[k][i].slice(5) + '%');
            }
          else{
            values[k].splice( i, 1, '%{color:#008000}' + values[k][i] + '%');
            }
          }
        if(resultStyles[k][i].isBold() == true){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '*' + values[k][i].slice(4) + '*');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '*' + values[k][i].slice(5) + '*');
            }
          else{
            values[k].splice( i, 1, '*' + values[k][i] + '*');
            }
          }
        if(resultStyles[k][i].isStrikethrough() == true){
          if(values[k][i].match(/[\/\\][1-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 4) + '-' + values[k][i].slice(4) + '-');
            }
          else if(values[k][i].match(/[\/\\][1-9][0-9]\. */) != null ){
            values[k].splice( i, 1, values[k][i].slice(0, 5) + '-' + values[k][i].slice(5) + '-');
            }
          else{
            values[k].splice( i, 1, '-' + values[k][i] + '-');
            }
          }
      }
    }


    for(k=0;k<values.length;k++){
      var newValue = [];
      // console.log(values[k]);
      for(var i=0; i<values[k].length; i++){
        if(results[k][i] == 'general-center' || results[k][i] == 'center' && reg.test(values[k][i]) == false && resultStyles[k][i].isBold() == true){
          if(values[k][i].match(/\\([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '_' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\\([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '_' + values[k][i].slice(3));
            }
          else if(values[k][i].match(/\/([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '_' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\/([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '_' + values[k][i].slice(3));
            }
          else{
            newValue.push('_. ' + values[k][i]);
            }
          }
        else if(results[k][i] == 'general-center' || results[k][i] == 'center' && reg.test(values[k][i]) == true && resultStyles[k][i].isBold() == true){
          newValue.push('_' + values[k][i]);
          }
        else if(results[k][i] == 'general-center' || results[k][i] == 'center'){
          if(values[k][i].match(/\\([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '=' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\\([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '=' + values[k][i].slice(3));
            }
          else if(values[k][i].match(/\/([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '=' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\/([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '=' + values[k][i].slice(3));
            }
          else{
            newValue.push('=. ' + values[k][i]);
            }
          }
        else if(results[k][i] == 'general-left' || results[k][i] == 'left'){
          if(values[k][i].match(/\\([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '<' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\\([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '<' + values[k][i].slice(3));
            }
          else if(values[k][i].match(/\/([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '<' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\/([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '<' + values[k][i].slice(3));
            }
          else{
            newValue.push('<. ' + values[k][i]);
            }
          }
        else if(results[k][i] == 'general-right' || results[k][i] == 'right'){
          // newValue.push('>. ' + values[k][i]);
          if(values[k][i].match(/\\([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '>' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\\([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '>' + values[k][i].slice(3));
            }
          else if(values[k][i].match(/\/([1-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 2) + '>' + values[k][i].slice(2));
            }
          else if(values[k][i].match(/\/([1-9][0-9])\. */) != null ){
            newValue.push(values[k][i].slice(0, 3) + '>' + values[k][i].slice(3));
            }
          else{
            newValue.push('>. ' + values[k][i]);
            }
          }  
        else{
            newValue.push(values[k][i]);
          }
        
        // var reg = new RegExp(values[k][i]);
        if(reg.test(values[k][i])){

          if(values[k][i].match(/\\([1-9])\. */) != null ){
            var test = values[k][i].match(/\\([1-9])\. */);
            console.log(test);
            console.log(test[1]);
            i += test[1]-1;
            }

          if(values[k][i].match(/\\([1-9][0-9])\. */) != null ){
            var test = values[k][i].match(/\\([1-9][0-9])\. */);
            console.log(test);
            console.log(test[1]);
            i += test[1]-1;
            }


          }
        }
      newValues.push(newValue);
    }
  console.log(newValues);

     for(k=0;k<newValues.length;k++){

      for(var i=0; i<newValues[k].length; i++){

        if(reg.test(newValues[k][i])){

          if(newValues[k][i].match(/\/([1-9])\. */) != null ){
            var test = newValues[k][i].match(/\/([1-9])\. */);
            for(n=1;n<test[1];++n){
              // newValues[k+n].splice(i,1);
              newValues[k+n].splice( i, 1, '＊未処理＊');
              }
            }
          if(newValues[k][i].match(/\/([1-9])[<>=]\. */) != null ){
            var test = newValues[k][i].match(/\/([1-9])[<>=]\. */);
            for(n=1;n<test[1];++n){
              // newValues[k+n].splice(i,1);
              newValues[k+n].splice( i, 1, '＊未処理＊');
              }
            }
          if(newValues[k][i].match(/\/([1-9][0-9])\. */) != null ){
            var test = newValues[k][i].match(/\/([1-9][0-9])\. */);
            for(n=1;n<test[1];++n){
              // newValues[k+n].splice(i,1);
              newValues[k+n].splice( i, 1, '＊未処理＊');
              }
            }
          if(newValues[k][i].match(/\/([1-9][0-9])[<>=]\. */) != null ){
            var test = newValues[k][i].match(/\/([1-9][0-9])[<>=]\. */);
            for(n=1;n<test[1];++n){
              // newValues[k+n].splice(i,1);
              newValues[k+n].splice( i, 1, '＊未処理＊');
              }
            }

          }
        }
      
    }
  console.log(newValues);

  let body = '';
  for(k=0;k<newValues.length;k++){
    body += '|';
    for(i=0; i<newValues[k].length; i++){
      // var exiteValue = newValues[k][i].replace(/\n/g,'<br>');
      if(newValues[k][i] =='＊未処理＊'){
        // body += '|';
      }else{
        body += newValues[k][i] + '|';
      }
  }
  body += '\n';

  }
  // Logger.log(body);
  return body;
  // return showWindow(body);
  // console.log(body);

  // var ui = SpreadsheetApp.getUi();
  // //HTMLダイアログを表示する
  // //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  // //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です
  // var output = HtmlService.createTemplateFromFile('index-報告書')
  // output.inputword = body;

  // var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
  //            //https://wakky.tech/gas-html-title/
  //           //  .setTitle('タイトルのテスト')
  //            .setWidth(700)
  //            .setHeight(700);
  // ui.showModelessDialog(html,'redmine用　表');
  
}
function showWindow(){
  var body = createExit();

  var ui = SpreadsheetApp.getUi();
  //HTMLダイアログを表示する
  //ui.showModalDialogで表示すればモーダルダイアログになります。ダイアログ以外の操作を一時的にロックする状態になります。
  //ui.showModelessDialogであればモードレスダイアログになります。ダイアログ以外の操作も可能です
  var output = HtmlService.createTemplateFromFile('index-redmine')
  output.inputword = body;

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
             //https://wakky.tech/gas-html-title/
            //  .setTitle('タイトルのテスト')
             .setWidth(700)
             .setHeight(700);
  ui.showModelessDialog(html,'redmine用　表');
  
}

function showSidebar() {
  var body = createExit();
  var output = HtmlService.createTemplateFromFile('index-redmine')
  output.inputword = body;
  output.inputsub = 'showSidebar';

  var html = output.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('サイドバー');

  SpreadsheetApp.getUi().showSidebar(html);  
}










