function deleteSteptNo1() {
  //スプレッドシートのシートオブジェクトを取得
  let  sheet=SpreadsheetApp.getActiveSheet();
  //選択されているアクティブなセルを取得
  let  sel=sheet.getActiveRange();
　// そのセル範囲にある値の多次元配列を取得
  let values = sel.getValues();
  console.log(values);

  let valuesist = [[]];

for(let k=0; k<values.length; k++){
    let value = values[k][0];
    console.log(value);
    let value0 = value.replace(/\. \[/gi, '.[');//https://qiita.com/katsukii/items/1c1550f064b4686c04d4 //https://www.acrovision.jp/service/gas/?p=602
    let value01 = value0.replace(/　\.\[/gi, '.[')  
    let value02 = value01.replace(/\] /gi, ']')  
    let value03 = value02.replace(/\]/gi, '] ')  
    let value04 =  value03.replace(/[1-9]\.\[/gi,".[");
    console.log(value04);
    let value1 = value04.split('.[');
    console.log(value1);
    console.log(value1.length);

    let valueList = [[]];
    for(let i=0; i<value1.length; i++) {
      if(value1 !="手順"){
      if(i < value1.length-1){
        
        let value2 =  value1[i]+[i+1];
        let value3 = [ value2 ];
        valueList[i] = value3;
      }else{
        let value2 =  value1[i];
        let value3 = [ value2 ];
        valueList[i] = value3;

      }
      }

    }
    let allOffices = valueList.join('.[');
    console.log(allOffices);
    valuesist[k] = [allOffices];

}


    console.log(valuesist);


  let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cell1 = sheet1.getActiveCell();
  let range = sheet1.getRange(cell1.getRow() , cell1.getColumn() ,  values.length , 1);
  range.setValues(valuesist);
}
