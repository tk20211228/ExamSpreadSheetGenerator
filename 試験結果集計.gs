function test(){  
    const startTime = new Date();
    result();
    //実行終了
    const endTime = new Date();
    //実行時間の計算
    const runTime = (endTime - startTime) / 1000;
    console.log(runTime); 
  }

function result() {
  const meinsheet1 = SpreadsheetApp.getActiveSpreadsheet();
  const meinsheet2 = SpreadsheetApp.getActiveSheet();
  // const meinsheet3 = SpreadsheetApp.getActive();
  // const meinsheetID = meinsheet3.getId();
  const selctsheetNeme = meinsheet2.getRange(11,2,33,1).getValues();
  // console.log(selctsheetNeme);
  // console.log(selctsheetNeme[0]);

  // let NGresult = [];
  let redminelist = [];
  let NAresult = [];
  let Outresult = [];
  let outcontents = [];
  let Keepresult = [];
  let keepcontents = [];
  let OKresult = [];
  for(k=0; k<selctsheetNeme.length; ++k){
    if(selctsheetNeme[k] != ''){

      let getResultsheet =  meinsheet1.getSheetByName(selctsheetNeme[k]);
      // console.log(getResultsheet.getLastRow());
      let getresultvalues = getResultsheet.getRange(11,17,getResultsheet.getLastRow()-10,3).getValues();
      let getItems = getResultsheet.getRange(11,4,getResultsheet.getLastRow()-10,1).getValues();
      // console.log(getresultvalues[2]);
          for(i = 0; i<getresultvalues.length; i++){
            if(getresultvalues[i][0] == 'NG'){
              // NGresult.push(getresultvalues[i]);
              redminelist.push(getresultvalues[i][1]);
            }else if(getresultvalues[i][0] == 'N/A'){
              NAresult.push(getItems[i] + '***joinword***' + getresultvalues[i][2]);
            }else if(getresultvalues[i][0] == '試験対象外'){
              Outresult.push(getItems[i] + '***joinword***' + getresultvalues[i][2]);
              outcontents.push(getresultvalues[i][2]);
            }else if(getresultvalues[i][0] == '保留'){
              Keepresult.push(getItems[i] + '***joinword***' + getresultvalues[i][2]);
              keepcontents.push(getresultvalues[i][2])
            }else if(getresultvalues[i][0] == 'OK'){
              OKresult.push(getresultvalues[i]);
            }
          }
    }
  }
  //----------------------------------------------------------------------【NG集計】----------------------------------------------------------------------
  // console.log(redminelist.length);
  let res = new Set(redminelist);

  let arr = Array.from(res);
  let arrsort = arr.sort();
  // console.log(NAresult);

  //https://qiita.com/saka212/items/408bb17dddefc09004c8
  let count = {};
  for(i=0; i<redminelist.length; ++i){
    var elm = redminelist[i];
    count[elm] = (count[elm] || 0) + 1 ;
  }

  let ExNGnumbarlist = meinsheet1.getRange('AH11:AH').getValues().flat();
  let ExNGtitlelist = meinsheet1.getRange('AI11:AI').getValues().flat();
  let ExNGSituationlist = meinsheet1.getRange('AJ11:AJ').getValues().flat();
  let ExNGNewExistlist = meinsheet1.getRange('AK11:AK').getValues().flat();

  let NGlist = [];

  for(i=0; i<arrsort.length; ++i){
    var list = [];
    list.push(arrsort[i]);
    var redminetitleNo = ExNGnumbarlist.indexOf(arrsort[i]);
    list.push(ExNGtitlelist[redminetitleNo]);
    list.push(count[arrsort[i]]);
    list.push(ExNGSituationlist[redminetitleNo]);
    list.push(ExNGNewExistlist[redminetitleNo]);

    NGlist.push(list);

  }
  // console.log(NGlist);
　//----------------------------------------------------------------------【N/A集計】----------------------------------------------------------------------
  let resNA = new Set(NAresult);
  let arrNA = Array.from(resNA);
  let arrsortNA = arrNA.sort();
  
  let countNA = {};
  for(i=0; i<NAresult.length; ++i){
    var elmNA = NAresult[i];
    countNA[elmNA] = (countNA[elmNA] || 0) + 1 ;
  }
  
  let NAlist = [];
  for(i=0; i<arrsortNA.length; ++i){
    var list = [];
    var naItems_contents = arrsortNA[i].split('***joinword***');
    list.push(naItems_contents[0]+'関連項目');
    list.push(naItems_contents[1]);
    list.push(countNA[arrsortNA[i]]);
    NAlist.push(list);
  }
  //  console.log(NAresult);
  //  console.log(NAlist);
　//----------------------------------------------------------------------【試験対象外集計】----------------------------------------------------------------------
  let resOut = new Set(Outresult);
  let arrOut = Array.from(resOut);
  let arrsortOut = arrOut.sort();
  
  let countOut = {};
  for(i=0; i<Outresult.length; ++i){
    var elmOut = Outresult[i];
    countOut[elmOut] = (countOut[elmOut] || 0) + 1 ;
  }
  
  let Outlist = [];
  for(i=0; i<arrsortOut.length; ++i){
    var list = [];
    var OutItems_contents = arrsortOut[i].split('***joinword***');
    list.push(OutItems_contents[0]+'関連項目');
    list.push(OutItems_contents[1]);
    list.push(countOut[arrsortOut[i]]);
    Outlist.push(list);
  }
  //  console.log(Outresult);
  //  console.log(Outlist);
　//----------------------------------------------------------------------【保留集計】----------------------------------------------------------------------
  let resKeep = new Set(Keepresult);
  let arrKeep = Array.from(resKeep);
  let arrsortKeep = arrKeep.sort();
  
  let countKeep = {};
  for(i=0; i<Keepresult.length; ++i){
    var elmKeep = Keepresult[i];
    countKeep[elmKeep] = (countKeep[elmKeep] || 0) + 1 ;
  }
  
  let Keeplist = [];
  for(i=0; i<arrsortKeep.length; ++i){
    var list = [];
    var KeepItems_contents = arrsortKeep[i].split('***joinword***');
    list.push(KeepItems_contents[0]+'関連項目');
    list.push(KeepItems_contents[1]);
    list.push(countKeep[arrsortKeep[i]]);
    Keeplist.push(list);
  }

　//----------------------------------------------------------------------【データ入力】----------------------------------------------------------------------
  var Totallsheet = meinsheet1.getSheetByName("集計シート");


  //NGset
  Totallsheet.getRange(11,16,34,5).clearContent();
  if(NGlist != ''){
      var range = Totallsheet.getRange(11,16,NGlist.length,5);
      range.setValues(NGlist);
      }

  //N/Aset
  Totallsheet.getRange(11,22,34,3).clearContent();
  if(NAlist != ''){
      var range = Totallsheet.getRange(11,22,NAlist.length,3);
      range.setValues(NAlist);
      }


  //試験対象外set
  Totallsheet.getRange(11,26,34,3).clearContent();
  if(Outlist != ''){
      var range = Totallsheet.getRange(11,26,Outlist.length,3);
      range.setValues(Outlist);
      }


  //保留set
  Totallsheet.getRange(11,30,34,3).clearContent();
  if(Keeplist != ''){
      var range = Totallsheet.getRange(11,30,Keeplist.length,3);
      range.setValues(Keeplist);
      }


  
};
