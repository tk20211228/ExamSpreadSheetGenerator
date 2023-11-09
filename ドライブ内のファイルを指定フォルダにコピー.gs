//ドライブ内のファイルを指定フォルダにコピー

//下記スクリプトを入力し、名前を設定して保存します。
//こちらは「入力データ_template」の名前を「入力データ_年月」に変換してフォルダにコピーする
// シートの自動作成スクリプト

function createSheet() {
  // テンプレートファイル
  var templateFile = DriveApp.getFileById('#ファイルID#');
  // 出力フォルダ
  var OutputFolder = DriveApp.getFolderById('#フォルダID#');
  // 出力ファイル名
  var OutputFileName = templateFile.getName().replace('_template', '')+'_'+Utilities.formatDate(new Date(), 'JST', 'yyyyMM')
  
  templateFile.makeCopy(OutputFileName, OutputFolder);
}
