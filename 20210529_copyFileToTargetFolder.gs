function copyFileToTargetFolder() {
  // 変数宣言
  var OutputFolderId,
    OutputFileName,
    OutputFileAmount,
    sheetName = '【出力結果】フォルダのURL',
    ss,
    column_index;
  ss    = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  column_index = 8; // H列のインデックス
  OutputFolderId   = sheet.getRange(4, column_index).getValue();         // 出力先フォルダID（H4セルのフォルダID）
  OutputFileName   = sheet.getRange(5, column_index).getValue();         // 出力ファイル名（H5セルのファイル名）
  OutputFileAmount = Number(sheet.getRange(6, column_index).getValue()); // 出力ファイル数（H6セルのファイル数）
  
  // テンプレートファイル（「yyyyMMdd(E)」）
  var TemplateFile = DriveApp.getFileById('1CBM96cEjVesKFxuKnlLYTETSWAElju2ZteZeEDskWDY');
  // 出力先フォルダ（H4セルのフォルダ）
  var OutputFolder = DriveApp.getFolderById(OutputFolderId);

  // ファイルをコピーする。
  for (let k = 0; k < OutputFileAmount; k++) {
    TemplateFile.makeCopy(OutputFileName, OutputFolder);
  }

  // ●実行ステータス入力：完了確認（H7セルへのステータス確認）
  // Utilities.sleep(1000);
  sheet.getRange(7, column_index).setValue('Terminated!');
  // ●実行ステータス入力：実行待ち（H7セルへのステータス入力）
  // Utilities.sleep(1000);
  // sheet.getRange(7, column_index).setValue('Waiting...');
}
