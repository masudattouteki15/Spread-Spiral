function copyFileToTargetFolder() {
  // 変数宣言
  var OutputFolderId,
    OutputFileName,
    OutputFileAmount,
    sheetName = '【出力結果】フォルダのURL',
    ss;
  ss    = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  OutputFolderId   = sheet.getRange(4, 7).getValue();         // 出力先フォルダID（G4セルのフォルダID）
  OutputFileName   = sheet.getRange(5, 7).getValue();         // 出力ファイル名（G5セルのファイル名）
  OutputFileAmount = Number(sheet.getRange(6, 7).getValue()); // 出力ファイル数（G6セルのファイル数）
  
  // テンプレートファイル（「yyyyMMdd(E)」）
  var TemplateFile = DriveApp.getFileById('1CBM96cEjVesKFxuKnlLYTETSWAElju2ZteZeEDskWDY');
  // 出力先フォルダ（G4セルのフォルダ）
  var OutputFolder = DriveApp.getFolderById(OutputFolderId);

  // ファイルをコピーする。
  for (let k = 0; k < OutputFileAmount; k++) {
    TemplateFile.makeCopy(OutputFileName, OutputFolder);
  }

  // ●実行ステータス入力：完了確認（G7セルへのステータス確認）
  Utilities.sleep(3000);
  sheet.getRange(7, 7).setValue('Terminated!');
  // ●実行ステータス入力：実行待ち（G7セルへのステータス入力）
  Utilities.sleep(5000);
  sheet.getRange(7, 7).setValue('Waiting...');
}
