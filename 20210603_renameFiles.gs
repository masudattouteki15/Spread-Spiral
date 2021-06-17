function renameFiles() {
  var file,
    sheetName = '【出力結果】フォルダのURL',
    ss;
    name = "",
    i = 4, //フォルダを処理する行位置
    column_for_fileList = 12; // L列
  ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  
  while(sheet.getRange(i, column_for_fileList).getValue() != '') {
    file = DriveApp.getFileById(sheet.getRange(i, 1+column_for_fileList).getValue());
    file.setName(sheet.getRange(i, column_for_fileList).getValue());
    i++;
  }
}
