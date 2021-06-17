function getFolderListInFolder() {
  var folder = DriveApp.getRootFolder(),
    // folders = folder.getFolders,
    sheetName = '【出力結果】フォルダのURL',
    key = DriveApp.getRootFolder().getId(),
    ss;
    name = "",
    i = 3, //フォルダを処理する行位置
    column_for_fileList = 1; // A列のインデックス
  ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  
  var folders = DriveApp.searchFolders("'"+key+"' in parents");
  while(folders.hasNext()) {
    i++;
    var folder = folders.next();
    sheet.getRange(i, 0+column_for_fileList).setValue(name + folder.getName());
    sheet.getRange(i, 1+column_for_fileList).setValue(folder.getId());
    sheet.getRange(i, 2+column_for_fileList).setValue(folder.getUrl());
  }
}
