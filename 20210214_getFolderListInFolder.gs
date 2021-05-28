function getFolderListInFolder() {
  var folder = DriveApp.getRootFolder(),
    // folders = folder.getFolders,
    sheetName = '【出力結果】フォルダのURL',
    key = DriveApp.getRootFolder().getId(),
    ss;
    name = "",
    i = 3; //フォルダを処理する行位置
  ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  
  var folders = DriveApp.searchFolders("'"+key+"' in parents");
  while(folders.hasNext()) {
    i++;
    var folder = folders.next();
    sheet.getRange(i, 1).setValue(name + folder.getName());
    sheet.getRange(i, 2).setValue(folder.getUrl());
  }
}
