function getFileListInFolder() {
  var folder_id,
    folder,
    files,
    sheetName = '【出力結果】フォルダのURL',
    key = DriveApp.getRootFolder().getId(),
    ss;
    name = "",
    i = 3, //フォルダを処理する行位置
    column_for_fileList = 3;
  ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  folder_id = sheet.getRange(1, 2).getValue();
  console.log(folder_id);
  folder = DriveApp.getFolderById(folder_id);
  files = folder.getFiles();

  // var folders = DriveApp.searchFolders("'"+key+"' in parents");
  while(files.hasNext()) {
    i++;
    var file = files.next();
    sheet.getRange(i, 1+column_for_fileList).setValue(name + file.getName());
    sheet.getRange(i, 2+column_for_fileList).setValue(file.getUrl());
  }

}
