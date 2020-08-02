function myFunction() {
  // スプレッドシートのインスタンスを設定する。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let num_sheet = ss.getNumSheets() - 1; // 対象のシートの枚数を取得
  
  // 諸々の変数を設定する。
  const column_of_value = 3; // 読み取る列のインデックス
  let row_submit   = 2;      // 「他の人に渡す単語」の行インデックス
  let arr_value_submit = []; // 入力した値
  
  for (let i = 0; i < num_sheet; i++) {
    arr_value_submit.push(readSubmitValue(ss.getSheets()[i], column_of_value, row_submit));
  }
  arr_value_submit.push(arr_value_submit[0]);
  arr_value_submit.splice(0, 1);
  
  for (let i = 0; i < num_sheet; i++) {
    InputValue(ss.getSheets()[i], column_of_value, arr_value_submit, num_sheet);
    
  }
}

function readSubmitValue(sheet, column_of_value, row_submit) {
  value_submit = sheet.getRange(row_submit, column_of_value).getValue();
  return value_submit;
}

function InputValue(sheet, column_of_value, arr_value_submit, num_sheet) {
  let row_player_base = 4; // 単語を格納する行の最上の行インデックス
  for (let i = 0; i < num_sheet; i++) {
    // シート名と同じプレイヤー名の行に書き込まない。
    if (sheet.getSheetName() != sheet.getRange(row_player_base + i,column_of_value - 1).getValue()) {
      sheet.getRange(row_player_base + i, column_of_value).setValue(arr_value_submit[i]);
      sheet.getRange(row_player_base + i, column_of_value).setBackground("#FFFFFF");
    }
    else {
      sheet.getRange(row_player_base + i, column_of_value).setValue('(不明)');
      sheet.getRange(row_player_base + i, column_of_value).setBackground("#808080");
    }
    
  }
}
