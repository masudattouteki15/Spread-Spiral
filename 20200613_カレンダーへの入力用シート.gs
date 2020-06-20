// カレンダーに予定を追加する。
function createEvents() {
  // カレンダーのインスタンスを設定する。
  const calendar = CalendarApp.getDefaultCalendar();
  
  // スプレッドシートのインスタンスを設定する。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('カレンダー入力');
  const range = sheet.getRange('C3');
  
  // 諸々の変数を設定する。
  const column_of_key = 2; // キーの列：数値
  let row = 1;             // 走査の開始行：数値
  let key;                 // キー：数値
  let array_to_input = []; // 入力用：配列
  let startDate;           // 開始日：文字列
  let endDate;             // 終了日：文字列
  let color;               // 色番号：数値
  let a_day;               // とある日にち：文字列
  let num_of_date;         // 日数：数値
  let startTime;           // 開始時間：文字列
  let endTime;             // 終了時間：文字列
  let options = {};        // 入力のオプション枠：辞書
  let event;               // イベント：Event
  let day_of_EOM = {'01': 31, '02': 28, '03': 31, 
                    '04': 30, '05': 31, '06': 30, 
                    '07': 31, '08': 31, '09': 30, 
                    '10': 31, '11': 30, '12': 31} // 月末の日付：辞書

  // カレンダー入力データの起点を決定。
  while(sheet.getRange(row,column_of_key).getValue() != '毎日入力（始まり）'){
    row++;
  }
  
  // 「options」行まで、入力値を読み込む。
  while(sheet.getRange(row,column_of_key).getValue() != 'options'){
    key = sheet.getRange(row,column_of_key).getValue();
    if(key == 'startDate'){
      startDate = Utilities.formatDate(sheet.getRange(row,column_of_key + 1).getValue(), "Asia/Tokyo", "yyyy/MM/dd");
    }
    if(key === 'endDate'){
      endDate = Utilities.formatDate(sheet.getRange(row,column_of_key + 1).getValue(), "Asia/Tokyo", "yyyy/MM/dd");
    }
    if(key === 'title'){
      array_to_input[0] = sheet.getRange(row,column_of_key + 1).getValue();
    }
    if(key === 'startTime'){
      startTime = sheet.getRange(row,column_of_key + 1).getValue();
    }
    if(key === 'endTime'){
      endTime = sheet.getRange(row,column_of_key + 1).getValue();
    }
    if(key === 'color'){
      color = Number(sheet.getRange(row,column_of_key + 1).getValue());
    }
    row++;
  }
  
  // 日数を設定する。（開始月と終了月が異なるかどうかで場合分け。閏年には非対応。）
  if (endDate.slice(5, 7) === startDate.slice(5, 7)) {
    num_of_date = Number(endDate.slice(-2)) - Number(startDate.slice(-2)) + 1;
  } else {
    num_of_date = Number(endDate.slice(-2)) + Number(day_of_EOM[startDate.slice(5, 7)]) - Number(startDate.slice(-2)) + 1;
  }
  
  // 繰り返し方を設定する。
//  let recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.WEDNESDAY).until(new Date("2018/08/31"));
  let recurrence = CalendarApp.newRecurrence().addDailyRule().times(Number(num_of_date));
  
  // 「option」の行をスキップ。
  row++;
  
  // 「options」内に入力する値を読み込む。
  while(sheet.getRange(row,column_of_key).getValue() != '毎日入力（終わり）'){
    key = sheet.getRange(row,column_of_key).getValue();
    options[String(key)] = sheet.getRange(row,column_of_key + 1).getValue();
    row++;
  }
  
  //  毎日入力①（欠点：繰り返しじゃない）
//  for (let i = 0; i <= num_of_date; i++) {
//    a_day = Number(startDate.slice(-2)) + i;
//    array_to_input[1] = new Date(startDate.slice(0,-2) + a_day + " " + startTime);
//    array_to_input[2] = new Date(startDate.slice(0,-2) + a_day + " " + endTime);
//    event = calendar.createEvent(array_to_input[0], array_to_input[1], array_to_input[2], options);
//    event.setColor(color);
//  }
  
  //  毎日入力②（欠点：特になし）
  array_to_input[1] = new Date(startDate + " " + startTime);
  array_to_input[2] = new Date(startDate + " " + endTime);
  // Browser.msgBox('テスト');
  event = calendar.createEventSeries(array_to_input[0], array_to_input[1], array_to_input[2], recurrence, options);
  event.setColor(color);
  
}
