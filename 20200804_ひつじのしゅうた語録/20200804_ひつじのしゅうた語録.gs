//// test code
//function myFunction() {
//  Browser.msgBox(getShutaMessages());
//  writeLog('うつうつつつ', 'fewfwqfewf');
//}

var CHANNEL_ACCESS_TOKEN = '8usFHhLgit4p5yPpROlwef3OtLQQ1qVYKADWrCXZvqnQbN7A739aJKSuUAmjpvlLDdB6MrSi6GA4cusG3NjSTInLn1A8Siw5TUQrgsm5RDbxPv4Gdesvw9Sxt5FbaVkRDyP00VLnGtFCfC1v5xOThAdB04t89/1O/w1cDnyilFU=';

function doGet(e) {
  return ContentService.createTextOutput(UrlFetchApp.fetch("http://ip-api.com/json"));
}

function doPost(e) {
  var event      = JSON.parse(e.postData.contents).events[0];
  var replyToken = event.replyToken;
  if (typeof replyToken === 'undefined') {
    return;
  }
  var userId   = event.source.userId; // なぜか知らんが、'undefined'になる。（event.sourceまでは取れてる。）
  // writeLog('うつうつつつ1', 'fewfwqfewf', event.source); // 友達追加した人がグループで話すと、「user」から「room」になる？
  var username = getUserName(userId);

  if(event.type == 'message') {
    var userMessage = event.message.text;
    var replyMessages = getShutaMessages();
    sendMessage(replyToken, replyMessages);
    writeLog(userMessage, replyMessages, event);
    // writeLog(userMessage, replyMessages);
    return ContentService.createTextOutput(
      JSON.stringify({'content': 'ok'})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendMessage(replyToken, replyMessages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var messages = replyMessages.map(function (v) {
    return {'type': 'text', 'text': v};
  });
  UrlFetchApp.fetch(url, {
     'headers': {
       'Content-Type' : 'application/json; charset=UTF-8',
       'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
     },
     'method' : 'post',
     'payload': JSON.stringify({
       'replyToken': replyToken,
       'messages'  : messages,
     }),
   });
}

function getUserName(userId) {
  var url         = 'https://api.line.me/v2/bot/profile/' + userId;
  var userProfile = UrlFetchApp.fetch(url,{
    'headers': {
      'Authorization' :  'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
  })
  return JSON.parse(userProfile).displayName;
}

function getShutaMessages() {
  const spreadsheetId = "1Wwaq5O7vzT0OdVongtVI0AmSRh5SCfEd4_HVXnfT3sw";
  const sheetName     = "shuta";
  let   spreadsheet   = SpreadsheetApp.openById(spreadsheetId);
  let   sheet         = spreadsheet.getSheetByName(sheetName);
  const column_of_key = 1;
  let   line          = "デフォルト";
  let   last_row      = 1; // 最終行のインデックス
  let   selected_row; // 選択された行
  
  last_row     = Number(sheet.getRange(2, column_of_key + 5).getValue());
  selected_row = Math.floor(Math.random() * (last_row)) + 1;
  if (selected_row == 0 || selected_row == last_row) {
    selected_row = 1;
  }
  let messages = [sheet.getRange(selected_row, column_of_key).getValue()];
  return messages;
}

function writeLog(userMessage, replyMessages, event) {
  const spreadsheetId = "1Wwaq5O7vzT0OdVongtVI0AmSRh5SCfEd4_HVXnfT3sw";
  const sheetName     = "log_shuta";
  let   spreadsheet   = SpreadsheetApp.openById(spreadsheetId);
  let   sheet         = spreadsheet.getSheetByName(sheetName);
  const column_of_key = 1;
  let   line          = "デフォルト";
  let   row           = 1;
  
  row = Number(sheet.getRange(2, column_of_key + 5).getValue());
  let date    = new Date();
  sheet.getRange(row, column_of_key).setValue(date);
  sheet.getRange(row, column_of_key + 1).setValue(String(userMessage));
  sheet.getRange(row, column_of_key + 2).setValue(String(replyMessages[0]));
  sheet.getRange(row, column_of_key + 3).setValue(String(event.type));
  sheet.getRange(2  , column_of_key + 5).setValue(row + 1);
}
