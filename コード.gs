//2021/01/08 既に貸し出しているものをoutできてしまう問題を発見

function doPost(e) {
  var verificationToken = e.parameter.token;
  if (verificationToken != 'OYKhypDN2XlKTCj2CEPLi2LP') { // AppのVerification Tokenを入れる
     throw new Error('Invalid token');
  }
  
  var command = e.parameter.text.split(' ');
  
  var result ='';
  var listStartRow = 1;
  var listStartColumn = 1;
  var listEndRow = 5; //機材の数
  var listEndColumn = 5; //
  var wifiList = getListRange(listStartRow, listStartColumn, listEndRow, listEndColumn).getValues();
  
  if(command[0]　== 'list') {
    result = getList(listStartRow, listStartColumn, listEndRow, listEndColumn);
    
  } else if(isEnableWifiName(command, 0, 0, wifiList, listEndRow) && isInOrOut(command, 2)) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inout');
    var lastRow = spreadsheet.getLastRow() + 1;
    var today = new Date();
    
    spreadsheet.getRange(lastRow, 1).setValue(command[0]);
    spreadsheet.getRange(lastRow, 2).setValue(command[1]);
    spreadsheet.getRange(lastRow, 3).setValue(command[2]);
    spreadsheet.getRange(lastRow, 4).setValue(today);
    
    result = "受け付けました。";
    
  } else {
    result = 'usage:\n/equip list\nequip一覧を表示します。\n\n/equip 端末ID 名前 in|out\nout（貸出）in（返却）を登録します。';
  }
  
  var response = {text: result};
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

function isInOrOut(command, column) {
  var result = false;
  
  if(command[column] == 'in' || command[column] == 'out'){
     result = true;
  }
  
  return result;
}

function isEnableWifiName(command, commandColumn, listColumn, wifiList, listEndRow) {
  var result = false;
  
  for(var i=0; i<listEndRow; i++){
      if(command[commandColumn] == wifiList[i][listColumn]){
      result = true;
      }
  }

  return result;
}

function getList(listStartRow, listStartColumn, listEndRow, listEndColumn){
  var result = '';
  var range;
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('list');
  range = spreadsheet.getRange(listStartRow, listStartColumn, listEndRow, listEndColumn);
  
  for(var i=0; i<listEndRow; i++){
    for(var j=0; j<listEndColumn; j++){
      result = result + range.getValues()[i][j] + ' | ';
    }
    result = result + '\n';
  }
  
  return result;
}

function getListRange(listStartRow, listStartColumn, listEndRow, listEndColumn){
  var result;
  var range;
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('list');
  range = spreadsheet.getRange(listStartRow, listStartColumn, listEndRow, listEndColumn);
  
  return range;
}