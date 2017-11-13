var columnA = 1
var columnB = 2
var columnC = 3
var columnD = 4
  
function input_data() {
  

  var name_list = get_random_name_()
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("スケジュール")
  var start_row = getLastRowNumber_ColumnD_(sheet) + 1
  
  for (var i=start_row; i < start_row + 100; i++) {
    var A = sheet.getRange(i,columnA)
    var B = sheet.getRange(i,columnB)
    var C = sheet.getRange(i,columnC)
    var D = sheet.getRange(i,columnD)
    
    if (B.getValue() == "土曜日" || B.getValue() == "日曜日") {
      B.setBackground("#FFAAAA")
      continue; 
    }
    if (C.getValue().length > 0) {
      B.setBackground("#FFAAAA")
      continue; 
    }
    D.setValue(name_list[0])
    name_list.shift();
    if (name_list.length == 0) {
      break 
    }
  }
}

function get_random_name_() {
  var sheet = convertSheet2Json.convert("名簿")
  var name_list = []
  for (var i=0; i < sheet.length; i++ ) {
    name_list.push(sheet[i].name)
  }
  return shuffle_(name_list)
}

function getLastRowNumber_ColumnD_(sheet){
　　var last_row = sheet.getLastRow();
　　for(var i = last_row; i >= 1; i--){
　　　　if(sheet.getRange(i, 4).getValue() != ''){
　　　　　　return i;
　　　　　　break;
　　　　}
　　}
}

function shuffle_(array) {
  var n = array.length, t, i;

  while (n) {
    i = Math.floor(Math.random() * n--);
    t = array[n];
    array[n] = array[i];
    array[i] = t;
  }

  return array;
}