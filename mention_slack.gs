
var columnA = 1
var columnB = 2
var columnC = 3
var columnD = 4
var columnE = 5

var slack = {
  postUrl:   'https://slack.com/api/chat.postMessage',
  userName:  '給水当番連絡botさん',
  iconEmoji: ':gori2:',
}

var meibo_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名簿")

function post2Slack(apply_data) {
  
  var sheet = convertSheet2Json.convert("スケジュール")
  var today = new Date()
  var formattedDate = Utilities.formatDate(today, "JST", "yyyy/MM/dd")
  
  var slack_mention_name = ""
  var remarks = ""
    
  for (var i=0; i < sheet.length; i++ ) {
    if (sheet[i].date == "") continue
    var v = new Date( sheet[i].date )
    
    var date = Utilities.formatDate(v, "JST", "yyyy/MM/dd")
    if (formattedDate != date) continue
    
    var assign = sheet[i].assign
    if (assign == "") {
      break
    }
    slack_mention_name = convert_slack_mention_name_(assign)
    remarks = sheet[i].remarks
    break
  }
  
  if (slack_mention_name != "") {
    var message = "本日の給水担当は <@"+slack_mention_name+"> さんです！ \n※担当者が初めて作業する場合は、チーム内でフォローをお願いします。また、担当者の急なお休みなどの対応もチーム内でお願いします:gori2:\n\n";
    if (remarks != "") {
      message += "\n備考: *"+remarks+"* \n\n";
    }
    message += "○やるべきこと\n・給水\n・排水タンクのお水を捨てる\n※加湿器は３台あります！！\n\n"
    message += "○お手入れのPDF \n xxxxxxxxxx\n\n"
    message += "○当番表\n https://docs.google.com/a/mixi.co.jp/spreadsheets/d/1myWeofPmcIXNfTHADQwJm5vkE_Q59LHWu_Uu8nGFAFM/edit?usp=sharing\n\n"
    message += "\nよろしくお願い致します!!"
    
    
    var slackApp = SlackApp.create(getSlackToken_());
    // var channel_id = "G5BPZ9ZPS" //test1
    var channel_id = "G5E1FNH8C" //test3
    slackApp.postMessage(channel_id,message,{ username: slack["userName"], icon_emoji: slack["iconEmoji"]});
  }
}

function convert_slack_mention_name_(name) {
  var row = findRow_(meibo_sheet, name, columnA)
  return meibo_sheet.getRange(row, columnB).getValue()
}

function findRow_(sheet,val,col){ 
  var data = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<data.length;i++){
    if(data[i][col-1] === val){
      return i+1;
    }
  }
  return 0;
}

function getSlackToken_() {
  var sheet = convertSheet2Json.convert("slack token")
  return sheet[0].slack_token
}
