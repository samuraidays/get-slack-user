function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('管理', [
      {name: 'Slackユーザ取得', functionName: 'getSlackUser'},
    ]);
}

function getSlackUser() {
      
  // アクティブなシート取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      
  var accessToken = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";
  var slackURLBase = "https://slack.com/api"
  var slackUserListAPI = slackURLBase + "/users.list?token=" + accessToken
  
  var res = UrlFetchApp.fetch(slackUserListAPI) // kanmu.esa.io に登録されているユーザの一覧を取得する API にアクセスする
  var data = JSON.parse(res);  // APIから得られたデータを連想配列に変換する

  var rows = [];
  var rowcount = 0;
  // ヘッダ行
  rows.push(["USERNAME", "DELETED", "ID", "EMAIL", "IS_ADMIN"]);
  
  Logger.log(data.members)
  for (var i in data.members) {
      var cols = [];         
      cols.push(data.members[i].name);
      cols.push(data.members[i].deleted);
      cols.push(data.members[i].id);
      cols.push(data.members[i].profile.email);
      cols.push(data.members[i].is_admin);
    rows.push(cols);
  }
  rowcount = rowcount + data.members.length;
  sheet.getRange(1, 1, rowcount+1, 5).setValues(rows);
}


