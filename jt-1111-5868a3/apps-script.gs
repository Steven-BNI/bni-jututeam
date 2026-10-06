// Google Apps Script：把報名資料寫進試算表
// 使用方式：開一份 Google 試算表 → 擴充功能 → Apps Script → 貼上這段 → 部署為「網頁應用程式」
//   執行身分：我；誰可以存取：所有人 → 複製網址，貼到網頁 index.html 的 REG_ENDPOINT
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('報名') || ss.insertSheet('報名');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['時間', '姓名', '手機', 'Email', '公司/產業', '身份', '邀請人', '備註']);
  }
  var d = JSON.parse(e.postData.contents);
  sh.appendRow([new Date(), d.name, d.phone, d.email, d.company, d.role, d.inviter, d.note]);
  return ContentService.createTextOutput('ok');
}
