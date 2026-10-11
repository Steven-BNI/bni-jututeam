/**
 * 贊助名額自動統計（貼到贊助表單所在的 Apps Script）
 *
 * 做兩件事：
 * 1) 把下面這個函式貼到整份程式最下面。
 * 2) 在 doPost 的 switch (data.action) 裡，緊接在 case 'sponsorSubmit' 那一行下面，新增一行：
 *      case 'sponsorCount':           return handleSponsorCount();
 * 然後「部署 → 管理部署作業 → 編輯 → 版本：新版本 → 部署」（網址不變）。
 */
function handleSponsorCount() {
  const counts = { main: 0, general: 0, gift: 0 };
  const sheet = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID).getSheetByName(SPONSOR_SHEET);
  if (sheet && sheet.getLastRow() > 1) {
    // 「方案代碼」在 SPONSOR_FIELDS 的第 1 欄，試算表第 2 欄（A 欄是報名時間）
    const vals = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    vals.forEach(function (r) {
      const k = String(r[0]).trim();
      if (counts.hasOwnProperty(k)) counts[k]++;
    });
  }
  return jsonResponse({ status: 'ok', counts: counts });
}
