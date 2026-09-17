/**
 * BNI 聚團隊｜培訓報名系統 - Google Apps Script
 * 金流：黑貓PAY 多元支付（統一金流 PAYUNi）
 * ════════════════════════════════════════════════
 * 部署方式：
 *   擴充功能 → Apps Script → 貼上此程式碼（取代全部）
 *   → 部署 → 管理部署 → 編輯 → 版本「新版本」→ 儲存
 *   （不需要重新部署，更新版本即可）
 * ════════════════════════════════════════════════
 */

// ══════════════════════════════════════
// ⚙️  設定區
// ══════════════════════════════════════
const SETTINGS = {
  SPREADSHEET_ID: '1JKHpemlHQd_iCp1FpRwyolwL1oL3LNNKtvd5AVEZ-eI',
  SHEET_NAME: '報名紀錄',

  // 黑貓PAY API 設定
  PAYUNI_USERNAME: '934900020001',
  PAYUNI_PASSWORD: 'O9241627H^',
  PAYUNI_CUST_ID:  '934900020001',
  PAYUNI_BASE_URL: 'https://cocs.4128888card.com.tw',
  // 測試環境（測試時改用這個）：
  // PAYUNI_BASE_URL: 'https://test.4128888card.com.tw/app',

  // 收單行（統一金流已開通）
  ACQUIRER_TYPE: 'payuni',

  // 付款完成後回傳網址（APN，需是可接收 POST 的公開網址）
  APN_URL: 'https://script.google.com/macros/s/AKfycbx4RlPROgjqOPWaI76XZ6Lh4kZIQMww5IeiDTIUM2y8VuBGf9OCPWoQHT4JNTl6OOWPYw/exec',

  // 付款成功後跳回的頁面（給用戶看的）
  SUCCESS_URL: 'https://steven-bni.github.io/bni-jututeam/register-return.html?result=success',
};

// ══════════════════════════════════════
// 欄位定義（新增「餐點」於最後一欄，不影響既有欄位索引）
// ══════════════════════════════════════
const HEADERS = [
  '報名時間', '培訓名稱', '培訓日期', '地點',
  '姓名', '分會名稱', '電話', 'Email',
  '報名身份', '費用', '付款狀態', '交易編號', '付款時間', '付款網址', '餐點', '付款方式', '人工核對', 'ATM末五碼',
];

// ══════════════════════════════════════
// 主入口
// ══════════════════════════════════════
function doPost(e) {
  try {
    const raw = e.postData ? e.postData.contents : '';

    // 黑貓PAY APN 回傳（JSON 格式，含 cust_order_no）
    let data;
    try { data = JSON.parse(raw); } catch(err) {
      // 嘗試 form-urlencoded 格式（黑貓PAY APN 可能用此格式）
      data = {};
      raw.split('&').forEach(function(pair) {
        var parts = pair.split('=');
        if (parts[0]) data[decodeURIComponent(parts[0].replace(/\+/g,' '))] = decodeURIComponent((parts[1]||'').replace(/\+/g,' '));
      });
    }

    if (raw && !data.action && !data.trainingName) {
      return handlePayUniAPN(data);
    }

    switch (data.action) {
      case 'register':              return handleRegistration(data);
      case 'cancelRequest':         return handleCancelRequest(data);
      case 'lookupPaidFee':         return handleLookupPaidFee(data);
      case 'submitFeedback':        return handleSubmitFeedback(data);
      case 'lookupMyRegistrations': return handleLookupMyRegistrations(data);
      case 'dnaLogin':              return handleDnaLogin(data);
      case 'dnaGetAllData':         return handleDnaGetAllData(data);
      case 'dnaCheckinV2':          return handleDnaCheckinV2(data);
      case 'dnaChangePassword':     return handleDnaChangePassword(data);
      default:                       return jsonResponse({ status: 'error', message: '未知 action' });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput('OK');
}

// ══════════════════════════════════════
// 1. 處理報名
// ══════════════════════════════════════
function handleRegistration(data) {
  const sheet   = getOrCreateSheet();

  // 防止重複報名：同一人（姓名+分會）已有這場（培訓名稱+日期）非失敗狀態的報名紀錄，就擋下
  const existingData = sheet.getDataRange().getValues();
  const targetDate = normalizeDateValue(data.trainingDate);
  for (let i = 1; i < existingData.length; i++) {
    const status = String(existingData[i][10] || '');
    const isFailed = status === '付款失敗' || status === '建單失敗';
    if (
      !isFailed &&
      String(existingData[i][1]).trim() === String(data.trainingName).trim() &&
      normalizeDateValue(existingData[i][2]) === targetDate &&
      String(existingData[i][4]).trim() === String(data.name).trim() &&
      String(existingData[i][5]).trim() === String(data.chapter).trim()
    ) {
      return jsonResponse({ status: 'error', message: '您已經報名過這場培訓了，若需要修改或取消，請使用「申請延期」功能，不要重複報名。' });
    }
  }

  const tradeNo = generateTradeNo();
  const now     = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const identityLabel = {
    general:    '個人報名',
    mentor:     '認證導師',
    staff:      '統籌',
    lecturer:   '講師',
    gold:       '金質獎章得主',
    ambassador: '區域培訓大使',
  }[data.identity] || data.identity;

  // 寫入試算表
  const payMethodLabel = data.fee === 0 ? '免費' : (data.payMethod === 'atm' ? 'ATM轉帳' : '信用卡');
  const atmLast5 = data.payMethod === 'atm' ? (data.atmLast5 || '未填') : '';

  sheet.appendRow([
    now,
    data.trainingName,
    data.trainingDate,
    data.location,
    data.name,
    data.chapter,
    data.phone,
    data.email,
    identityLabel,
    data.fee,
    data.fee === 0 ? '免費（已完成）' : (data.payMethod === 'atm' ? 'ATM待確認' : '待付款'),
    tradeNo,
    '',   // 付款時間
    '',   // 付款網址
    data.meal || '不需要',
    payMethodLabel,
    '',   // 人工核對
    atmLast5,
  ]);

  // 同步培訓公告報名人數
  updateRegistrationCount(data.trainingName, data.trainingDate);

  // 免費：直接回傳成功
  if (data.fee === 0) {
    return jsonResponse({ status: 'ok', free: true });
  }

  // ATM 轉帳：末五碼已存入獨立欄位，不再需要另外覆寫付款狀態
  if (data.payMethod === 'atm') {
    return jsonResponse({ status: 'ok', free: false, atm: true });
  }

  // 信用卡：取得 Token → 建立刷卡訂單 → 回傳付款網址
  try {
    const token  = getPayUniToken();
    const result = createCocsOrder(token, tradeNo, data.fee, data.trainingName);
    const paymentUrl   = result.url;
    const cocsOrderNo  = result.cocsOrderNo;

    // 用黑貓PAY 的訂單編號覆蓋原本的 BNI 編號，確保 APN 回傳時能對應
    if (cocsOrderNo && cocsOrderNo !== tradeNo) {
      updateTradeNo(tradeNo, cocsOrderNo);
      tradeNo = cocsOrderNo;
    }

    // 把付款網址存回試算表
    updatePaymentUrl(tradeNo, paymentUrl);

    return jsonResponse({ status: 'ok', free: false, paymentUrl: paymentUrl });
  } catch (err) {
    // 建立訂單失敗：把狀態改為「建單失敗」
    updatePaymentStatus(tradeNo, '建單失敗', '');
    return jsonResponse({ status: 'error', message: '建立付款訂單失敗：' + err.message });
  }
}

// ══════════════════════════════════════
// 2. 取得黑貓PAY Token（有效3小時）
// ══════════════════════════════════════
function getPayUniToken() {
  const url      = SETTINGS.PAYUNI_BASE_URL + '/Token';
  const payload  = 'grant_type=password'
    + '&username=' + encodeURIComponent(SETTINGS.PAYUNI_USERNAME)
    + '&password=' + encodeURIComponent(SETTINGS.PAYUNI_PASSWORD);

  const response = UrlFetchApp.fetch(url, {
    method:      'post',
    contentType: 'application/x-www-form-urlencoded',
    payload:     payload,
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (!result.access_token) {
    throw new Error('取得 Token 失敗：' + JSON.stringify(result));
  }
  return result.access_token;
}

// ══════════════════════════════════════
// 3. 建立刷卡訂單
// ══════════════════════════════════════
function createCocsOrder(token, tradeNo, amount, trainingName) {
  const url      = SETTINGS.PAYUNI_BASE_URL + '/api/Collect';
  const sendTime = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  const body = JSON.stringify({
    cmd:           'CocsOrderAppend',
    cust_id:       SETTINGS.PAYUNI_CUST_ID,
    cust_order_no: tradeNo,
    order_amount:  amount,
    order_detail:  'BNI培訓報名-' + trainingName,
    acquirer_type: SETTINGS.ACQUIRER_TYPE,
    send_time:     sendTime,
    success_url:   SETTINGS.SUCCESS_URL + '&trade_no=' + tradeNo + '&amount=' + amount,
    apn_url:       SETTINGS.APN_URL,
  });

  const response = UrlFetchApp.fetch(url, {
    method:      'post',
    contentType: 'application/json',
    headers:     { 'Authorization': 'Bearer ' + token },
    payload:     body,
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  console.log('建單回傳：' + JSON.stringify(result));
  if (result.status !== 'OK' || !result.url) {
    throw new Error(result.msg || '建立刷卡訂單失敗');
  }
  // 回傳 url 和黑貓PAY 的訂單編號（order_no 或 cust_order_no）
  return { url: result.url, cocsOrderNo: result.order_no || result.cust_order_no || tradeNo };
}

// ══════════════════════════════════════
// 3.5 查詢刷卡訂單真實狀態（官方 API，不信任 APN 內容本身）
// ══════════════════════════════════════
function queryCocsOrder(token, tradeNo) {
  const url  = SETTINGS.PAYUNI_BASE_URL + '/api/Collect';
  const body = JSON.stringify({
    cmd:           'CocsOrderQuery',
    cust_id:       SETTINGS.PAYUNI_CUST_ID,
    cust_order_no: tradeNo,
  });

  const response = UrlFetchApp.fetch(url, {
    method:      'post',
    contentType: 'application/json',
    headers:     { 'Authorization': 'Bearer ' + token },
    payload:     body,
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  console.log('訂單查詢回傳：' + JSON.stringify(result));
  return result; // { status, process_code, order_amount, ... } 或 { status:'ERROR', msg }
}

// ══════════════════════════════════════
// 4. 接收黑貓PAY APN 付款通知
// ══════════════════════════════════════
function handlePayUniAPN(data) {
  try {
    // Log 完整回傳內容以便 debug（僅供除錯，不作為付款依據）
    console.log('APN 收到（僅作觸發訊號，不直接信任內容）：' + JSON.stringify(data));

    const tradeNo = data.cust_order_no || data.order_no || data.OrderNo || data.MerchantOrderNo || '';

    if (!tradeNo) {
      console.error('APN 沒有交易編號');
      return ContentService.createTextOutput('ERROR');
    }

    const token = getPayUniToken();
    verifyAndUpdateOrder(token, tradeNo);

    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('handlePayUniAPN error:', err);
    return ContentService.createTextOutput('ERROR');
  }
}

// ══════════════════════════════════════
// 4.5 核心核實邏輯（APN 觸發與定時輪詢共用）
//     不相信外部通知內容，一律主動向黑貓PAY 官方查詢真實狀態
// ══════════════════════════════════════
function verifyAndUpdateOrder(token, tradeNo) {
  // 訂單必須存在且目前狀態為待付款/ATM待確認，才進一步查證（避免對亂猜編號浪費查詢）
  const rowInfo = findFullRowByTradeNo(tradeNo);
  if (!rowInfo) {
    console.log('查無對應訂單，略過：' + tradeNo);
    return;
  }
  if (rowInfo.status !== '待付款' && rowInfo.status.indexOf('ATM待確認') !== 0) {
    console.log('訂單目前狀態非待付款（' + rowInfo.status + '），略過重複處理：' + tradeNo);
    return;
  }

  // ⭐ 核心防偽：主動呼叫官方 CocsOrderQuery 查真實狀態
  const order = queryCocsOrder(token, tradeNo);

  if (order.status !== 'OK') {
    console.log('查詢訂單失敗，暫不更新狀態：' + tradeNo + ' ' + JSON.stringify(order));
    return;
  }

  // 金額核對：官方查詢金額須與試算表記錄一致
  if (Number(order.order_amount) !== Number(rowInfo.fee)) {
    console.error('查詢金額與訂單記錄不符（官方：' + order.order_amount + '，記錄：' + rowInfo.fee + '），轉人工核對：' + tradeNo);
    updatePaymentStatus(tradeNo, '待人工核對（金額不符）', '');
    return;
  }

  const paidAt = order.process_code_update_time || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const code   = Number(order.process_code);

  if (code === 15 || code === 22) {
    // 15=授權完成 22=請款完成，皆視為付款成功
    updatePaymentStatus(tradeNo, '已付款', paidAt);
    const rowData = findRowByTradeNo(tradeNo);
    if (rowData) updateRegistrationCount(rowData.trainingName, rowData.trainingDate);
    console.log('官方查證付款成功：' + tradeNo + ' process_code=' + code);
  } else if (code === 16 || code === 18) {
    // 16=授權失敗 18=取消授權失敗
    updatePaymentStatus(tradeNo, '付款失敗', paidAt);
    console.log('官方查證付款失敗：' + tradeNo + ' process_code=' + code);
  } else {
    // 其他中間狀態（例如13刷卡確認頁、14繳款人確認、20-21請款中）：先不變更，等下次再查
    console.log('訂單處於中間狀態，暫不變更：' + tradeNo + ' process_code=' + code);
  }
}

// ══════════════════════════════════════
// 4.8 一次性修正表頭（欄位新增後表頭沒跟著更新時使用）
//     用法：執行這個函式，會把「報名紀錄」第一列覆蓋成正確的完整欄位標題
// ══════════════════════════════════════
function fixHeaderRow() {
  const sheet = getOrCreateSheet();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  console.log('表頭已修正為：' + HEADERS.join('、'));
}

// ══════════════════════════════════════
// 4.9 排程主函式（觸發條件請指向這個函式，取代單獨的 pollPendingOrders）
//     每次執行會依序做兩件事：
//     1) 主動查證信用卡待付款訂單的真實狀態（黑貓PAY 官方 API）
//     2) 套用你在「人工核對」欄標記 Y 的 ATM 轉帳確認
// ══════════════════════════════════════
function scheduledTasks() {
  pollPendingOrders();
  applyManualConfirmations();
}

// ══════════════════════════════════════
// 4.7 套用人工核對（處理 ATM 或其他人工確認過的付款）
//     用法：在「報名紀錄」的「人工核對」欄填 Y，存檔後執行這個函式
// ══════════════════════════════════════
function applyManualConfirmations() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  const statusCol  = headers.indexOf('付款狀態');
  const paidAtCol  = headers.indexOf('付款時間');
  const methodCol  = headers.indexOf('付款方式');
  const checkCol   = headers.indexOf('人工核對');
  const nameCol    = headers.indexOf('培訓名稱');
  const dateCol    = headers.indexOf('培訓日期');

  if (checkCol === -1) {
    console.error('找不到「人工核對」欄位，請先在「報名紀錄」工作表最後新增一欄，標題填「人工核對」');
    return;
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  let processed = 0;

  for (let i = 1; i < data.length; i++) {
    const mark = String(data[i][checkCol] || '').trim();
    if (mark !== 'Y' && mark !== 'y') continue;

    const row = i + 1; // 1-indexed for setRange

    sheet.getRange(row, statusCol + 1).setValue('已付款（人工核對）');
    if (!data[i][paidAtCol]) {
      sheet.getRange(row, paidAtCol + 1).setValue(now);
    }
    if (methodCol !== -1 && !data[i][methodCol]) {
      sheet.getRange(row, methodCol + 1).setValue('ATM轉帳');
    }
    sheet.getRange(row, checkCol + 1).setValue('已處理');

    updateRegistrationCount(data[i][nameCol], data[i][dateCol]);
    processed++;
  }

  console.log('已套用 ' + processed + ' 筆人工核對');
}

// ══════════════════════════════════════
// 4.6 定時輪詢（不依賴 APN 是否送達／何時送達）
//     請在 Apps Script 設定「時間驅動」觸發條件，建議每 5～10 分鐘執行一次
// ══════════════════════════════════════
function pollPendingOrders() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  // 收集所有待付款 / ATM待確認 的交易編號
  const pending = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][10] || '');
    const tradeNo = data[i][11];
    if (tradeNo && (status === '待付款' || status.indexOf('ATM待確認') === 0)) {
      pending.push(tradeNo);
    }
  }

  if (pending.length === 0) {
    console.log('目前沒有待確認的訂單');
    return;
  }

  console.log('輪詢 ' + pending.length + ' 筆待付款訂單：' + pending.join(', '));

  const token = getPayUniToken(); // 共用同一組 Token，避免重複索取
  pending.forEach(function(tradeNo) {
    try {
      verifyAndUpdateOrder(token, tradeNo);
    } catch (err) {
      console.error('輪詢訂單發生錯誤：' + tradeNo + ' ' + err.message);
    }
  });
}

// ══════════════════════════════════════
// 5. 同步報名人數到「培訓公告」工作表
// ══════════════════════════════════════
// ══════════════════════════════════════
// 5.5 手動重新計算「所有場次」的報名人數
//     用途：欄位異動導致人數卡住不動時，手動執行一次修正全部資料
//     用法：Apps Script 編輯器上方選這個函式 → 執行
// ══════════════════════════════════════
function recalculateAllCounts() {
  const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  const announcementSheet = ss.getSheetByName('培訓公告');
  if (!announcementSheet) {
    console.error('找不到「培訓公告」工作表');
    return;
  }

  const annData    = announcementSheet.getDataRange().getValues();
  const annHeaders = annData[0];
  const nameCol = annHeaders.indexOf('培訓名稱');
  const dateCol = annHeaders.indexOf('培訓日期');

  if (nameCol === -1 || dateCol === -1) {
    console.error('「培訓公告」工作表找不到必要欄位（培訓名稱／培訓日期）');
    return;
  }

  let updated = 0;
  for (let i = 1; i < annData.length; i++) {
    const trainingName = annData[i][nameCol];
    const trainingDate = annData[i][dateCol];
    if (!trainingName || !trainingDate) continue;
    updateRegistrationCount(trainingName, trainingDate);
    updated++;
  }
  console.log('已重新計算 ' + updated + ' 場培訓的報名人數');
}

function updateRegistrationCount(trainingName, trainingDate) {
  try {
    const ss               = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
    const announcementSheet = ss.getSheetByName('培訓公告');
    if (!announcementSheet) return;

    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();
    const targetDate = normalizeDateValue(trainingDate);

    let count = 0;
    for (let i = 1; i < regData.length; i++) {
      if (String(regData[i][1]).trim() === String(trainingName).trim() && normalizeDateValue(regData[i][2]) === targetDate) {
        const st = String(regData[i][10] || '');
        // 免費／已付款（含人工核對）／ATM待確認 皆計入報名人數（避免因尚未收到付款確認而漏算）
        if (st.indexOf('已付款') === 0 || st === '免費（已完成）' || st.indexOf('ATM待確認') === 0) count++;
      }
    }

    const annData    = announcementSheet.getDataRange().getValues();
    const annHeaders = annData[0];
    const nameCol  = annHeaders.indexOf('培訓名稱');
    const dateCol  = annHeaders.indexOf('培訓日期');
    const countCol = annHeaders.indexOf('報名人數');

    if (nameCol === -1 || dateCol === -1 || countCol === -1) {
      console.error('「培訓公告」工作表找不到必要欄位（培訓名稱／培訓日期／報名人數），請確認欄位名稱是否正確');
      return;
    }

    for (let i = 1; i < annData.length; i++) {
      if (String(annData[i][nameCol]).trim() === String(trainingName).trim() && normalizeDateValue(annData[i][dateCol]) === targetDate) {
        announcementSheet.getRange(i + 1, countCol + 1).setValue(count);
        break;
      }
    }
  } catch (err) {
    console.error('updateRegistrationCount error:', err);
  }
}

// ══════════════════════════════════════
// 6. 處理延期申請
// ══════════════════════════════════════
const CANCEL_SHEET_NAME = '延期申請';
const CANCEL_HEADERS = [
  '申請時間', '姓名', '分會名稱', '原培訓名稱', '原培訓日期',
  '希望改期至', '目標場次費用', '價差（恕不退還）', '申請原因', '對應交易編號', '處理狀態',
];

// ══════════════════════════════════════
// 7. 查詢某筆報名的實際付款金額（供延期表單比對用）
// ══════════════════════════════════════
// ══════════════════════════════════════
// 8. 處理會後回饋
// ══════════════════════════════════════
const FEEDBACK_SHEET_NAME = '會後回饋';
const FEEDBACK_HEADERS = [
  '填寫時間', '姓名', '分會名稱', '培訓名稱', '培訓日期',
  '滿意度', '講義幫助程度', '印象最深的演練', '培訓建議', '提升參與建議',
  '願意擔任身份', '對應交易編號',
];

function handleSubmitFeedback(data) {
  try {
    // 核對報名紀錄：培訓名稱+日期+姓名+分會 若對得上，記錄交易編號方便追溯
    // 目前尚未全面強制系統報名，查無紀錄不擋下送出，僅在資料上註記，避免漏收真實回饋
    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();
    const targetDate = normalizeDateValue(data.trainingDate);
    let matchedTradeNo = '查無報名紀錄';

    for (let i = 1; i < regData.length; i++) {
      if (
        String(regData[i][1]).trim() === String(data.training).trim() &&
        normalizeDateValue(regData[i][2]) === targetDate &&
        String(regData[i][4]).trim() === String(data.name).trim() &&
        String(regData[i][5]).trim() === String(data.chapter).trim()
      ) {
        matchedTradeNo = regData[i][11];
        break;
      }
    }

    const sheet = getOrCreateFeedbackSheet();
    const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      now,
      data.name || '',
      data.chapter || '',
      data.training || '',
      data.trainingDate || '',
      data.satisfaction || '',
      data.materialHelp || '',
      data.memorablePractice || '',
      data.timeSuggestion || '',
      data.participationSuggestion || '',
      (data.roles || []).join('、'),
      matchedTradeNo,
    ]);

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    console.error('handleSubmitFeedback error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function getOrCreateFeedbackSheet() {
  const ss    = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(FEEDBACK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FEEDBACK_SHEET_NAME);
    sheet.appendRow(FEEDBACK_HEADERS);
    const hr = sheet.getRange(1, 1, 1, FEEDBACK_HEADERS.length);
    hr.setBackground('#1a1a2e');
    hr.setFontColor('#C9A84C');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(8, 220);
    sheet.setColumnWidth(9, 220);
    sheet.setColumnWidth(10, 220);
  }
  return sheet;
}

// ══════════════════════════════════════
// 9. 查詢某人（姓名+Email）的所有報名紀錄
// ══════════════════════════════════════
// ══════════════════════════════════════
// 10. DnA 帳密系統｜多層權限｜GPS 月會簽到
// ══════════════════════════════════════
// ⚠️ 這份名單是整個 DnA 專區帳密系統的唯一資料來源
//    DnA 成員異動時（新增/離開/密碼重設/改組織結構），只需要改這裡
//    reportsTo：直屬上層姓名，null 代表頂層（沒有上層）
//    isManager：true 代表管理者視角（例如 Steven）—— 看得到所有人，但不計入地基追蹤數據
const MEMBERS_FULL = [
  { id: 'Steven Chou',   name: 'Steven',  password: '669815', role: 'ED', reportsTo: null, isManager: true,  branches: [], hasF5: false },
  { id: 'Popo Lin',      name: '林綉蓉',  password: '856676', role: '董事顧問',     reportsTo: null, isManager: false, branches: [{name:'聚大',target:51}], hasF5: true },
  { id: 'Archie Wu',     name: '吳宗憲',  password: '212676', role: '董事顧問',     reportsTo: null, isManager: false, branches: [{name:'聚富',target:50}], hasF5: true },
  { id: 'Brenda Chen',   name: '陳虹君',  password: '541481', role: '區域培訓大使', reportsTo: null, isManager: false, branches: [], hasF5: false },
  { id: 'Shawn Chen',    name: '陳世祥',  password: '456840', role: '啟動大使',     reportsTo: '吳宗憲', isManager: false, branches: [], hasF5: false },
  { id: 'Penny Li',      name: '李佩玲',  password: '000392', role: '增長大使',     reportsTo: null, isManager: false, branches: [{name:'聚道',target:60}], hasF5: true },
  { id: 'Super Star Yu', name: '余明興',  password: '994576', role: '增長大使',     reportsTo: '吳宗憲', isManager: false, branches: [{name:'聚富',target:50}], hasF5: true },
  { id: 'Anthony Chen',  name: '陳臣勝',  password: '052580', role: '增長助理大使', reportsTo: '吳宗憲', isManager: false, branches: [{name:'聚富',target:50}], hasF5: true },
  { id: 'KK Yang',       name: '楊凱雯',  password: '032116', role: '助理大使',     reportsTo: null, isManager: false, branches: [], hasF5: false },
  { id: 'Teresa Tsai',   name: '蔡菱秝',  password: '889033', role: '助理大使',     reportsTo: null, isManager: false, branches: [], hasF5: false, startMonth: '2026-04', endMonth: '2026-06' },
  { id: 'Shun-Hao Wu',   name: '吳舜豪',  password: '574136', role: '助理大使',     reportsTo: '林綉蓉', isManager: false, branches: [], hasF5: false, startMonth: '2026-06' },
  { id: 'Dao-Ran Lin',   name: '林道然',  password: '644197', role: '啟動大使',     reportsTo: null, isManager: false, branches: [], hasF5: false, startMonth: '2026-06' },
  { id: 'I-CHEN CHIANG', name: '江宜真',  password: '614828', role: '助理大使',     reportsTo: '陳臣勝', isManager: false, branches: [], hasF5: false, startMonth: '2026-06' },
  { id: 'Bo-Lin Jiang',  name: '江柏林',  password: '399711', role: '助理大使',     reportsTo: null, isManager: false, branches: [], hasF5: false, startMonth: '2026-06', endMonth: '2026-07' },
  { id: 'Bo-Ting Chou',  name: '周柏廷',  password: '951453', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-06' },
  { id: 'One One',       name: '萬翎甄',  password: '838437', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-06' },
  { id: 'James Liao',    name: '廖灝明',  password: '787907', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Joanna Chou',   name: '周虹邑',  password: '342270', role: '助理大使',     reportsTo: '林道然', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Li-Chuan Chen', name: '陳力銓',  password: '620154', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Chu Wei Liang', name: '朱唯良',  password: '607767', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'OZ LIN',        name: '林煜珵',  password: '800644', role: '助理大使',     reportsTo: '余明興', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Jiang Tian yu', name: '江天昱',  password: '799424', role: '助理大使',     reportsTo: '林道然', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Melissa Chan',  name: '詹蕎瑀',  password: '669048', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Tina Huang',    name: '黃郁婷',  password: '413941', role: '助理大使',     reportsTo: '李佩玲', isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
  { id: 'Jeff Lin',      name: '林瑞營',  password: '640733', role: '助理大使',     reportsTo: null, isManager: false, branches: [], hasF5: false, startMonth: '2026-09' },
];

const CHECKIN_LOG_SHEET = 'DnA月會簽到記錄';
const CHECKIN_CONFIG_SHEET = 'DnA簽到設定';

// ── 登入 Token（無狀態，token = base64(姓名 + 分隔符 + 密碼)，每次請求都重新驗證）──
function makeToken(name, password) {
  return Utilities.base64Encode(name + '\u0001' + password, Utilities.Charset.UTF_8);
}
function parseToken(token) {
  try {
    const bytes = Utilities.base64Decode(token);
    const decoded = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    const idx = decoded.indexOf('\u0001');
    if (idx === -1) return null;
    return { name: decoded.substring(0, idx), password: decoded.substring(idx + 1) };
  } catch (e) {
    return null;
  }
}
const PASSWORD_SHEET = 'DnA帳號密碼';

// 從試算表讀取「姓名 → 目前密碼」，找不到工作表或該人時，退回程式碼裡的初始密碼（保底，避免表格還沒建好就整個掛掉）
function getPasswordMap() {
  const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PASSWORD_SHEET);
  if (!sheet) {
    // 第一次使用：自動建立並把程式碼裡的初始密碼寫進去，之後就以這張表為準
    sheet = ss.insertSheet(PASSWORD_SHEET);
    sheet.appendRow(['姓名', '密碼']);
    const hr = sheet.getRange(1, 1, 1, 2);
    hr.setBackground('#1a1a2e'); hr.setFontColor('#C9A84C'); hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    MEMBERS_FULL.forEach(m => sheet.appendRow([m.name, m.password]));
  }
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    if (name) map[name] = String(rows[i][1] || '').trim();
  }
  return map;
}

function getPasswordFor(name) {
  const map = getPasswordMap();
  if (map[name] !== undefined) return map[name];
  const m = MEMBERS_FULL.find(x => x.name === name);
  return m ? m.password : null; // 保底：表格裡萬一漏了這個人，退回程式碼裡的初始值
}

function setPasswordFor(name, newPassword) {
  const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PASSWORD_SHEET);
  if (!sheet) { getPasswordMap(); sheet = ss.getSheetByName(PASSWORD_SHEET); } // 確保表已存在
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === name) {
      sheet.getRange(i + 1, 2).setValue(newPassword);
      return true;
    }
  }
  sheet.appendRow([name, newPassword]); // 表裡沒有這個人，補一列
  return true;
}

function authenticate(token) {
  const parsed = parseToken(token);
  if (!parsed) return null;
  const m = MEMBERS_FULL.find(x => x.name === parsed.name);
  if (!m) return null;
  const realPassword = getPasswordFor(m.name);
  return (realPassword === parsed.password) ? m : null;
}

function handleDnaLogin(data) {
  const name = String(data.name || '').trim();
  const password = String(data.password || '').trim();
  const member = MEMBERS_FULL.find(m => m.name === name);
  if (!member || getPasswordFor(name) !== password) {
    return jsonResponse({ status: 'error', message: '帳號或密碼錯誤，請確認後再試一次。' });
  }
  return jsonResponse({
    status: 'ok',
    token: makeToken(member.name, password),
    name: member.name,
    role: member.role,
    isManager: !!member.isManager,
  });
}

function handleDnaChangePassword(data) {
  const member = authenticate(data.token);
  if (!member) return jsonResponse({ status: 'error', message: '登入已失效，請重新登入。' });

  const newPassword = String(data.newPassword || '').trim();
  if (!/^\d{4,10}$/.test(newPassword)) {
    return jsonResponse({ status: 'error', message: '新密碼請輸入 4~10 位數字。' });
  }
  setPasswordFor(member.name, newPassword);
  return jsonResponse({ status: 'ok', token: makeToken(member.name, newPassword) });
}

// 遞迴往下展開：找出這個人自己＋所有下線（多層）
function getVisibleNames(member) {
  if (member.isManager) {
    return MEMBERS_FULL.filter(m => !m.isManager).map(m => m.name);
  }
  const visible = [member.name];
  let frontier = [member.name];
  while (frontier.length > 0) {
    const next = [];
    MEMBERS_FULL.forEach(m => {
      if (m.reportsTo && frontier.includes(m.reportsTo) && !visible.includes(m.name)) {
        visible.push(m.name);
        next.push(m.name);
      }
    });
    frontier = next;
  }
  return visible;
}

// 取得某成員的完整可見清單（含角色、id等資訊，給前端渲染用）
// 這份月份清單需與 dna-index.html 的 MONTHS 保持同步（DnA 月會工作表建立新月份時，兩處都要加）
const DNA_MONTHS = [
  '2025-12','2026-01','2026-02','2026-03',
  '2026-04','2026-05','2026-06',
  '2026-07','2026-08','2026-09',
  '2026-10','2026-11','2026-12'
];

function readSheetAsMap(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const map = {};
  if (!sheet) return map;
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0] || [];
  for (let i = 1; i < rows.length; i++) {
    const rawId = String(rows[i][0] || '').trim();
    if (!rawId) continue;
    const cleanedId = rawId.replace(/\s+[\u4e00-\u9fff].*$/, '').trim();
    const rowObj = {};
    headers.forEach((h, idx) => { if (h) rowObj[h] = String(rows[i][idx] || ''); });
    map[cleanedId] = rowObj;
  }
  return map;
}

// 一次登入後，把這個人看得到的所有成員資訊＋全部月份資料＋地基4資料一次回傳
// （資料量小，一次拉完比較簡單，也不用前端一直來回問後端）
function handleDnaGetAllData(data) {
  try {
    const member = authenticate(data.token);
    if (!member) return jsonResponse({ status: 'error', message: '登入已失效，請重新登入。' });

    const visibleNames = getVisibleNames(member);
    const visibleMembers = MEMBERS_FULL.filter(m => !m.isManager && visibleNames.includes(m.name));

    const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);

    const monthData = {};
    DNA_MONTHS.forEach(month => { monthData[month] = readSheetAsMap(ss, month); });
    const f4Data = readSheetAsMap(ss, '地基4');

    const members = visibleMembers.map(m => ({
      id: m.id, name: m.name, role: m.role, branches: m.branches, hasF5: m.hasF5,
      startMonth: m.startMonth || null, endMonth: m.endMonth || null,
    }));

    return jsonResponse({
      status: 'ok', isManager: !!member.isManager, loginName: member.name,
      months: DNA_MONTHS, members: members, monthData: monthData, f4Data: f4Data,
    });
  } catch (err) {
    console.error('handleDnaGetAllData error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// Haversine 公式：計算兩組經緯度之間的距離（公尺）
function distanceMeters(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

function handleDnaCheckinV2(data) {
  try {
    const member = authenticate(data.token);
    if (!member) return jsonResponse({ status: 'error', message: '登入已失效，請重新登入。' });
    if (member.isManager) return jsonResponse({ status: 'error', message: '管理者帳號不需要簽到。' });

    const month = String(data.month || '').trim();
    const lat = Number(data.lat);
    const lng = Number(data.lng);
    if (!/^\d{4}-\d{2}$/.test(month)) return jsonResponse({ status: 'error', message: '月份格式錯誤。' });
    if (isNaN(lat) || isNaN(lng)) return jsonResponse({ status: 'error', message: '未取得您的定位資訊，請允許定位權限後再試一次。' });

    const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);

    // 讀取當月簽到設定
    let cfgSheet = ss.getSheetByName(CHECKIN_CONFIG_SHEET);
    if (!cfgSheet) {
      cfgSheet = ss.insertSheet(CHECKIN_CONFIG_SHEET);
      cfgSheet.appendRow(['月份','地點名稱','緯度','經度','允許範圍(公尺)','報到時間','遲到時間','未到時間','遲到罰款','未到罰款']);
      const hr = cfgSheet.getRange(1, 1, 1, 10);
      hr.setBackground('#1a1a2e'); hr.setFontColor('#C9A84C'); hr.setFontWeight('bold');
      cfgSheet.setFrozenRows(1);
    }
    const cfgRows = cfgSheet.getDataRange().getValues();
    const cfgHeaders = cfgRows[0];
    const monthColIdx = cfgHeaders.indexOf('月份');
    let cfg = null;
    for (let i = 1; i < cfgRows.length; i++) {
      if (String(cfgRows[i][monthColIdx]).trim() === month) {
        cfg = {};
        cfgHeaders.forEach((h, idx) => { cfg[h] = cfgRows[i][idx]; });
        break;
      }
    }
    if (!cfg) {
      return jsonResponse({ status: 'error', message: `尚未設定 ${month} 的月會地點與時間，請聯絡辦公室在「${CHECKIN_CONFIG_SHEET}」工作表新增這個月的設定。` });
    }

    const allowRadius = Number(cfg['允許範圍(公尺)']) || 150;
    const dist = distanceMeters(lat, lng, Number(cfg['緯度']), Number(cfg['經度']));
    if (dist > allowRadius) {
      return jsonResponse({ status: 'error', message: `您目前距離會場約 ${Math.round(dist)} 公尺，超過允許範圍（${allowRadius} 公尺），請到場後再簽到。` });
    }

    // 判斷準時／遲到／未到
    const now = new Date();
    const nowStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    const todayStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
    const onTimeDeadline = new Date(`${todayStr}T${cfg['報到時間']}:00+08:00`);
    const lateDeadline   = new Date(`${todayStr}T${cfg['遲到時間']}:00+08:00`);

    let attendanceStatus, fine;
    if (now <= onTimeDeadline) {
      attendanceStatus = '準時'; fine = 0;
    } else if (now <= lateDeadline) {
      attendanceStatus = '遲到'; fine = Number(cfg['遲到罰款']) || 100;
    } else {
      attendanceStatus = '嚴重遲到'; fine = Number(cfg['未到罰款']) || 200;
    }

    // 寫入當月工作表地基1
    const monthSheet = ss.getSheetByName(month);
    if (!monthSheet) return jsonResponse({ status: 'error', message: `找不到 ${month} 工作表。` });
    const data2d = monthSheet.getDataRange().getValues();
    const headers = data2d[0] || [];
    const f1Col = headers.indexOf('地基1');
    if (f1Col === -1) return jsonResponse({ status: 'error', message: `${month} 工作表找不到「地基1」欄位。` });

    let targetRow = -1;
    let alreadyChecked = false;
    for (let i = 1; i < data2d.length; i++) {
      const rawId = String(data2d[i][0] || '').trim();
      const cleanedId = rawId.replace(/\s+[\u4e00-\u9fff].*$/, '').trim();
      if (cleanedId === member.id) {
        targetRow = i + 1;
        alreadyChecked = String(data2d[i][f1Col] || '') === '1';
        break;
      }
    }
    if (targetRow === -1) {
      const newRow = new Array(headers.length).fill('');
      newRow[0] = `${member.id} ${member.name}`;
      newRow[f1Col] = '1';
      monthSheet.appendRow(newRow);
    } else {
      monthSheet.getRange(targetRow, f1Col + 1).setValue('1');
    }

    // 寫入簽到記錄（稽核用，含GPS與罰款資訊）
    let logSheet = ss.getSheetByName(CHECKIN_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(CHECKIN_LOG_SHEET);
      logSheet.appendRow(['簽到時間','月份','英文ID','姓名','出席狀態','罰款','距離會場(公尺)','緯度','經度']);
      const hr = logSheet.getRange(1, 1, 1, 9);
      hr.setBackground('#1a1a2e'); hr.setFontColor('#C9A84C'); hr.setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([nowStr, month, member.id, member.name, attendanceStatus, fine, Math.round(dist), lat, lng]);

    return jsonResponse({
      status: 'ok', name: member.name, alreadyChecked: alreadyChecked,
      attendanceStatus: attendanceStatus, fine: fine, distance: Math.round(dist),
    });
  } catch (err) {
    console.error('handleDnaCheckinV2 error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function handleLookupMyRegistrations(data) {
  try {
    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();
    const name  = String(data.name || '').trim();
    const email = String(data.email || '').trim().toLowerCase();

    const records = [];
    for (let i = 1; i < regData.length; i++) {
      const rowName  = String(regData[i][4] || '').trim();
      const rowEmail = String(regData[i][7] || '').trim().toLowerCase();
      if (rowName === name && rowEmail === email) {
        records.push({
          trainingName: regData[i][1],
          trainingDate: normalizeDateValue(regData[i][2]),
          status: regData[i][10],
          fee: regData[i][9],
          meal: regData[i][14] || '不需要',
        });
      }
    }

    return jsonResponse({ status: 'ok', records: records });
  } catch (err) {
    console.error('handleLookupMyRegistrations error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function handleLookupPaidFee(data) {
  try {
    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();
    const targetDate = normalizeDateValue(data.trainingDate);

    for (let i = 1; i < regData.length; i++) {
      if (
        String(regData[i][1]).trim() === String(data.training).trim() &&
        normalizeDateValue(regData[i][2]) === targetDate &&
        String(regData[i][4]).trim() === String(data.name).trim() &&
        String(regData[i][5]).trim() === String(data.chapter).trim()
      ) {
        return jsonResponse({ status: 'ok', fee: Number(regData[i][9]) || 0, identity: String(regData[i][8] || '') });
      }
    }
    return jsonResponse({ status: 'error', message: '查無此筆報名資料，請確認姓名與分會是否與報名時填寫的完全一致。' });
  } catch (err) {
    console.error('handleLookupPaidFee error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function normalizeDateValue(val) {
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy/M/d');
  }
  const s = String(val || '').trim();
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return m[1] + '/' + parseInt(m[2], 10) + '/' + parseInt(m[3], 10);
  return s;
}

function handleCancelRequest(data) {
  try {
    // 核對報名紀錄：培訓名稱+日期+姓名+分會 需完全對上才允許送出
    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();
    let matched = false;
    let matchedTradeNo = '';
    const targetDate = normalizeDateValue(data.trainingDate);

    for (let i = 1; i < regData.length; i++) {
      if (
        String(regData[i][1]).trim() === String(data.training).trim() &&  // 培訓名稱
        normalizeDateValue(regData[i][2]) === targetDate &&                // 培訓日期（正規化後比對）
        String(regData[i][4]).trim() === String(data.name).trim() &&      // 姓名
        String(regData[i][5]).trim() === String(data.chapter).trim()      // 分會名稱
      ) {
        matched = true;
        matchedTradeNo = regData[i][11];
        break;
      }
    }

    if (!matched) {
      return jsonResponse({ status: 'error', message: '查無此筆報名資料，請確認姓名與分會是否與報名時填寫的完全一致。' });
    }

    const sheet = getOrCreateCancelSheet();
    const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      now,
      data.name || '',
      data.chapter || '',
      data.training || '',
      data.trainingDate || '',
      data.deferTarget || '',
      data.deferTargetFee || '',
      data.priceDiff || '0',
      data.reason || '',
      matchedTradeNo,
      '待審核',
    ]);

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    console.error('handleCancelRequest error:', err);
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function getOrCreateCancelSheet() {
  const ss    = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(CANCEL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CANCEL_SHEET_NAME);
    sheet.appendRow(CANCEL_HEADERS);
    const hr = sheet.getRange(1, 1, 1, CANCEL_HEADERS.length);
    hr.setBackground('#1a1a2e');
    hr.setFontColor('#C9A84C');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(8, 220);
  }
  return sheet;
}

// ══════════════════════════════════════
// 工具函數
// ══════════════════════════════════════
function updateTradeNo(oldNo, newNo) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === oldNo) {
      sheet.getRange(i + 1, 12).setValue(newNo);
      break;
    }
  }
}

function updatePaymentStatus(tradeNo, status, paidAt) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === tradeNo) {
      sheet.getRange(i + 1, 11).setValue(status);
      if (paidAt) sheet.getRange(i + 1, 13).setValue(paidAt);
      break;
    }
  }
}

function updatePaymentUrl(tradeNo, url) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === tradeNo) {
      sheet.getRange(i + 1, 14).setValue(url);
      break;
    }
  }
}

function findRowByTradeNo(tradeNo) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === tradeNo) {
      return { trainingName: data[i][1], trainingDate: data[i][2] };
    }
  }
  return null;
}

// 取得完整訂單資訊（狀態、費用），供 APN 安全檢查使用
function findFullRowByTradeNo(tradeNo) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === tradeNo) {
      return { fee: data[i][9], status: String(data[i][10] || '') };
    }
  }
  return null;
}

function generateTradeNo() {
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmm');
  const rand  = Math.random().toString(36).substring(2, 6).toUpperCase();
  return 'BNI' + stamp + rand;
}

function computeMD5(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    str,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SETTINGS.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.SHEET_NAME);
    sheet.appendRow(HEADERS);
    const hr = sheet.getRange(1, 1, 1, HEADERS.length);
    hr.setBackground('#1a1a2e');
    hr.setFontColor('#C9A84C');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(8, 200);
    sheet.setColumnWidth(11, 120);
    sheet.setColumnWidth(12, 160);
    sheet.setColumnWidth(14, 300);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
