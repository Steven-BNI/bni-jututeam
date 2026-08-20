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
// 欄位定義
// ══════════════════════════════════════
const HEADERS = [
  '報名時間', '培訓名稱', '培訓日期', '地點',
  '姓名', '分會名稱', '電話', 'Email',
  '報名身份', '費用', '付款狀態', '交易編號', '付款時間', '付款網址',
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
      case 'register': return handleRegistration(data);
      default:         return jsonResponse({ status: 'error', message: '未知 action' });
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
  const tradeNo = generateTradeNo();
  const now     = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const identityLabel = {
    general: '個人報名',
    mentor:  '認證導師',
    staff:   '講師/籌劃小組'
  }[data.identity] || data.identity;

  // 寫入試算表
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
  ]);

  // 同步培訓公告報名人數
  updateRegistrationCount(data.trainingName, data.trainingDate);

  // 免費：直接回傳成功
  if (data.fee === 0) {
    return jsonResponse({ status: 'ok', free: true });
  }

  // ATM 轉帳：記錄末五碼，不建立刷卡訂單
  if (data.payMethod === 'atm') {
    updatePaymentStatus(tradeNo, 'ATM待確認（末五碼：' + (data.atmLast5||'未填') + '）', '');
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
// 4. 接收黑貓PAY APN 付款通知
// ══════════════════════════════════════
function handlePayUniAPN(data) {
  try {
    // Log 完整回傳內容以便 debug
    console.log('APN 收到：' + JSON.stringify(data));

    const tradeNo = data.cust_order_no || data.order_no || data.OrderNo || data.MerchantOrderNo || '';
    const paidAt  = data.pay_date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    if (!tradeNo) {
      console.error('APN 沒有交易編號');
      return ContentService.createTextOutput('ERROR');
    }

    // 統一金流付款成功判斷：status = 'D' 或 'S' 或 success = true 或 result_code = '00'
    const isSuccess = (
      data.status === 'D' ||
      data.status === 'S' ||
      data.success === true ||
      data.success === 'true' ||
      data.result_code === '00' ||
      data.ResCode === '00' ||
      data.RtnCode === 1 ||
      String(data.RtnCode) === '1'
    );

    if (isSuccess) {
      updatePaymentStatus(tradeNo, '已付款', paidAt);
      const rowData = findRowByTradeNo(tradeNo);
      if (rowData) updateRegistrationCount(rowData.trainingName, rowData.trainingDate);
      console.log('付款成功：' + tradeNo);
    } else {
      updatePaymentStatus(tradeNo, '付款失敗', paidAt);
      console.log('付款失敗：' + tradeNo + ' status=' + data.status);
    }

    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('handlePayUniAPN error:', err);
    return ContentService.createTextOutput('ERROR');
  }
}

// ══════════════════════════════════════
// 5. 同步報名人數到「培訓公告」工作表
// ══════════════════════════════════════
function updateRegistrationCount(trainingName, trainingDate) {
  try {
    const ss               = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
    const announcementSheet = ss.getSheetByName('培訓公告');
    if (!announcementSheet) return;

    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();

    let count = 0;
    for (let i = 1; i < regData.length; i++) {
      if (regData[i][1] === trainingName && regData[i][2] === trainingDate) {
        const st = regData[i][10];
        if (st === '已付款' || st === '免費（已完成）') count++;
      }
    }

    const annData = announcementSheet.getDataRange().getValues();
    for (let i = 1; i < annData.length; i++) {
      if (annData[i][1] === trainingName && annData[i][0] === trainingDate) {
        announcementSheet.getRange(i + 1, 4).setValue(count);
        break;
      }
    }
  } catch (err) {
    console.error('updateRegistrationCount error:', err);
  }
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
