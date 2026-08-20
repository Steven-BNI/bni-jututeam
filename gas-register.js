/**
 * BNI 聚團隊｜培訓報名系統 - Google Apps Script
 * ════════════════════════════════════════════════
 * 功能：
 *   1. 接收前端報名資料，寫入試算表一「報名紀錄」工作表
 *   2. 付費報名：產生 ECPay CheckMacValue 回傳給前端
 *   3. 免費報名：直接回傳成功
 *   4. ECPay 付款結果接收（ReturnURL）→ 更新付款狀態
 *
 * 部署方式：
 *   擴充功能 → Apps Script → 貼上此程式碼
 *   → 部署 → 新增部署 → 網頁應用程式
 *   → 執行身分：我（你的帳號）
 *   → 存取權：所有人
 *   → 複製「網頁應用程式網址」貼到 register.html 的 CONFIG.GAS_URL
 * ════════════════════════════════════════════════
 */

// ══════════════════════════════════════
// ⚙️  設定區（請填入你的資料）
// ══════════════════════════════════════
const SETTINGS = {
  // 試算表一的 ID
  SPREADSHEET_ID: '1JKHpemlHQd_iCp1FpRwyolwL1oL3LNNKtvd5AVEZ-eI',

  // 「報名紀錄」工作表名稱（不存在時自動建立）
  SHEET_NAME: '報名紀錄',

  // 統一金流 ECPay 金鑰（從廠商後台取得，只能放後端）
  ECPAY_MERCHANT_ID: 'YOUR_MERCHANT_ID',
  ECPAY_HASH_KEY:    'YOUR_HASH_KEY',
  ECPAY_HASH_IV:     'YOUR_HASH_IV',

  // 付款完成後 ECPay 呼叫此 GAS 的 URL（ReturnURL，需與前端一致）
  // 注意：ECPay ReturnURL 必須是可接收 POST 的公開網址
  // GitHub Pages 無法接收 POST，故 ReturnURL 指向本 GAS
  // 前端的 ClientBackURL 才是使用者看到的「返回網頁」
  GAS_RETURN_URL: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
};

// ══════════════════════════════════════
// 欄位定義（報名紀錄工作表）
// ══════════════════════════════════════
const HEADERS = [
  '報名時間', '培訓名稱', '培訓日期', '地點',
  '姓名', '分會名稱', '電話', 'Email',
  '報名身份', '費用', '付款狀態', '交易編號', '付款時間', '餐點',
];

// ══════════════════════════════════════
// 主入口：處理 GET / POST 請求
// ══════════════════════════════════════
function doPost(e) {
  try {
    const raw = e.postData.contents;
    const data = JSON.parse(raw);

    // ECPay 付款結果回傳（含 RtnCode 欄位）
    if (data.RtnCode !== undefined || e.parameter.RtnCode !== undefined) {
      return handleECPayReturn(e);
    }

    // 前端報名資料
    switch (data.action) {
      case 'register': return handleRegistration(data);
      default:         return jsonResponse({ status: 'error', message: '未知 action' });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  // ECPay 有時用 GET 回傳（ClientBackURL）
  return ContentService.createTextOutput('OK');
}

// ══════════════════════════════════════
// 1. 處理報名
// ══════════════════════════════════════
function handleRegistration(data) {
  const sheet = getOrCreateSheet();

  // 產生唯一交易編號（ECPay 限 20 字元以內，僅英數字）
  const tradeNo = generateTradeNo();

  // 寫入試算表
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const identityLabel = { general:'一般學員', mentor:'認證導師', staff:'講師/籌劃小組' }[data.identity] || data.identity;

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
    data.fee > 0 ? '待付款' : '免費（已完成）',
    tradeNo,
    '',  // 付款時間（付款後更新）
    data.meal || '不需要',
  ]);

  // 同步更新「培訓公告」工作表的報名人數
  updateRegistrationCount(data.trainingName, data.trainingDate);

  // 免費：直接回傳成功
  if (data.fee === 0) {
    return jsonResponse({ status: 'ok', free: true });
  }

  // 付費：產生 ECPay 參數（含 CheckMacValue）
  const ecpayParams = buildECPayParams({
    tradeNo:     tradeNo,
    totalAmount: data.fee,
    itemName:    data.trainingName,
    tradeDesc:   'BNI培訓報名',
  });

  return jsonResponse({ status: 'ok', free: false, ecpay: ecpayParams });
}

// ══════════════════════════════════════
// 2. 同步報名人數到「培訓公告」工作表
// ══════════════════════════════════════
function updateRegistrationCount(trainingName, trainingDate) {
  try {
    const ss = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
    const announcementSheet = ss.getSheetByName('培訓公告');
    if (!announcementSheet) return;

    const regSheet = getOrCreateSheet();
    const regData  = regSheet.getDataRange().getValues();

    // 計算該培訓的付款中+已完成人數（排除「待付款」逾期未付者視需求調整）
    let count = 0;
    for (let i = 1; i < regData.length; i++) {
      if (regData[i][1] === trainingName && regData[i][2] === trainingDate) {
        const status = regData[i][10];
        if (status !== '待付款') count++;  // 免費已完成 or 已付款
      }
    }

    // 找到對應列更新「報名人數」欄（第4欄，index 3）
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
// 3. 接收 ECPay 付款結果（ReturnURL）
// ══════════════════════════════════════
function handleECPayReturn(e) {
  try {
    // ECPay 用 application/x-www-form-urlencoded 格式 POST
    const params = e.parameter;
    const rtnCode    = params.RtnCode;
    const tradeNo    = params.MerchantTradeNo;
    const paidAmount = params.TradeAmt;
    const paidAt     = params.PaymentDate || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    // 驗證 CheckMacValue
    if (!verifyCheckMacValue(params)) {
      return ContentService.createTextOutput('0|CheckMacValue Error');
    }

    if (rtnCode === '1') {
      // 付款成功：更新試算表
      updatePaymentStatus(tradeNo, '已付款', paidAt);
      // 同步報名人數（付款後才計入）
      const rowData = findRowByTradeNo(tradeNo);
      if (rowData) updateRegistrationCount(rowData.trainingName, rowData.trainingDate);
    } else {
      updatePaymentStatus(tradeNo, '付款失敗', paidAt);
    }

    // ECPay 要求回傳 "1|OK"
    return ContentService.createTextOutput('1|OK');
  } catch (err) {
    console.error('handleECPayReturn error:', err);
    return ContentService.createTextOutput('0|Error');
  }
}

function updatePaymentStatus(tradeNo, status, paidAt) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === tradeNo) {                      // 第12欄：交易編號
      sheet.getRange(i + 1, 11).setValue(status);       // 第11欄：付款狀態
      sheet.getRange(i + 1, 13).setValue(paidAt);       // 第13欄：付款時間
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

// ══════════════════════════════════════
// ECPay 工具函數
// ══════════════════════════════════════

/**
 * 建立 ECPay 付款參數（含 CheckMacValue）
 */
function buildECPayParams({ tradeNo, totalAmount, itemName, tradeDesc }) {
  const tradeDate = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  const params = {
    MerchantID:        SETTINGS.ECPAY_MERCHANT_ID,
    MerchantTradeNo:   tradeNo,
    MerchantTradeDate: tradeDate,
    PaymentType:       'aio',
    TotalAmount:       totalAmount,
    TradeDesc:         tradeDesc,
    ItemName:          itemName,
    ReturnURL:         SETTINGS.GAS_RETURN_URL,
    ChoosePayment:     'ALL',
    EncryptType:       '1',
  };

  params.CheckMacValue = generateCheckMacValue(params);
  return params;
}

/**
 * 產生 ECPay CheckMacValue（SHA256）
 */
function generateCheckMacValue(params) {
  // 1. 依參數名稱英文字母排序
  const sorted = Object.keys(params)
    .filter(k => k !== 'CheckMacValue')
    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  // 2. 組成字串
  let raw = `HashKey=${SETTINGS.ECPAY_HASH_KEY}`;
  sorted.forEach(k => { raw += `&${k}=${params[k]}`; });
  raw += `&HashIV=${SETTINGS.ECPAY_HASH_IV}`;

  // 3. URL Encode（ECPay 規則）
  raw = encodeURIComponentECPay(raw);

  // 4. 轉小寫
  raw = raw.toLowerCase();

  // 5. SHA256
  const bytes  = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  const hex    = bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
  return hex.toUpperCase();
}

/**
 * 驗證 ECPay 回傳的 CheckMacValue
 */
function verifyCheckMacValue(params) {
  const received = params.CheckMacValue;
  const copy     = Object.assign({}, params);
  delete copy.CheckMacValue;
  const expected = generateCheckMacValue(copy);
  return received === expected;
}

/**
 * ECPay 專用 URL Encode
 * 將 .、-、_、* 等特殊字元保持不編碼
 */
function encodeURIComponentECPay(str) {
  return encodeURIComponent(str)
    .replace(/%20/g, '+')
    .replace(/!/g, '%21')
    .replace(/'/g, '%27')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
    .replace(/\*/g, '%2A')
    .replace(/%7E/g, '~');
}

/**
 * 產生唯一交易編號（英數字，最多20碼）
 * 格式：BNI + yyMMddHHmm + 4位隨機
 */
function generateTradeNo() {
  const now    = new Date();
  const stamp  = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMddHHmm');
  const rand   = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `BNI${stamp}${rand}`;
}

// ══════════════════════════════════════
// 試算表工具
// ══════════════════════════════════════
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SETTINGS.SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SETTINGS.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.SHEET_NAME);
    // 建立表頭
    sheet.appendRow(HEADERS);
    // 格式化表頭
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#1a1a2e');
    headerRange.setFontColor('#C9A84C');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 欄寬
    sheet.setColumnWidth(1, 160);  // 報名時間
    sheet.setColumnWidth(2, 200);  // 培訓名稱
    sheet.setColumnWidth(5, 100);  // 姓名
    sheet.setColumnWidth(6, 140);  // 分會名稱
    sheet.setColumnWidth(8, 200);  // Email
    sheet.setColumnWidth(11, 120); // 付款狀態
    sheet.setColumnWidth(12, 180); // 交易編號
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
