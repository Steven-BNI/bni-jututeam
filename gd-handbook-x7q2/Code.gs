// 廣東年會夥伴手冊 登記後台(Google Apps Script,綁定在試算表上)
// 流程:先登記(基本資料+轉帳+發票),之後再用「登記時填的 Email + 手機末四碼」補齊手冊資料與檔案。
var SHEET_NAME = '登記';
var FOLDER_NAME = '廣東年會手冊資料';
var FEATURE_LIMIT = 6;
var PRICES = { feature: 3600, cell: 700 };
var PLAN_LABEL = { feature: '一頁式曝光', cell: '一格曝光' };
var PAID = '已收款(請手動填 Y)';
var HEADERS = [
  '登記編號', '提交時間', '方案', '應付金額', PAID, '狀態',
  '中文全名', '英文名', '分會', '公司名稱', '手機', '發票寄送 Email', '轉帳後五碼', '轉帳日期',
  '發票類型', '統一編號', '發票抬頭',
  '補資料狀態', '補資料時間', '專業別', '產品或服務', '我有的資源', '我想要的資源',
  'LINE ID', '微信 ID', 'WhatsApp', 'QR 類型', '主標語', '優惠或服務內容',
  '資料夾', '形象照', 'QR 圖', 'LOGO', '廣告檔', '通知信', '提醒信', '取消信'
];
var TEXT_COLS = ['手機', '轉帳後五碼', '轉帳日期', '統一編號', 'LINE ID', '微信 ID', 'WhatsApp'];
var UPLOAD_URL = 'https://steven-bni.github.io/bni-jututeam/gd-handbook-x7q2/upload.html';
var SENDER_NAME = 'BNI 新北市西A區 聚團隊';
var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function col_(name) { return HEADERS.indexOf(name); }
function digits_(s) { return String(s || '').replace(/\D/g, ''); }
function t_(x, max) { return String(x || '').trim().slice(0, max); }
function safe_(s) { return String(s || '').replace(/[\\\/:*?"<>|\s]+/g, '').slice(0, 30); }

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (sh) {
    var have = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    if (HEADERS.slice(0, have.length).join('|') !== have.join('|')) {
      sh.setName(SHEET_NAME + '_舊_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMddHHmm'));
      sh = null;
    } else if (have.length < HEADERS.length) {
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#C8102E').setFontColor('#ffffff');
    }
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#C8102E').setFontColor('#ffffff');
    for (var i = 0; i < TEXT_COLS.length; i++) {
      sh.getRange(2, col_(TEXT_COLS[i]) + 1, 2000, 1).setNumberFormat('@');
    }
  }
  return sh;
}

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
}

function featureCount_(rows) {
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][col_('方案')] === PLAN_LABEL.feature && rows[i][col_('狀態')] !== '已取消') n++;
  }
  return n;
}

function totalCount_(rows) {
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][col_(PAID)]).trim().toUpperCase() === 'Y' && rows[i][col_('狀態')] !== '已取消') n++;
  }
  return n;
}

function nextId_(rows) {
  var max = 0;
  for (var i = 0; i < rows.length; i++) {
    var m = /^HB-(\d+)$/.exec(String(rows[i][0]));
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'HB-' + ('00' + (max + 1)).slice(-3);
}

function find_(sh, email, phone4) {
  var rows = rows_(sh);
  var e = String(email || '').trim().toLowerCase();
  if (!e || digits_(phone4).length !== 4) return null;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i][col_('狀態')] === '已取消') continue;
    if (String(rows[i][col_('發票寄送 Email')]).trim().toLowerCase() !== e) continue;
    if (digits_(rows[i][col_('手機')]).slice(-4) !== digits_(phone4)) continue;
    return { rowNum: i + 2, vals: rows[i] };
  }
  return null;
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === 'slots') {
    var rows = rows_(sheet_());
    return json_({ status: 'ok', featureLeft: Math.max(0, FEATURE_LIMIT - featureCount_(rows)), total: totalCount_(rows) });
  }
  return json_({ status: 'ok' });
}

function rootFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function idFromUrl_(url) {
  var m = /(?:folders\/|\/d\/)([\w-]+)/.exec(String(url || ''));
  return m ? m[1] : '';
}

function saveFile_(folder, f, base, type, oldUrl) {
  if (!f || !f.b64) return oldUrl || '';
  var name = base + '_' + type + '.' + (f.ext || 'bin');
  var blob = Utilities.newBlob(Utilities.base64Decode(f.b64), f.mime || 'application/octet-stream', name);
  var url = folder.createFile(blob).getUrl();
  var oldId = idFromUrl_(oldUrl);
  if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (x) {} }
  return url;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var d = JSON.parse(e.postData.contents);
    if (d.action === 'register') return register_(d);
    if (d.action === 'lookup') return lookup_(d);
    if (d.action === 'materials') return materials_(d);
    return json_({ status: 'error', message: '未知的動作' });
  } catch (err) {
    return json_({ status: 'error', message: '後台處理失敗:' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function register_(d) {
  if (!PRICES[d.plan]) return json_({ status: 'error', message: '方案不正確' });
  var must = ['nameZh', 'nameEn', 'chapter', 'company', 'phone', 'last5', 'payDate'];
  for (var i = 0; i < must.length; i++) {
    if (!String(d[must[i]] || '').trim()) return json_({ status: 'error', message: '有必填欄位沒填' });
  }
  if (digits_(d.phone).length < 8) return json_({ status: 'error', message: '手機號碼不正確' });
  if (!EMAIL_RE.test(String(d.invEmail || '').trim())) return json_({ status: 'error', message: '發票 Email 不正確' });
  if (d.invType === '3') {
    if (!/^\d{8}$/.test(String(d.taxId || '').trim()) || !String(d.invTitle || '').trim()) {
      return json_({ status: 'error', message: '三聯式發票需要 8 位統一編號與抬頭' });
    }
  } else if (d.invType !== '2') {
    return json_({ status: 'error', message: '發票類型不正確' });
  }

  var sh = sheet_();
  var rows = rows_(sh);
  if (d.plan === 'feature' && featureCount_(rows) >= FEATURE_LIMIT) {
    return json_({ status: 'error', message: '一頁式曝光 6 個名額已額滿,請改選一格曝光' });
  }
  var id = nextId_(rows);
  var folder = rootFolder_().createFolder(id + '_' + safe_(d.chapter) + '_' + safe_(d.nameZh));
  var o = {};
  o['登記編號'] = id; o['提交時間'] = new Date(); o['方案'] = PLAN_LABEL[d.plan]; o['應付金額'] = PRICES[d.plan];
  o['中文全名'] = t_(d.nameZh, 30); o['英文名'] = t_(d.nameEn, 60); o['分會'] = t_(d.chapter, 20);
  o['公司名稱'] = t_(d.company, 80); o['手機'] = t_(d.phone, 30); o['發票寄送 Email'] = t_(d.invEmail, 100);
  o['轉帳後五碼'] = t_(d.last5, 5); o['轉帳日期'] = t_(d.payDate, 20);
  o['發票類型'] = d.invType === '3' ? '三聯式' : '二聯式';
  o['統一編號'] = d.invType === '3' ? t_(d.taxId, 8) : '';
  o['發票抬頭'] = d.invType === '3' ? t_(d.invTitle, 80) : '';
  o['補資料狀態'] = '未補'; o['資料夾'] = folder.getUrl();
  sh.appendRow(HEADERS.map(function (h) { return o[h] === undefined ? '' : o[h]; }));
  return json_({ status: 'ok', id: id, amount: PRICES[d.plan] });
}

function lookup_(d) {
  var sh = sheet_();
  var r = find_(sh, d.email, d.phone4);
  if (!r) return json_({ status: 'error', message: 'Email 或手機末四碼不正確' });
  var v = r.vals;
  function g(n) { return String(v[col_(n)] || ''); }
  return json_({
    status: 'ok',
    id: g('登記編號'),
    name: g('中文全名'), chapter: g('分會'), company: g('公司名稱'),
    plan: g('方案') === PLAN_LABEL.feature ? 'feature' : 'cell',
    done: g('補資料狀態') === '已補',
    has: { photo: !!g('形象照'), qr: !!g('QR 圖'), logo: !!g('LOGO'), ad: !!g('廣告檔') },
    data: {
      trade: g('專業別'), service: g('產品或服務'), have: g('我有的資源'), want: g('我想要的資源'),
      line: g('LINE ID'), wechat: g('微信 ID'), wa: g('WhatsApp'), qrType: g('QR 類型'),
      slogan: g('主標語'), offer: g('優惠或服務內容')
    }
  });
}

function materials_(d) {
  var sh = sheet_();
  var r = find_(sh, d.email, d.phone4);
  if (!r) return json_({ status: 'error', message: 'Email 或手機末四碼不正確' });
  var v = r.vals;
  function g(n) { return String(v[col_(n)] || ''); }
  var feature = g('方案') === PLAN_LABEL.feature;
  var must = ['trade', 'service', 'have', 'want'];
  if (feature) must.push('slogan', 'offer');
  for (var i = 0; i < must.length; i++) {
    if (!String(d[must[i]] || '').trim()) return json_({ status: 'error', message: '有必填欄位沒填' });
  }
  if (!String(d.line || '').trim() && !String(d.wechat || '').trim() && !String(d.wa || '').trim()) {
    return json_({ status: 'error', message: 'LINE、微信、WhatsApp 至少填一項' });
  }
  var f = d.files || {};
  if (!f.photo && !g('形象照')) return json_({ status: 'error', message: '缺少形象照' });
  if (!f.qr && !g('QR 圖')) return json_({ status: 'error', message: '缺少 QR code 圖檔' });
  if (feature && !f.logo && !g('LOGO')) return json_({ status: 'error', message: '缺少公司 LOGO' });

  var folderId = idFromUrl_(g('資料夾'));
  var folder = folderId ? DriveApp.getFolderById(folderId) : rootFolder_();
  var base = safe_(g('分會')) + '_' + safe_(g('中文全名'));
  var vals = v.slice();
  function set(n, x) { vals[col_(n)] = x; }
  set('專業別', t_(d.trade, 20)); set('產品或服務', t_(d.service, 100));
  set('我有的資源', t_(d.have, 30)); set('我想要的資源', t_(d.want, 30));
  set('LINE ID', t_(d.line, 60)); set('微信 ID', t_(d.wechat, 60)); set('WhatsApp', t_(d.wa, 40));
  set('QR 類型', t_(d.qrType, 10));
  set('主標語', feature ? t_(d.slogan, 40) : ''); set('優惠或服務內容', feature ? t_(d.offer, 80) : '');
  set('形象照', saveFile_(folder, f.photo, base, '照片', g('形象照')));
  set('QR 圖', saveFile_(folder, f.qr, base, 'QR', g('QR 圖')));
  if (feature) {
    set('LOGO', saveFile_(folder, f.logo, base, 'LOGO', g('LOGO')));
    set('廣告檔', saveFile_(folder, f.ad, base, '廣告', g('廣告檔')));
  }
  set('補資料狀態', '已補'); set('補資料時間', new Date());
  sh.getRange(r.rowNum, 1, 1, HEADERS.length).setValues([vals]);
  return json_({ status: 'ok', id: g('登記編號') });
}

// ===== 收款確認信:在「已收款」欄填 Y,自動寄信給夥伴 =====
// 第一次使用:在 Apps Script 編輯器選 setupTrigger,按「執行」並授權。

function setupTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) {
    var fn = ts[i].getHandlerFunction();
    if (fn === 'onPaidEdit' || fn === 'dailyCheck') ScriptApp.deleteTrigger(ts[i]);
  }
  ScriptApp.newTrigger('onPaidEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  // 每天凌晨 1 點(台北時間)檢查:提醒轉帳、取消逾期名額
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(1).inTimezone('Asia/Taipei').create();
}

function onPaidEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    var pc = col_(PAID) + 1;
    if (e.range.getColumn() > pc || e.range.getLastColumn() < pc) return;
    for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
      if (r >= 2) notifyRow_(sh, r);
    }
  } catch (err) {}
}

// 補寄:若有填 Y 但沒寄出的列(例如觸發器沒跑到),在編輯器執行這個。
function sendPendingNotices() {
  var sh = sheet_();
  for (var r = 2; r <= sh.getLastRow(); r++) notifyRow_(sh, r);
}

function notifyRow_(sh, r) {
  var vals = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
  function g(n) { return String(vals[col_(n)] || '').trim(); }
  if (g(PAID).toUpperCase() !== 'Y' || g('狀態') === '已取消') return;
  var sent = g('通知信');
  if (sent && sent.indexOf('失敗') !== 0) return;
  var cell = sh.getRange(r, col_('通知信') + 1);
  var email = g('發票寄送 Email');
  if (!EMAIL_RE.test(email)) { cell.setValue('失敗:Email 不正確'); return; }

  var link = UPLOAD_URL + '?e=' + encodeURIComponent(email);
  var done = g('補資料狀態') === '已補';
  var body = g('中文全名') + ' 您好,\n\n' +
    '我們已確認收到你的款項,謝謝你參與廣東年會夥伴手冊。\n\n' +
    '登記編號:' + g('登記編號') + '\n' +
    '方案:' + g('方案') + '\n' +
    '金額:NT$ ' + g('應付金額') + '\n\n';
  if (done) {
    body += '你的手冊資料我們也已經收到。11/25 前如需修改,可以回到下面的連結更新:\n' + link + '\n\n';
  } else {
    body += '接下來請盡快補齊手冊資料(專業別、產品或服務介紹、資源、聯繫方式、照片與 QR code),11/25 前完成:\n' + link + '\n\n' +
      '進入後,請輸入你登記時填的 Email(就是這個信箱)與手機末四碼。\n\n';
  }
  body += '發票開立後,會另行寄到這個信箱。\n\n有任何問題,歡迎在群組聯繫區域窗口。\n\n' +
    'Givers Gain®\nBNI 新北市西A區 聚團隊';
  try {
    MailApp.sendEmail({ to: email, subject: '【廣東年會夥伴手冊】已收到你的款項 ' + g('登記編號') + (done ? '' : ',請盡快補資料'), body: body, name: SENDER_NAME });
    cell.setValue('已寄 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm'));
  } catch (err) {
    cell.setValue('失敗:' + String(err).slice(0, 80));
  }
}

// ===== 每日檢查:登記當天沒標 Y → 隔天凌晨寄轉帳提醒 → 再隔天凌晨取消名額 =====
// 登記日 D:D+1 凌晨寄提醒(今天 23:59 前要核對到款項);D+1 23:59 仍沒有 Y,D+2 凌晨取消。

var BANK_INFO = '國泰世華銀行(013)\n帳號:265035002662\n戶名:佳訊達國際開發有限公司';
var REGISTER_URL = 'https://steven-bni.github.io/bni-jututeam/gd-handbook-x7q2/';

function dayDiff_(a, b) {
  var da = new Date(a + 'T00:00:00Z'), db = new Date(b + 'T00:00:00Z');
  return Math.round((db - da) / 86400000);
}

function dailyCheck() {
  var sh = sheet_();
  var tz = 'Asia/Taipei';
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var rows = rows_(sh);
  var reminded = [], cancelled = [];
  for (var i = 0; i < rows.length; i++) {
    var vals = rows[i], rowNum = i + 2;
    function g(n) { return String(vals[col_(n)] || '').trim(); }
    if (g(PAID).toUpperCase() === 'Y' || g('狀態') === '已取消') continue;
    var ts = vals[col_('提交時間')];
    if (!(ts instanceof Date)) continue;
    var diff = dayDiff_(Utilities.formatDate(ts, tz, 'yyyy-MM-dd'), today);
    var label = g('登記編號') + ' ' + g('中文全名') + '(' + g('分會') + ' / ' + g('方案') + ')';
    if (diff >= 2) {
      sh.getRange(rowNum, col_('狀態') + 1).setValue('已取消');
      sendMail_(sh, rowNum, g('發票寄送 Email'), '取消信',
        '【廣東年會夥伴手冊】登記已取消 ' + g('登記編號'),
        g('中文全名') + ' 您好,\n\n' +
        '你登記的廣東年會夥伴手冊(' + g('登記編號') + ',' + g('方案') + ',NT$ ' + g('應付金額') + ')\n' +
        '因為在期限內沒有核對到款項,名額已經取消。\n\n' +
        '如果還要參與,請重新登記:\n' + REGISTER_URL + '\n\n' +
        '如果你已經完成轉帳,請在群組聯繫區域窗口確認,我們會協助處理。\n\n' +
        'Givers Gain®\n' + SENDER_NAME);
      cancelled.push(label);
    } else if (diff === 1 && !g('提醒信')) {
      sendMail_(sh, rowNum, g('發票寄送 Email'), '提醒信',
        '【廣東年會夥伴手冊】請於今天 23:59 前完成轉帳 ' + g('登記編號'),
        g('中文全名') + ' 您好,\n\n' +
        '你昨天登記了廣東年會夥伴手冊:\n' +
        '登記編號:' + g('登記編號') + '\n方案:' + g('方案') + '\n金額:NT$ ' + g('應付金額') + '\n\n' +
        '目前我們還沒有核對到你的款項。請於今天 23:59 前完成轉帳:\n' + BANK_INFO + '\n\n' +
        '區域的對帳是人工作業。如果你已經轉帳,請不用理會這封信,核對到款項後會另外通知你。\n' +
        '今天 23:59 前仍核對不到款項,名額會自動取消,需要重新登記。\n\n' +
        'Givers Gain®\n' + SENDER_NAME);
      reminded.push(label);
    }
  }
  if (reminded.length || cancelled.length) {
    var owner = Session.getEffectiveUser().getEmail();
    var body = '手冊登記每日檢查(' + today + ')\n\n' +
      '已寄轉帳提醒 ' + reminded.length + ' 人:\n' + (reminded.join('\n') || '無') + '\n\n' +
      '已取消名額 ' + cancelled.length + ' 人:\n' + (cancelled.join('\n') || '無') + '\n\n' +
      '如有人已轉帳但被取消,把該列「狀態」清空、「已收款」填 Y 即可恢復。';
    try { MailApp.sendEmail({ to: owner, subject: '【手冊登記】每日檢查 提醒' + reminded.length + ' 人、取消' + cancelled.length + ' 人', body: body, name: SENDER_NAME }); } catch (x) {}
  }
}

function sendMail_(sh, rowNum, to, statusCol, subject, body) {
  var cell = sh.getRange(rowNum, col_(statusCol) + 1);
  if (!EMAIL_RE.test(to)) { cell.setValue('失敗:Email 不正確'); return; }
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: SENDER_NAME });
    cell.setValue('已寄 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm'));
  } catch (err) {
    cell.setValue('失敗:' + String(err).slice(0, 80));
  }
}

// ===== 手動寄轉帳提醒:在試算表選取那一列,點選單「手冊登記 → 寄轉帳提醒給選取的列」 =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('手冊登記')
    .addItem('寄轉帳提醒給選取的列', 'remindSelected')
    .addToUi();
}

function remindSelected() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) { ui.alert('請先切到「' + SHEET_NAME + '」分頁,再選取要提醒的列。'); return; }
  var rg = sh.getActiveRange();
  var sent = 0, skipped = [];
  for (var r = rg.getRow(); r <= rg.getLastRow(); r++) {
    if (r < 2) continue;
    var vals = sh.getRange(r, 1, 1, HEADERS.length).getValues()[0];
    function g(n) { return String(vals[col_(n)] || '').trim(); }
    if (!g('登記編號')) continue;
    if (g(PAID).toUpperCase() === 'Y') { skipped.push(g('登記編號') + '(已收款)'); continue; }
    if (g('狀態') === '已取消') { skipped.push(g('登記編號') + '(已取消)'); continue; }
    sendMail_(sh, r, g('發票寄送 Email'), '提醒信',
      '【廣東年會夥伴手冊】請完成轉帳 ' + g('登記編號'),
      g('中文全名') + ' 您好,\n\n' +
      '你登記了廣東年會夥伴手冊:\n' +
      '登記編號:' + g('登記編號') + '\n方案:' + g('方案') + '\n金額:NT$ ' + g('應付金額') + '\n\n' +
      '目前我們還沒有核對到你的款項。請盡快完成轉帳:\n' + BANK_INFO + '\n\n' +
      '區域的對帳是人工作業。如果你已經轉帳,請不用理會這封信,核對到款項後會另外通知你。\n' +
      '登記隔天 23:59 前核對不到款項,名額會自動取消,需要重新登記。\n\n' +
      'Givers Gain®\n' + SENDER_NAME);
    if (String(sh.getRange(r, col_('提醒信') + 1).getValue()).indexOf('已寄') === 0) sent++;
  }
  ui.alert('已寄出 ' + sent + ' 封提醒。' + (skipped.length ? '\n略過:' + skipped.join('、') : ''));
}
