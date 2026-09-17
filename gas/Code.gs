/**
 * Floralway 線上預約系統 — Google Apps Script 後端
 *
 * 這個檔案要貼到「【勿刪除】Floralway預約系統_資料庫」試算表的
 * 擴充功能 > Apps Script 綁定指令碼專案裡（不是獨立專案）。
 *
 * 部署方式：部署 > 新增部署作業 > 類型選「網頁應用程式」
 *   - 執行身分：我（你自己的帳號）
 *   - 具有應用程式存取權的使用者：所有人
 * 部署後會拿到一個 exec 結尾的網址，那就是前端 LIFF 表單要打的 API 網址。
 *
 * 使用前必須先在「專案設定 > 指令碼屬性」加入：
 *   LINE_CHANNEL_ACCESS_TOKEN = (Messaging API channel access token)
 *   CALENDAR_ID = (要同步忙碌時段的 Google 日曆 ID，留空則用你的預設日曆)
 *   PAYMENT_BANK_NAME = (匯款銀行名稱，會即時顯示在預約頁面付款方式底下)
 *   PAYMENT_BANK_ACCOUNT = (匯款帳號，會即時顯示在預約頁面付款方式底下)
 *
 * 上面兩個 PAYMENT_ 開頭的屬性只會存在「指令碼屬性」裡，絕對不要寫進這個檔案
 * 或任何會上傳到公開 GitHub repo 的檔案，避免帳號外流。
 */

// ===================== 基本設定 =====================

var SHEET_NAMES = {
  SERVICES: '服務設定',
  BOOKINGS: '預約紀錄',
  NOTIFICATIONS: '通知記錄'
};

var BOOKING_COLUMNS = [
  '預約編號', '姓名', '電話', '生日', '服務區域', '服務項目', '訂金金額',
  '預約日期', '預約時段', '付款方式', '匯款後五碼', 'LINE_userId', 'LINE顯示名稱',
  '填單時間', '狀態', '確認到帳時間', '日曆事件ID',
  '提醒_保留中已發送', '提醒_未完成2', '提醒_行前7天', '提醒_行前3天', '備註'
];

// 2026-09-17 起：「是否已加入客戶日曆」欄改為「日曆事件ID」（存放保留時段用的日曆事件ID，
// 方便之後自動取消／更新），「提醒_未完成1」欄改為「提醒_保留中已發送」（改用24小時保留機制，
// 詳見下方「訂金保留與自動釋出」章節）。既有試算表的欄位標題需要跟著改名，程式才讀得到正確欄位。

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  var sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁：' + name);
  return sh;
}

function getScriptProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

// ===================== HTTP 進入點 =====================

function doGet(e) {
  try {
    var action = e.parameter.action || '';
    var result;
    if (action === 'services') {
      result = getActiveServices_();
    } else if (action === 'busyslots') {
      result = getBusySlots_(e.parameter.start, e.parameter.end);
    } else if (action === 'paymentinfo') {
      result = getPaymentInfo_();
    } else {
      result = { error: '未知的 action：' + action };
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'submit';
    var result;
    if (action === 'submit') {
      result = submitBooking_(body);
    } else {
      result = { success: false, error: '未知的 action：' + action };
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err) });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== 服務設定 =====================

/**
 * 回傳所有「是否上架」= TRUE 的服務項目
 * [{code, name, desc, price, deposit, durationMinutes, regions, requires}]
 */
function getActiveServices_() {
  var sh = getSheet_(SHEET_NAMES.SERVICES);
  var values = sh.getDataRange().getValues();
  var header = values[0];
  var idx = {};
  header.forEach(function (h, i) { idx[h] = i; });

  var list = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[idx['服務代碼']]) continue;
    var onShelf = row[idx['是否上架']];
    if (onShelf !== true && String(onShelf).toUpperCase() !== 'TRUE') continue;

    list.push({
      code: row[idx['服務代碼']],
      name: row[idx['服務名稱']],
      desc: row[idx['說明文字']] || '',
      price: row[idx['服務價格']] || '',
      deposit: row[idx['訂金金額']] || 0,
      durationMinutes: row[idx['服務時長（分鐘）']] || 0,
      regions: row[idx['可服務區域']] || '',
      requires: row[idx['必選搭配服務代碼']] || ''
    });
  }
  return { services: list };
}

// ===================== 匯款資訊（即時從指令碼屬性讀取） =====================
// 帳戶名稱/帳號只存在 Apps Script 的「指令碼屬性」，不會寫進這個檔案，
// 前端會在使用者選好付款方式後，用這個 action 即時打API拿資料顯示。

function getPaymentInfo_() {
  return {
    bankName: getScriptProp_('PAYMENT_BANK_NAME') || '',
    bankAccount: getScriptProp_('PAYMENT_BANK_ACCOUNT') || ''
  };
}

// ===================== 日曆忙碌時段 =====================

function getCalendar_() {
  var calId = getScriptProp_('CALENDAR_ID');
  if (calId) {
    var cal = CalendarApp.getCalendarById(calId);
    if (cal) return cal;
  }
  return CalendarApp.getDefaultCalendar();
}

/**
 * 回傳指定日期區間內的忙碌時段，以及週日整天視為忙碌
 * start/end: 'YYYY-MM-DD'
 */
function getBusySlots_(start, end) {
  if (!start || !end) {
    return { error: '缺少 start 或 end 參數' };
  }
  var cal = getCalendar_();
  var startDate = new Date(start + 'T00:00:00');
  var endDate = new Date(end + 'T23:59:59');
  var events = cal.getEvents(startDate, endDate);

  var busy = events.map(function (ev) {
    return {
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString(),
      allDay: ev.isAllDayEvent()
    };
  });

  // 週日整天列為不可預約
  var sundays = [];
  var cursor = new Date(startDate);
  while (cursor <= endDate) {
    if (cursor.getDay() === 0) {
      var dayStart = new Date(cursor);
      dayStart.setHours(0, 0, 0, 0);
      var dayEnd = new Date(cursor);
      dayEnd.setHours(23, 59, 59, 999);
      sundays.push({ start: dayStart.toISOString(), end: dayEnd.toISOString(), allDay: true });
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  return { busy: busy.concat(sundays) };
}

// ===================== 日曆保留（Hold）輔助函式 =====================
// 訂金 > 0 的預約，送出當下就會在日曆上建立一個「保留」事件，讓這個時段
// 立刻對其他客戶顯示為忙碌（不用等你手動確認）；24小時內沒收到款項會
// 自動刪除這個事件、釋出時段（詳見下方「訂金保留與自動釋出」章節）。
// 固定用90分鐘計算保留時長，跟前端「busyslots」判斷可預約時段時的假設一致。

function createHoldEvent_(name, serviceDisplayNames, dateStr, timeStr) {
  try {
    if (!dateStr || !timeStr) return '';
    var start = new Date(dateStr + 'T' + timeStr + ':00');
    if (isNaN(start.getTime())) return '';
    var end = new Date(start.getTime() + 90 * 60000);
    var cal = getCalendar_();
    var ev = cal.createEvent('[保留] ' + name + '－' + serviceDisplayNames, start, end);
    return ev.getId();
  } catch (err) {
    Logger.log('建立日曆保留事件失敗: ' + err);
    return '';
  }
}

/** 訂金確認到帳後，把日曆事件標題的「[保留]」拿掉，代表這是已確定的預約 */
function markCalendarEventConfirmed_(eventId, name, serviceDisplayNames) {
  if (!eventId) return;
  try {
    var cal = getCalendar_();
    var ev = cal.getEventById(eventId);
    if (ev) ev.setTitle(name + '－' + serviceDisplayNames);
  } catch (err) {
    Logger.log('更新日曆事件標題失敗: ' + err);
  }
}

/** 保留逾時釋出時，刪除對應的日曆保留事件 */
function deleteHoldEvent_(eventId) {
  if (!eventId) return;
  try {
    var cal = getCalendar_();
    var ev = cal.getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (err) {
    Logger.log('刪除日曆保留事件失敗: ' + err);
  }
}

// ===================== 預約送出 =====================

function generateBookingId_() {
  var now = new Date();
  var stamp = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyMMddHHmmss');
  var rand = Math.floor(Math.random() * 90 + 10); // 2位數
  return 'B' + stamp + rand;
}

/**
 * body 預期格式：
 * {
 *   name, phone, birthday,               // 姓名/電話/生日 (birthday: 'YYYY-MM-DD')
 *   regions: ['台中', ...],               // 服務區域（可複選）
 *   services: ['COLOR2', 'PRECONSULT'],   // 服務項目代碼陣列
 *   date: 'YYYY-MM-DD',
 *   time: 'HH:mm',
 *   paymentMethod: '轉帳',   // 2026-09-16 起僅保留轉帳（整筆付清）一種付款方式，分期付款選項已移除
 *   last5Digits,             // 匯款帳號後五碼，選填。填了視為「已付款」，沒填視為「未付款」
 *                            // ——訂金>0時兩種情況會分別發不同的LINE訊息（見下方），且未付款的
 *                            // 會進入「保留24小時，逾時自動釋出」機制，詳見「訂金保留與自動釋出」章節
 *   lineUserId, lineDisplayName,
 *   note
 * }
 */
function submitBooking_(body) {
  if (!body.name || !body.phone || !body.services || body.services.length === 0) {
    return { success: false, error: '缺少必要欄位（姓名/電話/服務項目）' };
  }

  // 業務規則檢查：選了年度陪跑（COACHING）必須同時選預諮詢（PRECONSULT）
  var servicesData = getActiveServices_().services;
  var requiresMap = {};
  var serviceNameMap = {};
  servicesData.forEach(function (s) {
    requiresMap[s.code] = s.requires;
    serviceNameMap[s.code] = s.name;
  });

  for (var i = 0; i < body.services.length; i++) {
    var code = body.services[i];
    var req = requiresMap[code];
    if (req && body.services.indexOf(req) === -1) {
      return {
        success: false,
        error: '選擇「' + code + '」時，必須同時選擇「' + req + '」'
      };
    }
  }

  // 週日禁止預約
  if (body.date) {
    var d = new Date(body.date + 'T00:00:00');
    if (d.getDay() === 0) {
      return { success: false, error: '週日無法預約，請選擇其他日期' };
    }
  }

  // 計算訂金總額
  var depositMap = {};
  servicesData.forEach(function (s) { depositMap[s.code] = Number(s.deposit) || 0; });
  var totalDeposit = body.services.reduce(function (sum, code) {
    return sum + (depositMap[code] || 0);
  }, 0);

  var serviceDisplayNames = body.services.map(function (code) {
    return serviceNameMap[code] || code;
  }).join('、');

  var bookingId = generateBookingId_();
  var now = new Date();
  var last5 = body.last5Digits || '';

  // 訂金 > 0 才需要在日曆上建立保留時段（訂金 $0 的預約，例如單選線上預諮詢，
  // 是走 Calendly 另外約時段，不使用這裡的日期/時段欄位，也不需要保留機制）
  var eventId = '';
  if (totalDeposit > 0) {
    eventId = createHoldEvent_(body.name, serviceDisplayNames, body.date, body.time);
  }

  var row = {
    '預約編號': bookingId,
    '姓名': body.name,
    '電話': body.phone,
    '生日': body.birthday || '',
    '服務區域': (body.regions || []).join('、'),
    '服務項目': body.services.join('、'),
    '訂金金額': totalDeposit,
    '預約日期': body.date || '',
    '預約時段': body.time || '',
    '付款方式': body.paymentMethod || '',
    '匯款後五碼': last5,
    'LINE_userId': body.lineUserId || '',
    'LINE顯示名稱': body.lineDisplayName || '',
    '填單時間': now,
    '狀態': '待確認',
    '確認到帳時間': '',
    '日曆事件ID': eventId,
    '提醒_保留中已發送': '',
    '提醒_未完成2': '',
    '提醒_行前7天': '',
    '提醒_行前3天': '',
    '備註': body.note || ''
  };

  var sh = getSheet_(SHEET_NAMES.BOOKINGS);
  var newRow = BOOKING_COLUMNS.map(function (col) { return row[col]; });
  sh.appendRow(newRow);

  // 訂金 > 0 才需要送出「收到預約申請」的即時LINE通知；依「後五碼是否已填」分兩種內容
  if (totalDeposit > 0 && body.lineUserId) {
    var dateStr = formatDateValue_(body.date);
    var bankName = getScriptProp_('PAYMENT_BANK_NAME') || '';
    var bankAccount = getScriptProp_('PAYMENT_BANK_ACCOUNT') || '';
    var msg;
    var notifyType;

    if (last5) {
      // 情境B：送出時已填後五碼，視為已付款，等待人工核對入帳
      msg = '嗨 ' + body.name + '，我們已收到您的預約申請與匯款資料！\n\n' +
        '✦ 預約項目：' + serviceDisplayNames + '\n' +
        '✦ 預約時間：' + dateStr + ' ' + body.time + '\n' +
        '✦ 填寫後五碼：' + last5 + '\n\n' +
        '我們會在核對定金入帳後，盡快發送「預約正式成立」通知給您，請稍候我們的確認訊息喔！';
      notifyType = '送出預約通知（已填後五碼）';
    } else {
      // 情境A：送出時未填後五碼，視為未付款，時段保留24小時
      var deadline = new Date(now.getTime() + 24 * 60 * 60 * 1000);
      var deadlineStr = Utilities.formatDate(deadline, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
      msg = '嗨 ' + body.name + '，我們已收到您的預約申請！\n\n' +
        '✦ 預約項目：' + serviceDisplayNames + '\n' +
        '✦ 預約時間：' + dateStr + ' ' + body.time + '\n' +
        '✦ 保留期限：至 ' + deadlineStr + ' 止\n\n' +
        '【定金匯款資訊】\n' +
        '✦ 銀行：' + bankName + '\n' +
        '✦ 帳號：' + bankAccount + '\n' +
        '✦ 定金金額：' + totalDeposit + ' 元\n\n' +
        '完成轉帳後，請直接回覆此訊息提供「帳號後五碼」，我們確認入帳後會立即為您正式保留時段✨';
      notifyType = '送出預約通知（未填後五碼，開始24小時保留）';
    }

    var ok = sendLinePush_(body.lineUserId, msg);
    logNotification_(bookingId, body.name, body.lineUserId, notifyType, ok, ok ? '' : '發送失敗');
  }

  return { success: true, bookingId: bookingId, totalDeposit: totalDeposit };
}

// ===================== 手動確認到帳 → 觸發LINE通知 =====================
// 當你在「預約紀錄」表把某一列的「狀態」欄手動改成「已確認」時，
// 這個簡易觸發器會自動發送LINE通知給客戶，並記錄確認到帳時間。
// 注意：簡易觸發器 onEdit 不能發出外部網路請求需要的完整權限時，
// 若無法自動觸發，請改用選單「Floralway工具 > 確認選取的預約」手動執行 confirmSelectedBooking()。

function onEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAMES.BOOKINGS) return;

    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var statusCol = header.indexOf('狀態') + 1;
    if (e.range.getColumn() !== statusCol) return;

    var newValue = e.range.getValue();
    if (newValue !== '已確認') return;

    var row = e.range.getRow();
    confirmBookingRow_(sh, row, header);
  } catch (err) {
    Logger.log('onEdit 錯誤: ' + err);
  }
}

/** 選單用：確認目前選取儲存格所在的那一列預約 */
function confirmSelectedBooking() {
  var sh = getSheet_(SHEET_NAMES.BOOKINGS);
  var row = SpreadsheetApp.getActiveRange().getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('請選取資料列，而不是標題列');
    return;
  }
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  confirmBookingRow_(sh, row, header);
  SpreadsheetApp.getUi().alert('已確認並發送LINE通知');
}

function confirmBookingRow_(sh, row, header) {
  var idx = {};
  header.forEach(function (h, i) { idx[h] = i; });
  var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

  var bookingId = values[idx['預約編號']];
  var lineUserId = values[idx['LINE_userId']];
  var name = values[idx['姓名']];
  var serviceNames = values[idx['服務項目']];
  var date = values[idx['預約日期']];
  var time = values[idx['預約時段']];

  // 狀態欄若還不是「已確認」就設定它（手動執行選單時用）
  sh.getRange(row, idx['狀態'] + 1).setValue('已確認');
  sh.getRange(row, idx['確認到帳時間'] + 1).setValue(new Date());
  markCalendarEventConfirmed_(values[idx['日曆事件ID']], name, serviceNames);

  if (lineUserId) {
    var dateStr = formatDateValue_(date);
    var region = values[idx['服務區域']];
    var location = region && String(region).trim() ? region : '線上（Google Meet，會議連結將另行提供）';
    var msg = '您好 ' + name + '，您的預約已確認！\n\n' +
      '服務項目：' + serviceNames + '\n' +
      '預約日期：' + dateStr + '\n' +
      '預約時段：' + time + '\n' +
      '地點：' + location + '\n\n' +
      '是否要將此預約加入行事曆？\n' +
      '(之後會提供加入 Google/Apple 行事曆的連結)';
    sendLinePush_(lineUserId, msg);
    logNotification_(bookingId, name, lineUserId, '預約確認通知', true, '');
  }
}

// ===================== 直接在「確認到帳時間」填日期 → 自動確認並回覆客戶 =====================
// 用途：你收到客戶用LINE回報的「匯款後五碼」、對完帳之後，
// 直接在「預約紀錄」分頁那一列的「確認到帳時間」欄位填入到帳日期，
// 這個安裝式觸發器會自動：
// 1) 把狀態改成「已確認」（如果還不是的話）
// 2) 用LINE回覆客戶：已完成預約、預約日期、服務項目、地點
//
// 設定方式（部署時只需設定一次）：
// 1. 開啟這個 Apps Script 專案 → 左側「觸發條件」（鬧鐘圖示）→ 右下角「新增觸發條件」
// 2. 執行的函式：onDepositTimeEdit
// 3. 事件來源：來自試算表
// 4. 事件類型：編輯時
// 5. 儲存
// （簡易觸發器 onEdit 沒有權限呼叫LINE API，所以要另外用這個安裝式觸發器；
//   跟最上面的 onEdit 是分開的兩個函式，不會互相干擾）

function onDepositTimeEdit(e) {
  try {
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAMES.BOOKINGS) return;

    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var idx = {};
    header.forEach(function (h, i) { idx[h] = i; });

    var depositCol = idx['確認到帳時間'] + 1;
    if (e.range.getColumn() !== depositCol) return;

    var newValue = e.range.getValue();
    if (!newValue) return;

    var row = e.range.getRow();
    if (row === 1) return;

    var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    var bookingId = values[idx['預約編號']];
    var lineUserId = values[idx['LINE_userId']];
    var name = values[idx['姓名']];
    var serviceNames = values[idx['服務項目']];
    var date = values[idx['預約日期']];
    var time = values[idx['預約時段']];
    var region = values[idx['服務區域']];

    if (values[idx['狀態']] !== '已確認') {
      sh.getRange(row, idx['狀態'] + 1).setValue('已確認');
    }
    markCalendarEventConfirmed_(values[idx['日曆事件ID']], name, serviceNames);

    if (lineUserId) {
      var dateStr = formatDateValue_(date);
      var location = region && String(region).trim() ? region : '線上（Google Meet，會議連結將另行提供）';
      var msg = '您好 ' + name + '，已收到您的訂金，預約完成！\n\n' +
        '服務項目：' + serviceNames + '\n' +
        '預約日期：' + dateStr + '\n' +
        '預約時段：' + time + '\n' +
        '地點：' + location + '\n\n' +
        '期待與您見面，如有任何問題歡迎透過LINE與我們聯繫。';
      sendLinePush_(lineUserId, msg);
      logNotification_(bookingId, name, lineUserId, '到帳確認通知（填寫確認到帳時間觸發）', true, '');
    } else {
      logNotification_(bookingId, name, '', '到帳確認通知（未綁定LINE，未發送）', false, '客戶未綁定LINE');
    }
  } catch (err) {
    Logger.log('onDepositTimeEdit 錯誤: ' + err);
  }
}

// ===================== 訂金保留與自動釋出（2026-09-17 起） =====================
// 訂金 > 0 且送出時「未填後五碼」的預約，送出當下已經發送過第1則通知（見 submitBooking_），
// 接下來由這個函式接手處理：
//   ✦ 滿12小時，且「提醒_保留中已發送」還是空的 → 發送第2則「保留中」提醒
//     （若這個時間點落在晚上11點～隔天上午9點之間，會延後到隔天上午9點後才發送，避免半夜打擾客戶）
//   ✦ 滿24小時 → 不論有沒有綁定LINE，都會：
//       1) 刪除日曆上的保留事件（真正把時段空出來給其他人預約）
//       2) 狀態改成「已釋出」
//       3) 有綁定LINE的話，發送第3則「已釋出」通知
// 規則：
//   1) 訂金金額 > 0 才進入這套流程（0元訂金，例如單選線上預諮詢，不受影響）
//   2) 還沒有「確認到帳時間」、狀態也還不是「已釋出」
//   3) 有綁定LINE才會推播對應的提醒/通知訊息（但「釋出時段」這個動作本身不受此限）
//
// 這個函式要設定「每小時一次」的定時觸發條件才會自動執行，設定方式：
// 在 Apps Script 選單「Floralway工具 > 設定每小時保留/催款/釋出檢查觸發條件（只需執行一次）」

function processReservationHolds() {
  var sh = getSheet_(SHEET_NAMES.BOOKINGS);
  var values = sh.getDataRange().getValues();
  var header = values[0];
  var idx = {};
  header.forEach(function (h, i) { idx[h] = i; });

  var servicesData = getActiveServices_().services;
  var serviceNameMap = {};
  servicesData.forEach(function (s) { serviceNameMap[s.code] = s.name; });

  var bankName = getScriptProp_('PAYMENT_BANK_NAME') || '';
  var bankAccount = getScriptProp_('PAYMENT_BANK_ACCOUNT') || '';
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var now = new Date();
  var currentHour = Number(Utilities.formatDate(now, tz, 'H'));
  var isNight = (currentHour >= 23 || currentHour < 9); // 晚上11點～隔天上午9點前，暫緩發送第2則提醒
  var TWELVE_HOURS_MS = 12 * 60 * 60 * 1000;
  var TWENTYFOUR_HOURS_MS = 24 * 60 * 60 * 1000;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var bookingId = row[idx['預約編號']];
    if (!bookingId) continue;

    var deposit = Number(row[idx['訂金金額']]) || 0;
    if (deposit <= 0) continue; // $0訂金不進入保留/催款/釋出流程

    if (row[idx['確認到帳時間']]) continue; // 已確認到帳，不用再處理
    if (row[idx['狀態']] === '已釋出') continue; // 已經釋出過，不重複處理

    var fillTime = row[idx['填單時間']];
    if (!(fillTime instanceof Date)) continue;
    var elapsedMs = now.getTime() - fillTime.getTime();

    var lineUserId = row[idx['LINE_userId']];
    var name = row[idx['姓名']];
    var serviceCodes = row[idx['服務項目']];
    var serviceDisplayNames = String(serviceCodes || '').split('、').map(function (code) {
      return serviceNameMap[code] || code;
    }).join('、');
    var eventId = row[idx['日曆事件ID']];

    if (elapsedMs >= TWENTYFOUR_HOURS_MS) {
      // ---- 滿24小時，未收到款項：釋出時段 ----
      deleteHoldEvent_(eventId);
      sh.getRange(r + 1, idx['狀態'] + 1).setValue('已釋出');
      sh.getRange(r + 1, idx['日曆事件ID'] + 1).setValue('');

      if (lineUserId) {
        var msg3 = name + ' 您好，由於預約保留期限已到，目前尚未收到您的款項，系統已先釋出該諮詢時段。\n\n' +
          '若後續仍想安排諮詢，歡迎隨時點擊下方選單重新挑選合適的時段，謝謝您！';
        var ok3 = sendLinePush_(lineUserId, msg3);
        logNotification_(bookingId, name, lineUserId, '保留逾時釋出通知', ok3, ok3 ? '' : '發送失敗');
      } else {
        logNotification_(bookingId, name, '', '保留逾時釋出（未綁定LINE，未發送通知）', false, '客戶未綁定LINE');
      }
      continue;
    }

    if (elapsedMs >= TWELVE_HOURS_MS && !row[idx['提醒_保留中已發送']] && !isNight && lineUserId) {
      // ---- 滿12小時：發送「保留中」提醒（避開半夜） ----
      var deadline = new Date(fillTime.getTime() + TWENTYFOUR_HOURS_MS);
      var deadlineStr = Utilities.formatDate(deadline, tz, 'yyyy/MM/dd HH:mm');
      var dateStr = formatDateValue_(row[idx['預約日期']]);
      var time = row[idx['預約時段']];

      var msg2 = '哈囉 ' + name + '，貼心提醒您～\n\n' +
        '您的諮詢時段目前保留中，保留將於 ' + deadlineStr + ' 截止喔！\n\n' +
        '✦ 預約項目：' + serviceDisplayNames + '\n' +
        '✦ 預約時間：' + dateStr + ' ' + time + '\n\n' +
        '【定金匯款資訊】\n' +
        '✦ 銀行：' + bankName + '\n' +
        '✦ 帳號：' + bankAccount + '\n' +
        '✦ 定金金額：' + deposit + ' 元\n\n' +
        '若手邊忙碌，記得抽空完成轉帳並回傳後五碼，期待與您見面！';

      var ok2 = sendLinePush_(lineUserId, msg2);
      if (ok2) {
        sh.getRange(r + 1, idx['提醒_保留中已發送'] + 1).setValue(now);
      }
      logNotification_(bookingId, name, lineUserId, '保留中提醒（滿12小時）', ok2, ok2 ? '' : '發送失敗');
    }
  }
}

/** 選單用：設定每小時自動檢查保留/催款/釋出的時間觸發條件（只需執行一次）
 *  注意：這個函式故意不加結尾底線，因為 Apps Script 編輯器的「執行函式」
 *  下拉選單會自動隱藏名稱結尾是底線的函式（視為私有輔助函式），
 *  這個函式需要讓 Evelyn 手動執行一次，所以特意保留可被選取。
 */
function setupReservationHoldTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'processReservationHolds' || fn === 'sendUnpaidReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processReservationHolds')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('已設定：每小時自動檢查一次保留中/需催款/需釋出的預約');
}

function formatDateValue_(dateValue) {
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy/MM/dd');
  }
  return String(dateValue);
}

// ===================== LINE Messaging API =====================

function sendLinePush_(userId, text) {
  var token = getScriptProp_('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) {
    Logger.log('尚未設定 LINE_CHANNEL_ACCESS_TOKEN，無法發送LINE訊息');
    return false;
  }
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    to: userId,
    messages: [{ type: 'text', text: text }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var resp = UrlFetchApp.fetch(url, options);
  var ok = resp.getResponseCode() === 200;
  if (!ok) Logger.log('LINE推播失敗: ' + resp.getContentText());
  return ok;
}

function logNotification_(bookingId, name, lineUserId, type, success, errorMsg) {
  var sh = getSheet_(SHEET_NAMES.NOTIFICATIONS);
  sh.appendRow([new Date(), bookingId, name, lineUserId, type, new Date(), success, errorMsg || '']);
}

// ===================== 自訂選單 =====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Floralway工具')
    .addItem('確認選取的預約（發送LINE通知）', 'confirmSelectedBooking')
    .addItem('設定每小時保留/催款/釋出檢查觸發條件（只需執行一次）', 'setupReservationHoldTrigger')
    .addToUi();
}
