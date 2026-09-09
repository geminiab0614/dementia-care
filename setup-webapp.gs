/**
 * 一般精神科 飲料點單 — 收單後端（Apps Script 網頁應用程式）
 *
 * 為什麼要這一支：
 *   Google 表單本身不會擋重複，純前端也查不到別人填過什麼（查得到就等於同仁
 *   互相看得到訂單）。這支程式跑在你的帳號下，收到訂單時先查姓名，
 *   重複就直接回「已經點過了」而不寫入，同仁那邊看不到任何別人的資料。
 *
 * 部署步驟：
 *   1. script.google.com →「新增專案」，把原本的內容刪掉，貼上這整段
 *   2. 右上角「部署」→「新增部署作業」
 *   3. 齒輪圖示 → 選「網頁應用程式」
 *   4. 執行身分：「我」　　誰可以存取：「任何人」　←── 這兩個要設對
 *   5. 按「部署」，照指示授權（跟上次一樣，進階 → 前往…（不安全）→ 允許）
 *   6. 複製「網頁應用程式」的網址（.../exec 結尾），貼到點單頁的主辦人設定
 *
 * 改動之後要重新部署才會生效：部署 →「管理部署作業」→ 鉛筆 → 版本選「新版本」→ 部署
 */

/** 訂單要寫進哪一份試算表（表單的回應試算表） */
var SHEET_ID = '1BvvEFIwf1mRgEShsqGzXJQE6wSOs05fdy2l7PPPZgT8';

/**
 * 比對方式。
 *   false = 只看姓名：全院同名的人只能有一個點成功
 *   true  = 看病房＋姓名：不同病房的同名同仁可以各點一杯
 * 兩個病房剛好有同名同仁時，false 會誤擋後面那位，改成 true 即可。
 */
var MATCH_WARD_TOO = false;

/** 欄位在試算表的位置（表單建立的預設順序：時間戳記, 病房, 姓名, 飲料, 甜度, 冰塊） */
var COL_WARD = 2;
var COL_NAME = 3;

function doGet(e) {
  var p = (e && e.parameter) || {};
  var result;
  try {
    result = (p.action === 'submit') ? submitOrder(p) : {ok: false, reason: 'bad_action'};
  } catch (err) {
    result = {ok: false, reason: 'error', detail: String(err)};
  }
  var json = JSON.stringify(result);

  // 靜態網頁沒辦法直接讀跨網域的回應，用 JSONP 把結果包成一段可執行的 JS 送回去
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function submitOrder(p) {
  var ward  = trim(p.ward);
  var name  = trim(p.name);
  var drink = trim(p.drink);
  var sugar = trim(p.sugar);
  var ice   = trim(p.ice);

  if (!ward || !name || !drink || !sugar || !ice) return {ok: false, reason: 'incomplete'};

  // 同時有兩個人在送出時，鎖住避免兩筆同名同時通過檢查
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return {ok: false, reason: 'busy'};
  }

  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var last = sheet.getLastRow();

    if (last >= 2) {
      var rows = sheet.getRange(2, 1, last - 1, Math.max(COL_WARD, COL_NAME)).getValues();
      var key = personKey(ward, name);
      for (var i = 0; i < rows.length; i++) {
        var existing = personKey(rows[i][COL_WARD - 1], rows[i][COL_NAME - 1]);
        if (existing && existing === key) return {ok: false, reason: 'duplicate'};
      }
    }

    sheet.appendRow([new Date(), ward, name, drink, sugar, ice]);
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

/** 產生比對用的鍵值：去掉所有空白、統一大小寫，避免「王小美」和「王 小美」被當成兩個人 */
function personKey(ward, name) {
  var n = norm(name);
  if (!n) return '';
  return MATCH_WARD_TOO ? (norm(ward) + '|' + n) : n;
}

function norm(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
}

function trim(s) {
  return String(s == null ? '' : s).trim();
}
