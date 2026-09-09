/**
 * 一般精神科 飲料點單 — 一鍵建立 Google 表單
 *
 * 用法：
 *   1. 打開 script.google.com，按「新增專案」
 *   2. 把編輯器裡原本的內容全部刪掉，貼上這整段
 *   3. 按上方 ▶ 執行，照指示授權
 *   4. 下方「執行記錄」會印出一行 https://docs.google.com/forms/... 的網址，
 *      整行複製，貼到點單頁的「主辦人設定」裡按「解析」
 *
 * 這支程式會做四件事：
 *   - 建立一份有 病房／姓名／飲料／甜度／冰塊 五個簡答題的表單
 *   - 關掉收集電子郵件、關掉登入限制（同仁不用登入就能填）
 *   - 建立並連結回應試算表
 *   - 產生帶有 entry 代號的預先填入網址，讓點單頁知道要往哪送
 */
function setupDrinkForm() {
  var form = FormApp.create('一般精神科 飲料點單');
  form.setDescription('這是點單網頁的後台，請不要直接填這一份，改用點單網頁。');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setAcceptingResponses(true);

  // 公務帳號預設會限制「僅限機構內使用者」，關掉才能讓所有同仁填。
  // 個人 Gmail 帳號沒有這個設定，會丟例外，忽略即可。
  try { form.setRequireLogin(false); } catch (e) {}

  var titles  = ['病房', '姓名', '飲料', '甜度', '冰塊'];
  var markers = ['WARD', 'NAME', 'DRINK', 'SUGAR', 'ICE'];

  var items = titles.map(function (title) {
    return form.addTextItem().setTitle(title).setRequired(true);
  });

  // 用標記字串產生預先填入網址，網址裡就會帶著五個 entry 代號
  var resp = form.createResponse();
  items.forEach(function (item, i) {
    resp = resp.withItemResponse(item.createResponse(markers[i]));
  });
  var prefillUrl = resp.toPrefilledUrl();

  var sheetUrl = '（建立失敗，可到表單的「回覆」分頁自行建立）';
  try {
    var ss = SpreadsheetApp.create('飲料點單 回應');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    sheetUrl = ss.getUrl();
  } catch (e) {}

  var out = [
    '',
    '===== 複製下面這一整行，貼到點單頁的「主辦人設定」=====',
    prefillUrl,
    '',
    '===== 回應試算表（訂單會進到這裡）=====',
    sheetUrl,
    '',
    '===== 表單本身（通常用不到）=====',
    form.getEditUrl(),
    ''
  ].join('\n');

  Logger.log(out);
  return out;
}
