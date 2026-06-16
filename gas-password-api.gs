/**
 * 課程網站密碼驗證 API
 * ════════════════════════════════════════════════
 *
 * 【設定步驟】
 * 1. 到 https://script.google.com 建立新專案
 * 2. 把這個檔案全部內容貼入編輯器（取代預設的 myFunction）
 * 3. 點「部署」→「新部署」→ 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：所有人
 * 4. 複製部署網址，貼到 config.js 的 GAS_URL 那行
 *
 * 【修改密碼方法（不需要改任何程式碼）】
 * 1. 打開下方 SHEET_ID 對應的 Google 試算表
 * 2. 切換到「設定」工作表（第一次執行 API 會自動建立）
 * 3. 找到「網站密碼」那列，修改右邊 B 欄的值
 * 4. 完成！立即生效，不需要重新部署
 * ════════════════════════════════════════════════
 */

const SHEET_ID = '1cs3oY6ncFJgwaE8W5v4lTz6CEhveEIurw_pGEEoEIiQ';

// 從試算表讀取密碼（第一次執行時自動建立「設定」工作表）
function getPassword() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('設定');

  if (!sheet) {
    sheet = ss.insertSheet('設定');
    sheet.appendRow(['設定項目', '設定值']);
    sheet.appendRow(['網站密碼', 'sc1234']);
    sheet.getRange('A1:B1').setFontWeight('bold');
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === '網站密碼') {
      return String(data[i][1]).trim();
    }
  }
  return 'sc1234'; // 找不到時的預設值
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const cb     = (e.parameter && e.parameter.cb)     || '';
  let result;

  try {
    if (action === 'checkPw') {
      const storedPw = getPassword();
      const inputPw  = String(e.parameter.pw || '').trim();
      result = { ok: inputPw === storedPw };
    } else {
      result = { ok: false, msg: '課程密碼驗證 API 運作中' };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  const json = JSON.stringify(result);

  // 支援 JSONP（跨網域呼叫用）
  if (cb) {
    return ContentService.createTextOutput(`${cb}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
