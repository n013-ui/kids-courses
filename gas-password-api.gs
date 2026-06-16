/**
 * 課程網站帳號驗證 API
 * ════════════════════════════════════════════════
 *
 * 【設定步驟】
 * 1. 到 https://script.google.com 開啟已建立的專案
 * 2. 把這個檔案全部內容貼入編輯器取代舊內容
 * 3. 點「部署」→「管理部署」→ 編輯（鉛筆）→ 版本選「新版本」→ 部署
 *
 * 【新增 / 修改帳號密碼（不需要改程式碼）】
 * 1. 打開下方 SHEET_ID 對應的 Google 試算表
 * 2. 切換到「帳號」工作表（第一次執行會自動建立）
 * 3. A 欄填帳號、B 欄填密碼，一列一個孩子
 * 4. 儲存後立即生效，不需重新部署
 * ════════════════════════════════════════════════
 */

const SHEET_ID = '1cs3oY6ncFJgwaE8W5v4lTz6CEhveEIurw_pGEEoEIiQ';

function getAccounts() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('帳號');

  if (!sheet) {
    sheet = ss.insertSheet('帳號');
    sheet.appendRow(['帳號', '密碼']);
    sheet.appendRow(['孩子一', 'child1']);
    sheet.appendRow(['孩子二', 'child2']);
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 150);
  }

  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => String(r[0]).trim() !== '');
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const cb     = (e.parameter && e.parameter.cb)     || '';
  let result;

  try {
    if (action === 'checkLogin') {
      const inputUser = String(e.parameter.user || '').trim();
      const inputPw   = String(e.parameter.pw   || '').trim();
      const accounts  = getAccounts();
      const match     = accounts.find(r =>
        String(r[0]).trim() === inputUser &&
        String(r[1]).trim() === inputPw
      );
      result = match ? { ok: true, user: inputUser } : { ok: false };
    } else {
      result = { ok: false, msg: '課程平台 API 運作中' };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }

  const json = JSON.stringify(result);
  if (cb) {
    return ContentService.createTextOutput(`${cb}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
