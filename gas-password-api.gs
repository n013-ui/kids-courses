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
 * 1. 打開這個 GAS 專案
 * 2. 左側點「⚙️ 專案設定」
 * 3. 拉到最下面「指令碼屬性」
 * 4. 找到 SITE_PW，點旁邊的鉛筆圖示
 * 5. 把值改成新密碼 → 儲存
 * 6. 完成！不需要重新部署
 * ════════════════════════════════════════════════
 */

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const cb     = (e.parameter && e.parameter.cb)     || '';

  let result;

  if (action === 'checkPw') {
    const props    = PropertiesService.getScriptProperties();
    let   storedPw = props.getProperty('SITE_PW');

    // 第一次執行時自動設定預設密碼
    if (!storedPw) {
      storedPw = 'sc1234';
      props.setProperty('SITE_PW', storedPw);
    }

    const inputPw = String(e.parameter.pw || '').trim();
    result = { ok: inputPw === storedPw };

  } else {
    result = { ok: false, msg: '課程密碼驗證 API 運作中' };
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
