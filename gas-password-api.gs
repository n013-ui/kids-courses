/**
 * 課程網站 API（帳號驗證 + 雲端進度同步）
 * ════════════════════════════════════════════════
 *
 * 【更新步驟】
 * 1. 打開 GAS 專案，把這個檔案內容全部貼入取代舊的
 * 2. 點「部署」→「管理部署」→ 編輯（鉛筆）→ 版本選「新版本」→ 部署
 * 3. 若試算表已有舊的「進度」工作表，請先刪除再重新部署
 *
 * 【帳號管理】
 * 試算表 →「帳號」工作表 → A欄帳號、B欄密碼
 *
 * 【進度查看】
 * 試算表 →「進度」工作表 → 每列 = 一個孩子的一門課
 * ════════════════════════════════════════════════
 */

const SHEET_ID = '1cs3oY6ncFJgwaE8W5v4lTz6CEhveEIurw_pGEEoEIiQ';

// ── 帳號工作表 ──────────────────────────────
function getAccounts() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('帳號');
  if (!sheet) {
    sheet = ss.insertSheet('帳號');
    sheet.appendRow(['帳號', '密碼', '預設頁面']);
    sheet.appendRow(['孩子一', 'child1', 'G8']);
    sheet.appendRow(['孩子二', 'child2', 'G8']);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 120);
  }
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => String(r[0]).trim() !== '');
}

// ── 進度工作表（一列 = 一個孩子的一門課）──
function getProgressSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('進度');
  if (!sheet) {
    sheet = ss.insertSheet('進度');
    sheet.appendRow(['帳號', '課程鍵', '已看影片', '最後更新']);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 320);
    sheet.setColumnWidth(4, 160);
  }
  return sheet;
}

// 載入進度：讀出該帳號所有課程的已看影片，合併成一個逗號清單回傳
function loadProgress(user) {
  const sheet = getProgressSheet();
  if (sheet.getLastRow() < 2) return { ok: true, watched: '' };

  const data  = sheet.getDataRange().getValues();
  const parts = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === user && data[i][2]) {
      parts.push(String(data[i][2]));
    }
  }
  return { ok: true, watched: parts.join(',') };
}

// 儲存進度：接收 JSON {"G8|理化|上":"B3C1-1,B3C1-2",...}，每門課寫一列
function saveProgress(user, courseDataJson) {
  const sheet = getProgressSheet();
  const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  let courseData;
  try { courseData = JSON.parse(courseDataJson); }
  catch(e) { return { ok: false, error: 'JSON 格式錯誤' }; }

  const rows = sheet.getLastRow() >= 2
    ? sheet.getDataRange().getValues()
    : [['帳號','課程鍵','已看影片','最後更新']];

  for (const [courseKey, watched] of Object.entries(courseData)) {
    let rowNum = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === user &&
          String(rows[i][1]).trim() === courseKey) {
        rowNum = i + 1;
        rows[i][2] = watched;
        break;
      }
    }
    if (rowNum > 0) {
      sheet.getRange(rowNum, 3, 1, 2).setValues([[watched, now]]);
    } else {
      sheet.appendRow([user, courseKey, watched, now]);
      rows.push([user, courseKey, watched, now]);
    }
  }
  return { ok: true };
}

// ── HTTP 進入點 ─────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const cb     = (e.parameter && e.parameter.cb)     || '';
  let result;

  try {
    if (action === 'checkLogin') {
      const inputUser = String(e.parameter.user || '').trim();
      const inputPw   = String(e.parameter.pw   || '').trim();
      const match     = getAccounts().find(r =>
        String(r[0]).trim() === inputUser && String(r[1]).trim() === inputPw
      );
      result = match
        ? { ok: true, user: inputUser, defaultGrade: String(match[2] || 'G8').trim() || 'G8' }
        : { ok: false };

    } else if (action === 'loadProgress') {
      result = loadProgress(String(e.parameter.user || '').trim());

    } else if (action === 'saveProgress') {
      result = saveProgress(
        String(e.parameter.user || '').trim(),
        String(e.parameter.data || '{}')
      );

    } else {
      result = { ok: true, msg: '課程平台 API 運作中' };
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
