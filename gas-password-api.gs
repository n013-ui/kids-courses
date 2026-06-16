/**
 * 課程網站 API（帳號驗證 + 雲端進度同步）
 * ════════════════════════════════════════════════
 *
 * 【更新步驟】
 * 1. 打開 GAS 專案，把這個檔案內容全部貼入取代舊的
 * 2. 點「部署」→「管理部署」→ 編輯（鉛筆）→ 版本選「新版本」→ 部署
 *
 * 【帳號管理】
 * Google 試算表 →「帳號」工作表 → A欄帳號、B欄密碼
 *
 * 【進度查看】
 * Google 試算表 →「進度」工作表 → 可看到每位孩子已觀看哪些影片
 * ════════════════════════════════════════════════
 */

const SHEET_ID = '1cs3oY6ncFJgwaE8W5v4lTz6CEhveEIurw_pGEEoEIiQ';

// ── 帳號工作表 ──────────────────────────────
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

// ── 進度工作表 ──────────────────────────────
function getProgressSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('進度');
  if (!sheet) {
    sheet = ss.insertSheet('進度');
    sheet.appendRow(['帳號', '已觀看影片（逗號分隔）', '最後更新']);
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}

function loadProgress(user) {
  const sheet = getProgressSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === user) {
      return { ok: true, watched: String(data[i][1]) };
    }
  }
  return { ok: true, watched: '' };
}

function saveProgress(user, watched) {
  const sheet = getProgressSheet();
  const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === user) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[watched, now]]);
      return { ok: true };
    }
  }
  sheet.appendRow([user, watched, now]);
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
      result = match ? { ok: true, user: inputUser } : { ok: false };

    } else if (action === 'loadProgress') {
      result = loadProgress(String(e.parameter.user || '').trim());

    } else if (action === 'saveProgress') {
      result = saveProgress(
        String(e.parameter.user    || '').trim(),
        String(e.parameter.watched || '').trim()
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
