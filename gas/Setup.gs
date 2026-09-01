/**
 * 一次性初始化工具：建立五個分頁與標題列。
 * 在 Apps Script 編輯器選擇 initSheets 這個函式，按「執行」一次即可。
 * 之後要匯入國中部資料，再執行 SeedJunior.gs 裡的 seedJunior()。
 */
function initSheets() {
  const ss = ss_();
  const studentHeaders = ['學號', '班級', '座號', '姓名', 'vh', 'level', '身分證後四碼'];
  createSheetWithHeaders_(ss, SHEETS.STUDENTS, studentHeaders);
  // 身分證後四碼可能 0 開頭，欄位一開始就設成純文字，避免手動輸入時前導 0 被 Sheets 自動吃掉
  forceIdLast4ColumnAsText_(sheet_(SHEETS.STUDENTS), studentHeaders.indexOf('身分證後四碼') + 1);
  createSheetWithHeaders_(ss, SHEETS.CLUBS, ['id', 'level', 'name', 'teacher', 'cap', 'fee', 'location', 'intro']);
  createSheetWithHeaders_(ss, SHEETS.CONFIG, ['level', 'label', 'excludedNote', 'infoLink_href', 'infoLink_text', 'SELECT_OPEN', 'SELECT_CLOSE', 'MAKEUP_OPEN', 'MAKEUP_CLOSE']);
  createSheetWithHeaders_(ss, SHEETS.NOTICES, ['level', 'order', 'text']);

  const respHeaders = ['timestamp', 'sid', 'level', 'cls', 'seat', 'name'];
  for (let i = 1; i <= REQUIRED_CHOICES; i++) respHeaders.push('choice' + i);
  respHeaders.push('note', 'updatedAt');
  createSheetWithHeaders_(ss, SHEETS.RESPONSES, respHeaders);

  Logger.log('五個分頁已就緒：Students / Clubs / Config / Notices / Responses');
}

function createSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

/** 一次性遷移工具：既有的 Students 分頁是在加入「身分證後四碼」欄位之前建立的，
 * initSheets() 只會幫「全新」的分頁補標題列，不會動已經有資料的舊分頁，
 * 所以要用這支手動補上這欄。在 Apps Script 編輯器選這個函式執行一次即可，
 * 已經有這欄的話會直接跳過，重複執行也不會出問題。
 * 補上欄位後，記得到 Students 分頁把每位學生的身分證後四碼實際填進去，
 * 在填之前 lookup 會因為欄位是空的而一律判定「身分證後四碼不正確」。
 *
 * ⚠️ 身分證後四碼可能是 0 開頭（例如 0023）。這裡會把整欄格式設成「純文字」，
 * 這樣不管是這支函式自己寫入、或之後有人直接在 Sheets 手動輸入，開頭的 0
 * 都不會被自動吃掉變成數字（Code.gs 的 lookup 比對也另外做了補零防呆，
 * 但欄位本身設成文字格式才是根本解法，兩邊一起處理比較保險）。 */
function addIdLast4ColumnToStudents() {
  const sh = sheet_(SHEETS.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf('身分證後四碼');
  if (col !== -1) {
    forceIdLast4ColumnAsText_(sh, col + 1);
    Logger.log('Students 分頁已經有「身分證後四碼」欄位，不用重複新增，但已確保欄位格式是純文字。');
    return;
  }
  const newCol = headers.length + 1;
  sh.getRange(1, newCol).setValue('身分證後四碼');
  forceIdLast4ColumnAsText_(sh, newCol);
  Logger.log('已在 Students 分頁新增「身分證後四碼」欄位（格式為純文字），請記得手動填入每位學生的資料。');
}

function forceIdLast4ColumnAsText_(sh, colIndex) {
  sh.getRange(1, colIndex, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@');
}

/** 一次性遷移工具：既有的 Config 分頁是在加入「補選」時間窗之前建立的，
 * 補上 MAKEUP_OPEN / MAKEUP_CLOSE 兩欄（已經有的話會跳過，可重複執行）。
 * 補完欄位後，記得呼叫 writeMakeupWindow_() 把各 level 實際的補選開放/截止時間填進去，
 * 否則 submitResponse_ 會判斷成「未設定」而不卡控補選窗口。 */
function addMakeupWindowColumnsToConfig() {
  const sh = sheet_(SHEETS.CONFIG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const missing = ['MAKEUP_OPEN', 'MAKEUP_CLOSE'].filter(h => headers.indexOf(h) === -1);
  if (!missing.length) {
    Logger.log('Config 分頁已經有補選時間窗欄位，不用重複新增。');
    return;
  }
  missing.forEach((h, i) => sh.getRange(1, headers.length + 1 + i).setValue(h));
  Logger.log('已在 Config 分頁新增欄位：%s，請記得呼叫 writeMakeupWindow_() 填入各 level 的補選時間。', missing.join(', '));
}

/** 寫入（或覆蓋）某個 level 的補選開放/截止時間，只更新 MAKEUP_OPEN/MAKEUP_CLOSE 這兩欄，
 * 不影響同一列的其他欄位（例如 SELECT_OPEN/SELECT_CLOSE）。
 * openIso / closeIso 格式範例："2026-09-07T00:00:00"。
 * 範例：writeMakeupWindow_('junior', '2026-09-07T00:00:00', '2026-09-07T20:00:00'); */
function writeMakeupWindow_(level, openIso, closeIso) {
  const sh = sheet_(SHEETS.CONFIG);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const levelCol = headers.indexOf('level');
  const openCol = headers.indexOf('MAKEUP_OPEN');
  const closeCol = headers.indexOf('MAKEUP_CLOSE');
  if (openCol === -1 || closeCol === -1) {
    throw new Error('Config 分頁還沒有 MAKEUP_OPEN/MAKEUP_CLOSE 欄位，請先執行 addMakeupWindowColumnsToConfig()。');
  }
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][levelCol]) === level) {
      sh.getRange(i + 1, openCol + 1).setValue(openIso);
      sh.getRange(i + 1, closeCol + 1).setValue(closeIso);
      Logger.log('已更新 %s 的補選時間：%s ～ %s', level, openIso, closeIso);
      return;
    }
  }
  throw new Error('Config 分頁找不到 level=' + level + ' 的那一列，請先確認 Config 分頁已有這個 level 的設定。');
}

/* ---------- 給 Seed*.gs 呼叫的寫入輔助函式 ---------- */

/** 寫入學生名冊，以「學號」為 key 檢查重複：已經在 Students 分頁裡的學號、
 * 或本次傳入的陣列裡自己重複的學號，都只保留第一筆、其餘略過不寫入，
 * 這樣同一個 seed 函式重跑第二次，也不會把同一位學生寫成兩列。 */
function writeStudents_(students) {
  const sh = sheet_(SHEETS.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const sidCol = headers.indexOf('學號');

  const existing = new Set(
    sh.getLastRow() > 1
      ? sh.getRange(2, sidCol + 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim())
      : []
  );

  const seen = new Set();
  const skipped = [];
  const toWrite = students.filter(s => {
    const sid = String(s['學號']).trim();
    if (existing.has(sid) || seen.has(sid)) {
      skipped.push(sid);
      return false;
    }
    seen.add(sid);
    return true;
  });

  if (skipped.length) {
    Logger.log('writeStudents_：略過 %s 筆已存在的學號（不重複寫入）：%s', skipped.length, skipped.join(', '));
  }

  const rows = toWrite.map(s => headers.map(h => s[h] !== undefined ? s[h] : ''));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  Logger.log('writeStudents_：實際寫入 %s 筆', rows.length);
}

/** 依學號更新既有學生列的座號（用在名冊已經匯入、之後才補到座號資料的情況）。
 * seatMap 格式：{ 學號: 座號 }。找不到對應學號的列會略過並記錄在 Logger，不會新增列。 */
function updateStudentSeats_(seatMap) {
  const sh = sheet_(SHEETS.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const sidCol = headers.indexOf('學號');
  const seatCol = headers.indexOf('座號');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const sids = sh.getRange(2, sidCol + 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());

  let updated = 0;
  const notFound = [];
  Object.keys(seatMap).forEach(sid => {
    const rowIdx = sids.indexOf(String(sid).trim());
    if (rowIdx === -1) {
      notFound.push(sid);
      return;
    }
    sh.getRange(rowIdx + 2, seatCol + 1).setValue(seatMap[sid]);
    updated++;
  });

  Logger.log('updateStudentSeats_：更新 %s 筆座號，找不到 %s 筆：%s', updated, notFound.length, notFound.join(', '));
}

/** 依學號更新既有學生列的「身分證後四碼」（用在名冊已經匯入、之後才補這欄身分驗證資料的情況）。
 * idLast4Map 格式：{ 學號: '後四碼' }。找不到對應學號的列會略過並記錄在 Logger，不會新增列。
 * 執行前要先確認 Students 分頁已經有「身分證後四碼」欄位（沒有的話先執行 addIdLast4ColumnToStudents()）。 */
function updateStudentIdLast4_(idLast4Map) {
  const sh = sheet_(SHEETS.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const sidCol = headers.indexOf('學號');
  const idCol = headers.indexOf('身分證後四碼');
  if (idCol === -1) {
    throw new Error('Students 分頁還沒有「身分證後四碼」欄位，請先執行 addIdLast4ColumnToStudents()。');
  }
  forceIdLast4ColumnAsText_(sh, idCol + 1); // 保險起見每次都重新確保欄位是純文字格式，避免前導 0 被吃掉
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const sids = sh.getRange(2, sidCol + 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());

  let updated = 0;
  const notFound = [];
  Object.keys(idLast4Map).forEach(sid => {
    const rowIdx = sids.indexOf(String(sid).trim());
    if (rowIdx === -1) {
      notFound.push(sid);
      return;
    }
    // 身分證後四碼可能有前導 0（例如 "0016"）：用 setValue() 寫入字串時 Sheets 不會自動轉成數字，
    // 前導 0 不會被吃掉，所以這裡直接寫字串就好，不要自己加單引號（那樣反而會多一個字元進儲存格）。
    sh.getRange(rowIdx + 2, idCol + 1).setValue(String(idLast4Map[sid]));
    updated++;
  });

  Logger.log('updateStudentIdLast4_：更新 %s 筆身分證後四碼，找不到 %s 筆：%s', updated, notFound.length, notFound.join(', '));
}

function writeClubs_(clubs) {
  const sh = sheet_(SHEETS.CLUBS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = clubs.map(c => headers.map(h => c[h] !== undefined ? c[h] : ''));
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/** 寫入（或覆蓋）某個 level 的一列設定 */
function writeConfig_(cfg) {
  const sh = sheet_(SHEETS.CONFIG);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const values = sh.getDataRange().getValues();
  const levelCol = headers.indexOf('level');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][levelCol]) === cfg.level) { targetRow = i + 1; break; }
  }
  const rowArr = headers.map(h => cfg[h] !== undefined ? cfg[h] : '');
  if (targetRow === -1) {
    sh.appendRow(rowArr);
  } else {
    sh.getRange(targetRow, 1, 1, headers.length).setValues([rowArr]);
  }
}

/** 覆蓋某個 level 的所有公告（會先刪掉該 level 的舊公告再寫入新的） */
function writeNotices_(level, notices) {
  const sh = sheet_(SHEETS.NOTICES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const levelCol = headers.indexOf('level');

  // 由下往上刪除該 level 的舊資料列，避免刪除時列號位移出錯
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][levelCol]) === level) sh.deleteRow(i + 1);
  }
  const rows = notices.map((text, i) => [level, i + 1, text]);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}
