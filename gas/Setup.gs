/**
 * 一次性初始化工具：建立五個分頁與標題列。
 * 在 Apps Script 編輯器選擇 initSheets 這個函式，按「執行」一次即可。
 * 之後要匯入國中部資料，再執行 SeedJunior.gs 裡的 seedJunior()。
 */
function initSheets() {
  const ss = ss_();
  createSheetWithHeaders_(ss, SHEETS.STUDENTS, ['學號', '班級', '座號', '姓名', 'vh', 'level']);
  createSheetWithHeaders_(ss, SHEETS.CLUBS, ['id', 'level', 'name', 'teacher', 'cap', 'fee', 'location', 'intro']);
  createSheetWithHeaders_(ss, SHEETS.CONFIG, ['level', 'label', 'excludedNote', 'infoLink_href', 'infoLink_text', 'SELECT_OPEN', 'SELECT_CLOSE']);
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
