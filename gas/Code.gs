/**
 * 志願填寫系統後端（Google Apps Script Web App）
 *
 * 社團清單／公告／開放時間現在是「固定文字」，直接寫死在前端 index.html
 * 的 LEVEL_INFO 裡，這支後端不再提供、也不再讀 Clubs / Notices 分頁
 * （這兩個分頁留著沒關係，只是後端已經不會用到）。
 *
 * 資料來源：本試算表底下的分頁
 *   Students  學號 / 班級 / 座號 / 姓名 / vh / level / 身分證後四碼  —— 個資，只給 lookup／算 classes 用，不整包送出
 *             （身分證後四碼只用來在 lookup 時比對驗證，絕不會回傳給前端）
 *   Config    level / ... / SELECT_OPEN / SELECT_CLOSE  —— submit 送出時用來做「真正」的開放時間卡控
 *   Responses timestamp / sid / level / cls / seat / name / choice1..choice9 / note / updatedAt
 *
 * 部署方式：「部署」→「新增部署作業」→ 類型選「網頁應用程式」，
 *   執行身分：我，誰可以存取：任何人。部署後把 /exec 網址貼到前端 index.html 的 WEBAPP_URL。
 */

const REQUIRED_CHOICES = 9;

// 學生名冊的伺服器端快取秒數。改了 Students 分頁之後，最多要等這麼久
// 前端的學號查詢／班級清單才會看到新版本；想馬上生效可以在編輯器手動執行 clearCache()。
const CACHE_TTL_SECONDS = 300;

const SHEETS = {
  STUDENTS: 'Students',
  CLUBS: 'Clubs',
  CONFIG: 'Config',
  NOTICES: 'Notices',
  RESPONSES: 'Responses'
};

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'classes') {
      return jsonOut_(classesForLevel_(e.parameter.level));
    }
    if (action === 'lookup') {
      return jsonOut_(lookupStudent_(e.parameter.level, e.parameter.sid, e.parameter.idLast4));
    }
    if (action === 'query') {
      return jsonOut_(queryResponse_(e.parameter.level, e.parameter.sid));
    }
    return jsonOut_({ ok: false, error: '未知的 action：' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'submit') {
      return jsonOut_(submitResponse_(payload));
    }
    return jsonOut_({ ok: false, error: '未知的 action：' + payload.action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁：' + name + '，請先執行 initSheets()。');
  return sh;
}

/** 把整張表讀成 [{欄名: 值, ...}, ...]，第一列視為標題列，跳過整列空白的資料列 */
function readObjects_(sheetName) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function toIso_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(v || '');
}

/**
 * 某個就讀階段的學生名冊（伺服器端快取，不含 vh —— 目前沒有任何地方用到這欄）。
 * 只給後端自己查（lookupStudent_ / 算 classes 用），不會整包送到瀏覽器。
 */
function getStudentsForLevel_(level) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'students_' + level;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const students = readObjects_(SHEETS.STUDENTS)
    .filter(r => String(r.level) === level)
    .map(r => ({ 學號: r['學號'], 班級: r['班級'], 座號: r['座號'], 姓名: r['姓名'], 身分證後四碼: r['身分證後四碼'] }));

  cache.put(cacheKey, JSON.stringify(students), CACHE_TTL_SECONDS);
  return students;
}

/** classes：查無學號、要手動填寫時，班級下拉選單用的資料（該階段所有學生的班級去重排序）。 */
function classesForLevel_(level) {
  if (!level) return { ok: false, error: '缺少 level 參數' };
  const classes = [...new Set(getStudentsForLevel_(level).map(s => s['班級']))].sort();
  return { ok: true, level: level, classes: classes };
}

/* ---------- lookup：輸入學號 + 身分證後四碼時，查這一個學生的班級／座號／姓名，
 * 並比對身分證後四碼做身分驗證。
 * 回傳三種狀態：
 *   found=false                查無此學號 → 前端走手動填寫備援
 *   found=true, verified=false 學號存在但身分證後四碼不符 → 前端顯示錯誤，不給資料、不給手動備援
 *   found=true, verified=true  兩者都對 → 回傳學生資料 */
function lookupStudent_(level, sid, idLast4) {
  if (!level) return { ok: false, error: '缺少 level 參數' };
  if (!sid) return { ok: false, error: '缺少學號' };
  const target = String(sid).trim();
  const student = getStudentsForLevel_(level).find(s => String(s['學號']) === target);
  if (!student) return { ok: true, found: false };

  const expected = String(student['身分證後四碼'] || '').trim();
  const actual = String(idLast4 || '').trim();
  if (!expected || expected !== actual) {
    return { ok: true, found: true, verified: false };
  }
  return { ok: true, found: true, verified: true, 學號: student['學號'], 班級: student['班級'], 座號: student['座號'], 姓名: student['姓名'] };
}

/** 手動清除快取：改了 Students 分頁後想馬上生效，在編輯器選這個函式執行即可。 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  ['junior', 'senior'].forEach(level => {
    cache.remove('students_' + level);
  });
  Logger.log('快取已清除');
}

/* ---------- query：查詢單一學號目前送出的志願（只回傳這個學生自己的資料） ---------- */
function queryResponse_(level, sid) {
  if (!sid) return { ok: false, error: '缺少學號' };
  const rows = readObjects_(SHEETS.RESPONSES).filter(r => String(r.sid).trim() === String(sid).trim());
  if (!rows.length) return { ok: true, found: false };

  const last = rows[rows.length - 1];
  const choices = [];
  for (let i = 1; i <= REQUIRED_CHOICES; i++) choices.push(last['choice' + i] || '');
  return {
    ok: true,
    found: true,
    name: last.name,
    cls: last.cls,
    seat: last.seat,
    choices: choices,
    note: last.note || '',
    updatedAt: toIso_(last.updatedAt)
  };
}

/* ---------- submit：送出／覆蓋志願（以學號為 key，重複送出視為更新） ---------- */
function submitResponse_(payload) {
  const level = payload.level;
  const sid = String(payload.sid || '').trim();
  const cls = payload.cls || '';
  const seat = payload.seat || '';
  const name = payload.name || '';
  const note = payload.note || '';
  // 社團清單現在是前端固定資料，這裡直接收社團「名稱」字串（不再是 id），
  // 不用另外讀 Clubs 分頁比對——社團名稱本身有沒有效，交給前端的固定清單把關。
  const choiceNames = (payload.choices || []).map(v => String(v || '').trim());

  if (!level) return { ok: false, error: '缺少 level' };
  if (!sid) return { ok: false, error: '缺少學號' };
  if (!name) return { ok: false, error: '缺少姓名' };
  if (choiceNames.length !== REQUIRED_CHOICES || choiceNames.some(n => !n)) {
    return { ok: false, error: '志願數量須為 ' + REQUIRED_CHOICES + ' 個' };
  }
  if (new Set(choiceNames).size !== choiceNames.length) {
    return { ok: false, error: '志願不可重複' };
  }

  // 伺服器端再次確認選填時間窗，避免有人繞過前端的時間卡控。
  // 正式選填（SELECT_OPEN/CLOSE）跟補選（MAKEUP_OPEN/CLOSE）是兩組獨立的時間窗，
  // 只要現在落在「其中一個」窗口內就放行；兩組都沒設定或都不在窗內才擋下來。
  const cfg = readObjects_(SHEETS.CONFIG).filter(r => String(r.level) === level)[0];
  if (cfg) {
    const now = new Date();
    const inWindow = (openVal, closeVal) => {
      if (!openVal || !closeVal) return null; // 該組窗口未設定，不列入判斷
      const open = new Date(toIso_(openVal));
      const close = new Date(toIso_(closeVal));
      return now >= open && now < close;
    };
    const inSelectWindow = inWindow(cfg.SELECT_OPEN, cfg.SELECT_CLOSE);
    const inMakeupWindow = inWindow(cfg.MAKEUP_OPEN, cfg.MAKEUP_CLOSE);
    const anyWindowConfigured = inSelectWindow !== null || inMakeupWindow !== null;
    if (anyWindowConfigured && !inSelectWindow && !inMakeupWindow) {
      return { ok: false, error: '目前不在選填開放時間內' };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEETS.RESPONSES);
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const sidCol = headers.indexOf('sid');
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][sidCol]).trim() === sid) { targetRow = i + 1; break; }
    }

    const now = new Date();
    const rowObj = {
      timestamp: (targetRow === -1) ? now : values[targetRow - 1][headers.indexOf('timestamp')],
      sid: sid, level: level, cls: cls, seat: seat, name: name,
      note: note, updatedAt: now
    };
    choiceNames.forEach((n, i) => { rowObj['choice' + (i + 1)] = n; });

    const rowArr = headers.map(h => (h in rowObj) ? rowObj[h] : '');

    if (targetRow === -1) {
      sh.appendRow(rowArr);
    } else {
      sh.getRange(targetRow, 1, 1, headers.length).setValues([rowArr]);
    }
  } finally {
    lock.releaseLock();
  }

  return { ok: true, updatedAt: new Date().toISOString() };
}
