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
    if (action === 'config') {
      return jsonOut_(configForLevel_(e.parameter.level));
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

/** config：讓前端可以即時讀到 Config 分頁目前設定的選填／補選開放時間，
 * 不用把時間寫死在 index.html 裡才能生效——改 Config 分頁就會立刻反映在畫面上
 * （「補選」按鈕能不能按之類）。只回傳時間窗相關欄位，不含 excludedNote/infoLink
 * 這些已經固定寫在前端 LEVEL_INFO 的內容。 */
function configForLevel_(level) {
  if (!level) return { ok: false, error: '缺少 level 參數' };
  const cfg = readObjects_(SHEETS.CONFIG).filter(r => String(r.level) === level)[0];
  if (!cfg) return { ok: true, found: false };
  return {
    ok: true,
    found: true,
    SELECT_OPEN: toIso_(cfg.SELECT_OPEN),
    SELECT_CLOSE: toIso_(cfg.SELECT_CLOSE),
    MAKEUP_OPEN: toIso_(cfg.MAKEUP_OPEN),
    MAKEUP_CLOSE: toIso_(cfg.MAKEUP_CLOSE)
  };
}

/** 身分證後四碼可能有前導 0（例如 0023）。Students 分頁的儲存格如果不是「純文字」格式，
 * 手動輸入 0023 會被 Sheets 自動當數字存成 23，讀回來就變成 "23" 而不是 "0023"；
 * 這裡統一補回 4 碼再比對，避免純粹因為儲存格格式問題就誤判驗證失敗。
 * 只在「整段都是數字、且不超過 4 碼」時才補零，避免誤把格式明顯錯誤的資料也硬湊成 4 碼。 */
function normalizeIdLast4_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  return /^\d{1,4}$/.test(s) ? s.padStart(4, '0') : s;
}

/* ---------- lookup：輸入學號（＋可選的身分證後四碼）時，查這一個學生的班級／座號／姓名，
 * 有給身分證後四碼的話會比對做身分驗證。
 * 回傳四種狀態：
 *   found=false                     查無此學號 → 前端走手動填寫備援
 *   found=true, verified=false      有給身分證後四碼但不符 → 前端顯示錯誤，不給資料、不給手動備援
 *   found=true, skippedVerification 沒有給身分證後四碼 → 略過驗證，直接回傳資料
 *                                   （查詢我的選填紀錄由前端擋掉，不會用這個狀態；
 *                                   只有填寫志願／補選允許在沒有身分證後四碼的情況下繼續）
 *   found=true, verified=true       兩者都對 → 回傳學生資料 */
function lookupStudent_(level, sid, idLast4) {
  if (!level) return { ok: false, error: '缺少 level 參數' };
  if (!sid) return { ok: false, error: '缺少學號' };
  const target = String(sid).trim();
  const student = getStudentsForLevel_(level).find(s => String(s['學號']) === target);
  if (!student) return { ok: true, found: false };

  const actual = normalizeIdLast4_(idLast4);
  if (!actual) {
    return { ok: true, found: true, skippedVerification: true, 學號: student['學號'], 班級: student['班級'], 座號: student['座號'], 姓名: student['姓名'] };
  }

  const expected = normalizeIdLast4_(student['身分證後四碼']);
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

/* ---------- query：查詢單一學號送出過的志願（只回傳這個學生自己的資料，
 * 現在會列出所有送出歷史，不是只有最新一筆） ---------- */
function queryResponse_(level, sid) {
  if (!sid) return { ok: false, error: '缺少學號' };
  const rows = readObjects_(SHEETS.RESPONSES).filter(r => String(r.sid).trim() === String(sid).trim());
  if (!rows.length) return { ok: true, found: false };

  // 依 updatedAt（沒有的話退回 timestamp）由舊到新排序，就算 Sheet 裡的列被手動搬動過順序也不會出錯。
  const timeOf = r => new Date(toIso_(r.updatedAt || r.timestamp)).getTime();
  const sorted = rows.slice().sort((a, b) => timeOf(a) - timeOf(b));

  const toRecord = r => {
    const choices = [];
    for (let i = 1; i <= REQUIRED_CHOICES; i++) choices.push(r['choice' + i] || '');
    return { name: r.name, cls: r.cls, seat: r.seat, choices: choices, note: r.note || '', updatedAt: toIso_(r.updatedAt || r.timestamp) };
  };

  const records = sorted.map(toRecord); // 由舊到新，最後一筆是最新
  const latest = records[records.length - 1];

  return {
    ok: true,
    found: true,
    // 保留舊版欄位（等於最新一筆），舊前端不用改也能繼續運作。
    name: latest.name,
    cls: latest.cls,
    seat: latest.seat,
    choices: latest.choices,
    note: latest.note,
    updatedAt: latest.updatedAt,
    // 新增：由舊到新的完整送出歷史，前端可以列出所有紀錄、把最後一筆標成目前有效。
    records: records,
    count: records.length
  };
}

/* ---------- submit：送出志願，每次都新增一列（不覆蓋舊資料，保留同一學號的完整送出歷史） ---------- */
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

  // 每次送出都新增一列，不覆蓋舊資料——Responses 分頁保留同一學號每一次送出的完整歷史，
  // 查詢時（queryResponse_）會把同一學號的所有列都列出來，最新一列視為目前有效的選填結果。
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEETS.RESPONSES);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    const now = new Date();
    const rowObj = {
      timestamp: now, sid: sid, level: level, cls: cls, seat: seat, name: name,
      note: note, updatedAt: now
    };
    choiceNames.forEach((n, i) => { rowObj['choice' + (i + 1)] = n; });

    const rowArr = headers.map(h => (h in rowObj) ? rowObj[h] : '');
    sh.appendRow(rowArr);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, updatedAt: new Date().toISOString() };
}
