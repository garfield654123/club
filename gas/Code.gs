/**
 * 志願填寫系統後端（Google Apps Script Web App）
 *
 * 資料來源：本試算表底下的分頁
 *   Students  學號 / 班級 / 座號 / 姓名 / vh / level
 *   Clubs     id / level / name / teacher / cap / fee / location / intro
 *   Config    level / label / excludedNote / infoLink_href / infoLink_text / SELECT_OPEN / SELECT_CLOSE
 *   Notices   level / order / text
 *   Responses timestamp / sid / level / cls / seat / name / choice1..choice9 / note / updatedAt
 *
 * 部署方式：「部署」→「新增部署作業」→ 類型選「網頁應用程式」，
 *   執行身分：我，誰可以存取：任何人。部署後把 /exec 網址貼到前端 index.html 的 WEBAPP_URL。
 */

const REQUIRED_CHOICES = 9;

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
    if (action === 'bootstrap') {
      return jsonOut_(bootstrap_(e.parameter.level));
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

/* ---------- bootstrap：某個就讀階段的名冊／社團／公告／設定 ---------- */
function bootstrap_(level) {
  if (!level) return { ok: false, error: '缺少 level 參數' };

  const students = readObjects_(SHEETS.STUDENTS)
    .filter(r => String(r.level) === level)
    .map(r => ({ 學號: r['學號'], 班級: r['班級'], 座號: r['座號'], 姓名: r['姓名'], vh: r['vh'] }));

  const clubs = readObjects_(SHEETS.CLUBS)
    .filter(r => String(r.level) === level)
    .map(r => ({ id: r.id, name: r.name, teacher: r.teacher, cap: r.cap, fee: r.fee, location: r.location, intro: r.intro }));

  const cfg = readObjects_(SHEETS.CONFIG).filter(r => String(r.level) === level)[0] || {};

  const notices = readObjects_(SHEETS.NOTICES)
    .filter(r => String(r.level) === level)
    .sort((a, b) => Number(a.order) - Number(b.order))
    .map(r => r.text);

  return {
    ok: true,
    level: level,
    label: cfg.label || level,
    excludedNote: cfg.excludedNote || '',
    infoLink: cfg.infoLink_href ? { href: cfg.infoLink_href, text: cfg.infoLink_text || '' } : null,
    notices: notices,
    students: students,
    clubs: clubs,
    config: {
      SELECT_OPEN: toIso_(cfg.SELECT_OPEN),
      SELECT_CLOSE: toIso_(cfg.SELECT_CLOSE)
    }
  };
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
  const choiceIds = payload.choices || [];

  if (!level) return { ok: false, error: '缺少 level' };
  if (!sid) return { ok: false, error: '缺少學號' };
  if (!name) return { ok: false, error: '缺少姓名' };
  if (choiceIds.length !== REQUIRED_CHOICES) {
    return { ok: false, error: '志願數量須為 ' + REQUIRED_CHOICES + ' 個' };
  }
  if (new Set(choiceIds.map(String)).size !== choiceIds.length) {
    return { ok: false, error: '志願不可重複' };
  }

  const clubs = readObjects_(SHEETS.CLUBS).filter(r => String(r.level) === level);
  const clubMap = {};
  clubs.forEach(c => { clubMap[String(c.id)] = c; });

  let choiceNames;
  try {
    choiceNames = choiceIds.map(id => {
      const c = clubMap[String(id)];
      if (!c) throw new Error('志願包含無效的社團 id：' + id);
      return c.name;
    });
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }

  // 伺服器端再次確認選填時間窗，避免有人繞過前端的時間卡控
  const cfg = readObjects_(SHEETS.CONFIG).filter(r => String(r.level) === level)[0];
  if (cfg && cfg.SELECT_OPEN && cfg.SELECT_CLOSE) {
    const now = new Date();
    const open = new Date(toIso_(cfg.SELECT_OPEN));
    const close = new Date(toIso_(cfg.SELECT_CLOSE));
    if (now < open || now >= close) {
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
