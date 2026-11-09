/**
 * 幾 A 幾 B｜版本四：Google 帳號與個人存檔（後端）
 * 以「存取網頁應用程式的使用者」身分執行；每位玩家一份自己的試算表，一局一列。
 * 本版沒有 setup：玩家在網頁按「建立我的遊戲存檔」後才建立。
 * 存檔位置記在 User Properties（GUESS_V4_BOOK_ID），請勿刪除。
 */
const BOOK_NAME = '猜數字版本四｜我的遊戲存檔';
const SHEET_NAME = '遊戲存檔';
const PROP_KEY = 'GUESS_V4_BOOK_ID';
const HEADERS = ['Google帳號', '局ID', '局數', '狀態', '開始時間', '更新時間', '猜測次數', '猜測歷程', '完整存檔JSON'];
const MAX_ROUNDS = 50;
const MAX_GUESSES = 100;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('幾A幾B版本四｜雲端紀錄')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 前端唯一入口：action = load / create / new / guess / debug。身分一律取自 Google，忽略前端傳入的帳號或試算表 ID。 */
function api(request) {
  let lock = null;
  try {
    const req = request || {};
    const action = String(req.action || '');
    if (['load', 'create', 'new', 'guess', 'debug'].indexOf(action) < 0) throw userError_('不支援的操作。');
    const email = identity_();

    lock = LockService.getUserLock();
    if (!lock.tryLock(20000)) throw userError_('雲端目前忙碌，請稍後再試。');

    const props = PropertiesService.getUserProperties();
    let id = props.getProperty(PROP_KEY);

    if (!id) {
      if (action === 'load') return notReady_(email, '請閱讀說明，再按「建立我的遊戲存檔」。');
      if (action !== 'create') throw userError_('尚未建立遊戲存檔，請先按「建立我的遊戲存檔」。');
      if (req.consent !== true) throw userError_('需要你按下確認，才會建立遊戲存檔。');
      const ss = SpreadsheetApp.create(BOOK_NAME);
      props.setProperty(PROP_KEY, ss.getId());   // 先記下位置，避免重複建立
      ensureSheet_(ss);
      return loaded_(email, ss, '已建立你的遊戲存檔。按「開新局」開始。');
    }

    const ss = openBook_(id);
    if (action === 'create') {
      ensureSheet_(ss);   // 已有存檔：沿用原檔，不新增
      return loaded_(email, ss, '已找到你原本的遊戲存檔，沒有重複建立。');
    }
    const sheet = getSheet_(ss);
    const list = readRounds_(sheet);

    if (action === 'load') {
      const msg = list.length ? '已讀取 ' + list.length + ' 局雲端紀錄。' : '這個帳號還沒有遊戲紀錄。按「開新局」開始。';
      return res_(email, ss, list, defaultSelected_(list), msg);
    }

    if (action === 'new') {
      const rid = requestId_(req);
      const existing = list.filter(function (x) { return x.game.rid === rid; })[0];
      if (existing) return res_(email, ss, list, existing.game.id, '新局先前已建立，沒有重複新增。');
      if (list.length >= MAX_ROUNDS) throw userError_('已達每個帳號 ' + MAX_ROUNDS + ' 局上限，不會自動刪除舊紀錄。');
      const now = new Date().toISOString();
      const number = list.reduce(function (m, x) { return Math.max(m, x.game.number); }, 0) + 1;
      const game = {
        id: Utilities.getUuid(), rid: rid, number: number, status: 'playing',
        answer: newAnswer_(), startedAt: now, updatedAt: now, endedAt: null,
        revision: 1, guesses: []
      };
      writeRow_(sheet, email, game, 0);
      list.push({ row: -1, game: game });
      return res_(email, ss, list, game.id, '新局已建立並存到雲端。');
    }

    // guess / debug 需要指定局（只會在這個帳號自己的存檔裡找）
    const item = list.filter(function (x) { return x.game.id === req.gameId; })[0];
    if (!item) throw userError_('找不到這一局，請重新讀取。');
    const game = item.game;

    if (action === 'debug') {
      // 唯讀：不改答案、不新增猜測、不開新局
      return res_(email, ss, list, game.id, '已從雲端讀取本局答案。', { debugAnswer: game.answer, debugGameId: game.id });
    }

    const rid = requestId_(req);
    if (game.guesses.some(function (g) { return g.rid === rid; })) {
      return res_(email, ss, list, game.id, '這次猜測先前已儲存，沒有重複新增。');
    }
    if (Number(req.revision) !== game.revision) {
      return res_(email, ss, list, game.id, '這一局已在其他地方更新，已讀回最新資料，請確認後再猜。', { conflict: true });
    }
    if (game.status !== 'playing') throw userError_('這一局已猜對，只能查看。');
    if (game.guesses.length >= MAX_GUESSES) throw userError_('這一局已達 ' + MAX_GUESSES + ' 次猜測上限。');

    const value = String(req.value === undefined || req.value === null ? '' : req.value).trim();
    if (!/^\d{4}$/.test(value) || new Set(value.split('')).size !== 4) {
      throw userError_('請輸入四位不重複的數字（可 0 開頭）；這次不會存入紀錄。');
    }

    const result = judge_(game.answer, value);
    const now = new Date().toISOString();
    game.guesses.push({ order: game.guesses.length + 1, value: value, result: result, at: now, rid: rid });
    game.updatedAt = now;
    game.revision += 1;
    if (result.a === 4) {
      game.status = 'won';
      game.endedAt = now;
    }
    writeRow_(sheet, email, game, item.row);
    const msg = result.a === 4 ? '猜對了！已儲存。' : '已儲存：' + value + ' → ' + result.a + 'A' + result.b + 'B';
    return res_(email, ss, list, game.id, msg);
  } catch (e) {
    if (e && e.user) return { ok: false, error: e.message };
    console.error(e);
    return { ok: false, uncertain: true, error: '雲端處理發生錯誤，尚未確認結果。請重試同一操作或重新讀取。' };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) { /* ignore */ } }
  }
}

/* ---------- 身分 ---------- */

function identity_() {
  let active = '', effective = '';
  try { active = Session.getActiveUser().getEmail(); } catch (e) { active = ''; }
  try { effective = Session.getEffectiveUser().getEmail(); } catch (e) { effective = ''; }
  active = String(active || '').trim().toLowerCase();
  effective = String(effective || '').trim().toLowerCase();
  if (!active || !effective) throw userError_('無法確認 Google 帳號，請檢查授權、資訊清單與部署的執行身分。');
  if (active !== effective) throw userError_('執行身分設定不正確：請將部署的執行身分改為「存取網頁應用程式的使用者」。');
  return active;
}

/* ---------- 遊戲規則 ---------- */

function newAnswer_() {
  const digits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
  for (let i = digits.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = digits[i]; digits[i] = digits[j]; digits[j] = t;
  }
  return digits.slice(0, 4).join('');
}

function judge_(answer, guess) {
  let a = 0, b = 0;
  for (let i = 0; i < 4; i++) {
    if (guess[i] === answer[i]) a++;
    else if (answer.indexOf(guess[i]) >= 0) b++;
  }
  return { a: a, b: b };
}

/* ---------- 存檔 ---------- */

function userError_(message) {
  const e = new Error(message);
  e.user = true;
  return e;
}

function requestId_(req) {
  const rid = String(req.requestId || '');
  if (!rid || rid.length > 80) throw userError_('操作識別碼不正確，請重新整理後再試。');
  return rid;
}

function openBook_(id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    // 不自動另建空表掩蓋資料遺失
    throw userError_('找不到你原本的遊戲存檔（可能已被刪除或無法存取）。請先恢復原檔或請教師協助。');
  }
}

function ensureSheet_(ss) {
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const first = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && first.getLastRow() === 0) {
      sheet = first;
      sheet.setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME);
    }
  }
  sheet.getRange('A:B').setNumberFormat('@');
  sheet.getRange('E:F').setNumberFormat('@');
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function getSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw userError_('找不到「' + SHEET_NAME + '」分頁，請先恢復原檔或請教師協助。');
  return sheet;
}

function readRounds_(sheet) {
  const last = sheet.getLastRow();
  const list = [];
  if (last < 2) return list;
  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][8]) list.push({ row: i + 2, game: JSON.parse(values[i][8]) });
  }
  list.sort(function (x, y) { return x.game.number - y.game.number; });
  return list;
}

function historyText_(game) {
  return game.guesses.map(function (g) {
    return g.order + '. ' + g.value + ' → ' + g.result.a + 'A' + g.result.b + 'B';
  }).join('\n');
}

function writeRow_(sheet, email, game, row) {
  const values = [
    email, game.id, game.number, game.status === 'won' ? '已猜對' : '未完成',
    game.startedAt, game.updatedAt, game.guesses.length, historyText_(game), JSON.stringify(game)
  ];
  if (row > 0) sheet.getRange(row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
  SpreadsheetApp.flush();
}

function publicGame_(game) {
  const out = {
    id: game.id, number: game.number, status: game.status,
    startedAt: game.startedAt, updatedAt: game.updatedAt, endedAt: game.endedAt,
    revision: game.revision,
    guesses: game.guesses.map(function (g) {
      return { order: g.order, value: g.value, result: g.result, at: g.at };
    })
  };
  // 進行中的答案不回傳；只有猜對後才公開
  if (game.status === 'won') out.answer = game.answer;
  return out;
}

function defaultSelected_(list) {
  const playing = list.filter(function (x) { return x.game.status === 'playing'; });
  const pick = playing.length ? playing[playing.length - 1] : list[list.length - 1];
  return pick ? pick.game.id : null;
}

function notReady_(email, message) {
  return {
    ok: true, code: email, ready: false, sheetUrl: '', rounds: [], selectedId: null,
    limits: { rounds: MAX_ROUNDS, guesses: MAX_GUESSES }, message: message
  };
}

function loaded_(email, ss, message) {
  const list = readRounds_(getSheet_(ss));
  return res_(email, ss, list, defaultSelected_(list), message);
}

function res_(email, ss, list, selectedId, message, extra) {
  const out = {
    ok: true, code: email, ready: true, sheetUrl: ss.getUrl(),
    rounds: list.map(function (x) { return publicGame_(x.game); }),
    selectedId: selectedId,
    limits: { rounds: MAX_ROUNDS, guesses: MAX_GUESSES },
    message: message
  };
  if (extra) Object.keys(extra).forEach(function (k) { out[k] = extra[k]; });
  return out;
}
