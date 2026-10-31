/**
 * 軽食 在庫管理アプリ（Google Apps Script / スプレッドシート連携）
 *
 * シート構成
 *   大会         : 大会ID / 大会名 / スタッフPIN / 有効 / 更新日時
 *   在庫設定     : 大会ID / 大会名 / 開催日 / 表示順 / 品目 / 在庫数 / 店舗 / 受取時間 / 賞味期限 / 受取方法 / 住所 / 電話   ← 開催日ごとに設定
 *   在庫切れログ : 記録日時 / 大会ID / 大会名 / 開催日 / 品目 / 入力者 / 状態 / 取消日時 / 取消者 / リクエストID / 在庫切れ時刻 / 備考
 *   日次残数     : 大会ID / 大会名 / 開催日 / 品目 / 在庫数 / 期限切れ数 / 消費数 / 在庫切れ時刻 / 入力者 / 記録日時
 */

const SPREADSHEET_ID = '1lBZbAOruZRdXtgd28UfQmsqQb0FgwO_kZZBTRCTA1p8';
const TZ = 'Asia/Tokyo';
const SH = { TOURN: '大会', STOCK: '在庫設定', LOG: '在庫切れログ', CLOSE: '日次残数' };
const HEAD = {
  [SH.TOURN]: ['大会ID', '大会名', 'スタッフPIN', '有効', '更新日時'],
  [SH.STOCK]: ['大会ID', '大会名', '開催日', '表示順', '品目', '在庫数', '店舗', '受取時間', '賞味期限', '受取方法', '住所', '電話'],
  [SH.LOG]: ['記録日時', '大会ID', '大会名', '開催日', '品目', '入力者', '状態', '取消日時', '取消者', 'リクエストID', '在庫切れ時刻', '備考'],
  [SH.CLOSE]: ['大会ID', '大会名', '開催日', '品目', '在庫数', '期限切れ数', '消費数', '在庫切れ時刻', '入力者', '記録日時'],
};
// 文字列として保持する列（日付・PIN・時刻の自動変換を防ぐ）
const TEXT_COLS = { [SH.TOURN]: [1, 3], [SH.STOCK]: [1, 3, 8, 9, 12], [SH.LOG]: [2, 4, 10, 12], [SH.CLOSE]: [1, 3, 8] };
const TIME_COLS = { [SH.TOURN]: [5], [SH.LOG]: [1, 8, 11], [SH.CLOSE]: [10] };

/* ---------- Web アプリ ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('軽食 在庫管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/** 外部に置いた画面（GitHub Pages など）からの API 呼び出し。本文は {"fn": 関数名, "args": [...]} の JSON */
const API = {
  getTournaments: getTournaments, staffLogin: staffLogin, getDayState: getDayState,
  markSoldOut: markSoldOut, cancelSoldOut: cancelSoldOut, submitClose: submitClose,
  adminList: adminList, adminSave: adminSave, adminSetActive: adminSetActive,
};
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('不明な操作です');
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- スプレッドシートのメニュー ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('軽食アプリ')
    .addItem('初期設定', 'setup')
    .addItem('管理者PINを表示', 'showAdminPin')
    .addItem('管理者PINを再発行', 'resetAdminPin')
    .addToUi();
}

function setup() {
  const book = book_();
  book.setSpreadsheetTimeZone(TZ);
  Object.keys(HEAD).forEach(function (name) {
    const sh = book.getSheetByName(name) || book.insertSheet(name);
    const h = HEAD[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
    (TEXT_COLS[name] || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
    (TIME_COLS[name] || []).forEach(function (c) { sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss'); });
  });
  closeColors_(book.getSheetByName(SH.CLOSE));
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', newPin_());
  showAdminPin();
}

/**
 * 日次残数シートの期限切れ数（F列）を自動で色分けする（条件付き書式）。
 * 薄い赤：在庫数−消費数（＝期限切れ数）が在庫数の半分以上 / 薄い青：期限切れ数が0 / 薄い緑：それ以外
 * （上のルールが優先される）
 */
function closeColors_(sh) {
  if (!sh || !sh.setConditionalFormatRules) return;
  const range = sh.getRange('F2:F');
  const red = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($E2),ISNUMBER($G2),$E2>0,$E2-$G2>=$E2/2)')
    .setBackground('#f4cccc').setRanges([range]).build();
  const blue = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($F2),$F2=0)')
    .setBackground('#cfe2f3').setRanges([range]).build();
  const green = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISNUMBER($F2)')
    .setBackground('#d9ead3').setRanges([range]).build();
  sh.setConditionalFormatRules([red, blue, green]);
}

function showAdminPin() {
  const msg = '管理者PIN: ' + PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* エディタから実行した場合 */ }
  Logger.log(msg);
}

function resetAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', newPin_());
  showAdminPin();
}

/* ---------- スタッフ用 API ---------- */

function getTournaments() {
  return {
    today: today_(),
    list: tournaments_().filter(function (t) { return t.active; }).map(function (t) { return { id: t.id, name: t.name }; }),
  };
}

function staffLogin(tid, pin) {
  const t = staff_(tid, pin);
  return { name: t.name, days: Object.keys(config_(t.id)).sort(), today: today_() };
}

function getDayState(tid, pin, date) {
  return state_(staff_(tid, pin), date);
}

/**
 * 在庫切れを記録する。clientAt は端末でボタンを押した時刻（オフライン再送時も押した時刻で記録するため）。
 * outAt（HH:MM）を指定すると、後から記録する場合の実際の在庫切れ時刻として扱う（5時より前は翌日）。note は備考。
 */
function markSoldOut(tid, pin, date, names, staff, reqId, clientAt, outAt, note) {
  const t = staff_(tid, pin);
  return withLock_(function () {
    const log = rows_(SH.LOG);
    if (reqId && log.some(function (r) { return str_(r[9]) === reqId; })) return state_(t, date); // 二重送信
    const valid = (config_(t.id)[date] || []).map(function (i) { return i.name; });
    const active = new Set(log.filter(function (r) {
      return str_(r[1]) === t.id && ds_(r[3]) === date && str_(r[6]) === '有効';
    }).map(function (r) { return str_(r[4]); }));
    const at = when_(clientAt);
    let outTime = at;
    const hm = str_(outAt);
    if (hm) {
      const m = hm.match(/^(\d{1,2}):(\d{2})$/);
      if (!m || +m[1] > 23 || +m[2] > 59) throw new Error('在庫切れ時刻が正しくありません');
      outTime = new Date(date + 'T' + ('0' + m[1]).slice(-2) + ':' + m[2] + ':00+09:00');
      if (+m[1] < 5) outTime = new Date(outTime.getTime() + 864e5); // 深夜（0〜4時台）は開催日の翌日
      if (outTime.getTime() > Math.max(at.getTime(), Date.now()) + 5 * 60000) throw new Error('在庫切れ時刻が記録時刻より後になっています');
    }
    const add = [];
    (names || []).forEach(function (n) {
      n = String(n);
      if (valid.indexOf(n) < 0) throw new Error('「' + n + '」はこの日の品目にありません');
      if (active.has(n)) return; // 他のスタッフが記録済み
      active.add(n);
      add.push([at, t.id, t.name, date, n, str_(staff), '有効', '', '', reqId || '', outTime, str_(note).slice(0, 200)]);
    });
    append_(SH.LOG, add);
    return state_(t, date);
  });
}

function cancelSoldOut(tid, pin, date, name, staff) {
  const t = staff_(tid, pin);
  return withLock_(function () {
    const sh = sheet_(SH.LOG);
    const log = rows_(SH.LOG);
    for (let i = log.length - 1; i >= 0; i--) {
      const r = log[i];
      if (str_(r[1]) === t.id && ds_(r[3]) === date && str_(r[4]) === name && str_(r[6]) === '有効') {
        sh.getRange(i + 2, 7, 1, 3).setValues([['取消', new Date(), str_(staff)]]);
      }
    }
    return state_(t, date);
  });
}

/** 期限切れ数を記録する。送信された品目のみ記録し、同じ品目の既存データは上書き */
function submitClose(tid, pin, date, entries, staff, clientAt) {
  const t = staff_(tid, pin);
  return withLock_(function () {
    const items = config_(t.id)[date] || [];
    if (!items.length) throw new Error('この日の品目が設定されていません');
    const outTime = {};
    state_(t, date).items.forEach(function (i) { outTime[i.name] = i.soldOut ? i.soldOut.time : ''; });
    const byName = {};
    (entries || []).forEach(function (e) { byName[String(e.item)] = e.remain; });
    const at = when_(clientAt);
    const add = [];
    items.forEach(function (it) {
      if (!Object.prototype.hasOwnProperty.call(byName, it.name)) return;
      const v = byName[it.name];
      const n = Number(v);
      if (v === '' || v == null || !Number.isInteger(n) || n < 0) throw new Error('「' + it.name + '」の期限切れ数が正しくありません');
      add.push([t.id, t.name, date, it.name, it.stock, n, it.stock - n, outTime[it.name], str_(staff), at]);
    });
    if (!add.length) throw new Error('期限切れ数を1つ以上入力してください');
    const sent = new Set(add.map(function (r) { return r[3]; }));
    replace_(SH.CLOSE, function (r) { return str_(r[0]) === t.id && ds_(r[2]) === date && sent.has(str_(r[3])); }, add);
    return state_(t, date);
  });
}

/* ---------- 管理者用 API ---------- */

function adminList(pin) {
  admin_(pin);
  return tournaments_().map(function (t) {
    const c = config_(t.id);
    return {
      id: t.id, name: t.name, staffPin: t.staffPin, active: t.active,
      days: Object.keys(c).sort().map(function (d) {
        return { date: d, items: c[d].map(function (i) { return { name: i.name, stock: i.stock, store: i.store, recv: i.recv, exp: i.exp, method: i.method, addr: i.addr, tel: i.tel }; }) };
      }),
    };
  });
}

function adminSave(pin, f) {
  admin_(pin);
  return withLock_(function () {
    const name = str_(f.name);
    const sp = str_(f.staffPin);
    if (!name) throw new Error('大会名を入力してください');
    if (!/^\d{4,8}$/.test(sp)) throw new Error('スタッフPINは数字4〜8桁で入力してください');
    const days = f.days || [];
    if (!days.length) throw new Error('開催日を1日以上登録してください');
    const id = str_(f.id) || ('T' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss'));
    const seen = new Set();
    const stockRows = [];
    days.forEach(function (d) {
      const date = str_(d.date);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('開催日を入力してください');
      if (seen.has(date)) throw new Error(date + ' が重複しています');
      seen.add(date);
      const items = (d.items || []).filter(function (i) { return str_(i.name) || str_(i.stock); });
      if (!items.length) throw new Error(date + ' の品目を1つ以上登録してください');
      const names = new Set();
      items.forEach(function (i, k) {
        const n = str_(i.name);
        const s = Number(i.stock);
        if (!n) throw new Error(date + ' に品目名が空の行があります');
        if (names.has(n)) throw new Error(date + ' の「' + n + '」が重複しています');
        names.add(n);
        if (str_(i.stock) === '' || !Number.isInteger(s) || s < 0) throw new Error(date + ' の「' + n + '」の在庫数が正しくありません');
        const recv = str_(i.recv);
        if (recv && !/^\d{1,2}:\d{2}$/.test(recv)) throw new Error(date + ' の「' + n + '」の受取時間が正しくありません');
        const exp = str_(i.exp).replace('T', ' ');
        if (exp && !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(exp)) throw new Error(date + ' の「' + n + '」の賞味期限が正しくありません');
        stockRows.push([id, name, date, k + 1, n, s, str_(i.store), recv, exp, str_(i.method), str_(i.addr), str_(i.tel)]);
      });
    });
    replace_(SH.TOURN, function (r) { return str_(r[0]) === id; }, [[id, name, sp, f.active ? '有効' : '無効', new Date()]]);
    replace_(SH.STOCK, function (r) { return str_(r[0]) === id; }, stockRows);
    return adminList(pin);
  });
}

/** 大会の表示・非表示（スタッフのログイン画面に出すか）だけを切り替える */
function adminSetActive(pin, id, active) {
  admin_(pin);
  return withLock_(function () {
    const v = rows_(SH.TOURN);
    for (let i = 0; i < v.length; i++) {
      if (str_(v[i][0]) === str_(id)) {
        sheet_(SH.TOURN).getRange(i + 2, 4, 1, 2).setValues([[active ? '有効' : '無効', new Date()]]);
        return adminList(pin);
      }
    }
    throw new Error('大会が見つかりません');
  });
}

/* ---------- 内部処理 ---------- */

function state_(t, date) {
  const items = config_(t.id)[date] || [];
  const out = {};
  rows_(SH.LOG).forEach(function (r) {
    if (str_(r[1]) === t.id && ds_(r[3]) === date && str_(r[6]) === '有効') {
      const n = str_(r[4]);
      if (out[n]) return;
      const o = r[10] instanceof Date ? r[10] : r[0]; // 在庫切れ時刻（旧データは記録日時）
      const late = r[10] instanceof Date && r[0] instanceof Date && Math.abs(r[0] - r[10]) >= 5 * 60000;
      out[n] = { time: o instanceof Date ? hm_(o) : str_(o), by: str_(r[5]), late: late, at: late ? hm_(r[0]) : '', note: str_(r[11]) };
    }
  });
  const remain = {};
  let closed = null;
  rows_(SH.CLOSE).forEach(function (r) {
    if (str_(r[0]) === t.id && ds_(r[2]) === date) {
      remain[str_(r[3])] = Number(r[5]);
      if (r[9] instanceof Date && (!closed || r[9] > closed.d)) closed = { d: r[9], by: str_(r[8]) };
    }
  });
  return {
    date: date,
    items: items.map(function (it) {
      return {
        name: it.name, stock: it.stock, store: it.store, recv: it.recv, exp: it.exp, method: it.method, addr: it.addr, tel: it.tel,
        soldOut: out[it.name] || null,
        remain: Object.prototype.hasOwnProperty.call(remain, it.name) ? remain[it.name] : null,
      };
    }),
    closed: closed ? { at: Utilities.formatDate(closed.d, TZ, 'M/d HH:mm'), by: closed.by } : null,
  };
}

function tournaments_() {
  return rows_(SH.TOURN).filter(function (r) { return str_(r[0]); }).map(function (r) {
    return { id: str_(r[0]), name: str_(r[1]), staffPin: str_(r[2]), active: str_(r[3]) !== '無効' };
  });
}

function config_(tid) {
  const map = {};
  rows_(SH.STOCK).forEach(function (r) {
    if (str_(r[0]) !== tid) return;
    const d = ds_(r[2]);
    (map[d] = map[d] || []).push({ name: str_(r[4]), stock: Number(r[5]) || 0, order: Number(r[3]) || 0,
      store: str_(r[6]), recv: r[7] instanceof Date ? hm_(r[7]) : str_(r[7]), exp: r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, 'yyyy-MM-dd HH:mm') : str_(r[8]),
      method: str_(r[9]), addr: str_(r[10]), tel: str_(r[11]) });
  });
  Object.keys(map).forEach(function (d) { map[d].sort(function (a, b) { return a.order - b.order; }); });
  return map;
}

function staff_(tid, pin) {
  const t = tournaments_().filter(function (x) { return x.id === tid && x.active; })[0];
  if (!t) throw new Error('大会が見つかりません');
  if (t.staffPin !== String(pin).trim()) throw new Error('PINが正しくありません');
  return t;
}

function admin_(pin) {
  const p = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!p || p !== String(pin).trim()) throw new Error('管理者PINが正しくありません');
}

function sheet_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。メニュー「軽食アプリ → 初期設定」を実行してください');
  return sh;
}

function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, HEAD[name].length).getValues();
}

function ensureRows_(sh, lastNeeded) {
  if (lastNeeded > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), lastNeeded - sh.getMaxRows() + 100);
    if (sh.getName() === SH.CLOSE) closeColors_(sh); // 行を増やしたら色分けの範囲も広げる
  }
}

function append_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name);
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + rows.length - 1);
  sh.getRange(start, 1, rows.length, HEAD[name].length).setValues(rows);
}

/** 条件に合う行を削除し、add を末尾に追加して書き戻す */
function replace_(name, removeFn, add) {
  const sh = sheet_(name);
  const w = HEAD[name].length;
  const last = sh.getLastRow();
  const keep = rows_(name).filter(function (r) { return !removeFn(r); }).concat(add);
  if (last >= 2) sh.getRange(2, 1, last - 1, w).clearContent();
  if (keep.length) {
    ensureRows_(sh, keep.length + 1);
    sh.getRange(2, 1, keep.length, w).setValues(keep);
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('混み合っています。もう一度押してください');
  try { return fn(); } finally { lock.releaseLock(); }
}

function when_(clientAt) {
  const now = new Date();
  const c = Number(clientAt);
  if (c && c <= now.getTime() + 5 * 60000 && c >= now.getTime() - 24 * 3600000) return new Date(c);
  return now;
}

function book_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function str_(v) { return v == null ? '' : String(v).trim(); }
function ds_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : str_(v); }
function hm_(d) { return Utilities.formatDate(d, TZ, 'HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function newPin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
