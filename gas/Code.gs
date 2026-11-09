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
const SH = { TOURN: '大会', STOCK: '在庫設定', LOG: '在庫切れログ', CLOSE: '日次残数', POP: 'POP設定', TABLE: '卓', ORDER: '注文', NOTIFY: '通知待ち' };
const HEAD = {
  [SH.TOURN]: ['大会ID', '大会名', 'スタッフPIN', '有効', '更新日時', '注文受付', '受付時間'],
  [SH.STOCK]: ['大会ID', '大会名', '開催日', '表示順', '品目', '在庫数', '店舗', '受取時間', '賞味期限', '受取方法', '住所', '電話'],
  [SH.LOG]: ['記録日時', '大会ID', '大会名', '開催日', '品目', '入力者', '状態', '取消日時', '取消者', 'リクエストID', '在庫切れ時刻', '備考'],
  [SH.CLOSE]: ['大会ID', '大会名', '開催日', '品目', '在庫数', '期限切れ数', '消費数', '在庫切れ時刻', '入力者', '記録日時'],
  [SH.POP]: ['品目', '価格', '写真', '更新日時'], // デジタルPOP用。品目名（受取時間などの（ ）を外した名前）ごとに1行。大会をまたいで使い回す
  [SH.TABLE]: ['大会ID', '卓番号', '席数', 'トークン', '有効', '更新日時'], // 卓上注文のQR。トークンを作り直すと過去のQRが一斉に失効する
  [SH.ORDER]: ['注文ID', '明細番号', '大会ID', '大会名', '開催日', '卓', '席', '品目', '個数', 'ドリチケ', '状態', '取消理由', '照会キー', 'リクエストID', '受付日時', '担当者', '担当端末', '担当日時', '完了日時', '操作履歴'],
  [SH.NOTIFY]: ['通知ID', '種別', '本文', '状態', '作成日時', '送信日時', '試行回数', 'エラー'], // 外部通知はロックの外で、ここを経由して送る
};
// 文字列として保持する列（日付・PIN・時刻の自動変換を防ぐ）
const TEXT_COLS = {
  [SH.TOURN]: [1, 3, 7], [SH.STOCK]: [1, 3, 8, 9, 12], [SH.LOG]: [2, 4, 10, 12], [SH.CLOSE]: [1, 3, 8], [SH.POP]: [1, 3],
  [SH.TABLE]: [1, 2, 4], [SH.ORDER]: [1, 3, 5, 6, 7, 13, 14], [SH.NOTIFY]: [1],
};
const TIME_COLS = {
  [SH.TOURN]: [5], [SH.LOG]: [1, 8, 11], [SH.CLOSE]: [10], [SH.POP]: [4],
  [SH.TABLE]: [6], [SH.ORDER]: [15, 18, 19], [SH.NOTIFY]: [5, 6],
};
/* 注文の設定（必要なら変える） */
const ORD = {
  MAX_QTY: 5,        // 1注文の合計個数
  MAX_KINDS: 5,      // 1注文の品目の種類
  SEAT_SEC: 60,      // 同じ卓×席の間隔（秒）
  TABLE_MAX: 5,      // 1つの卓からの注文数
  TABLE_SEC: 600,    // その集計期間（秒）
  OPEN_MAX: 50,      // 未完了の注文がこれを超えたら新規受付を止める
  LIST_MAX: 50,      // スタッフ画面に返す未完了注文の上限
};

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
  getPopState: getPopState, getPopImages: getPopImages, adminPopList: adminPopList, adminPopSave: adminPopSave,
  getOrderMenu: getOrderMenu, placeOrder: placeOrder, getOrder: getOrder, setOrder: setOrder,
  adminTables: adminTables, adminSaveTable: adminSaveTable, adminOrderOpen: adminOrderOpen,
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
    .addItem('通知のトリガーを設定', 'setupTriggers')
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
  tournaments_().forEach(function (t) { tournSheet_(t.id, t.name); }); // 既存の大会にも集計シートを用意
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
  const t = staff_(tid, pin);
  const st = state_(t, date);
  st.orders = openOrders_(tid, date);   // 未完了の注文だけ。完了履歴は返さない（取得が重くならないように）
  st.orderOpen = t.orderOpen;
  return st;
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
    // 注文受付・受付時間は大会編集では触らない。列数を HEAD に合わせないと replace_ が書き戻しに失敗し、大会一覧が消える
    const prev = rows_(SH.TOURN).filter(function (r) { return str_(r[0]) === id; })[0] || [];
    replace_(SH.TOURN, function (r) { return str_(r[0]) === id; }, [[id, name, sp, f.active ? '有効' : '無効', new Date(), str_(prev[5]), str_(prev[6])]]);
    replace_(SH.STOCK, function (r) { return str_(r[0]) === id; }, stockRows);
    try { tournSheet_(id, name); } catch (e) { console.warn('大会シートを作成できませんでした: ' + e); } // 失敗しても保存は有効
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

/* ---------- デジタルPOP（お客様向けの在庫表示） ---------- */

/**
 * POP画面用。PIN不要（品目名・価格・在庫の有無だけを返す）。表示中の大会のみ。
 * 営業日は5時切り替え（0〜4時台は前日の開催日として扱う）。date を渡すとその日を返す（確認用）。
 */
function getPopState(tid, date) {
  const list = tournaments_().filter(function (t) { return t.active; });
  if (!tid) return { tournaments: list.map(function (t) { return { id: t.id, name: t.name }; }) };
  const t = list.filter(function (x) { return x.id === tid; })[0];
  if (!t) throw new Error('大会が見つかりません（非表示の大会は表示できません）');
  const now = new Date();
  const d = date || Utilities.formatDate(new Date(now.getTime() - 5 * 3600000), TZ, 'yyyy-MM-dd');
  const st = state_(t, d);
  const master = popMaster_(false);
  const items = st.items.map(function (it) {
    return { base: popBase_(it.name), store: it.store, recv: it.recv, exp: it.exp, sold: !!it.soldOut };
  });
  const info = {};
  items.forEach(function (it) { if (master[it.base] && !info[it.base]) info[it.base] = { price: master[it.base].price, v: master[it.base].v, img: master[it.base].has }; });
  return { name: t.name, date: d, now: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm'), items: items, info: info };
}

/** POP画面用：品目名の写真（data URL）を返す */
function getPopImages(names) {
  const m = popMaster_(true), out = {};
  (names || []).slice(0, 100).forEach(function (n) { n = str_(n); if (m[n] && m[n].img) out[n] = m[n].img; });
  return out;
}

/** 管理画面用：大会に含まれる品目名ごとの価格・写真 */
function adminPopList(pin, tid) {
  admin_(pin);
  const names = [];
  rows_(SH.STOCK).forEach(function (r) {
    if (str_(r[0]) !== tid) return;
    const b = popBase_(r[4]);
    if (b && names.indexOf(b) < 0) names.push(b);
  });
  const m = popMaster_(true);
  return names.map(function (n) { return { name: n, price: m[n] ? m[n].price : '', img: m[n] ? m[n].img : '' }; });
}

/** 管理画面用：価格・写真を保存。img は undefined/null で変更なし、'' で削除、data URL で登録 */
function adminPopSave(pin, name, price, img) {
  admin_(pin);
  name = popBase_(name);
  if (!name) throw new Error('品目名がありません');
  price = str_(price).replace(/[^\d]/g, '');
  if (img != null && img !== '' && !/^data:image\/(jpeg|png|webp);base64,/.test(img)) throw new Error('写真の形式が正しくありません');
  if (img && img.length > 49000) throw new Error('写真のサイズが大きすぎます');
  return withLock_(function () {
    const sh = sheet_(SH.POP);
    const v = rows_(SH.POP);
    let i = v.findIndex(function (r) { return str_(r[0]) === name; });
    const cur = i >= 0 ? v[i] : [name, '', '', ''];
    const row = [name, price === '' ? '' : Number(price), img == null ? str_(cur[2]) : img, new Date()];
    if (i < 0) { ensureRows_(sh, v.length + 2); i = v.length; }
    sh.getRange(i + 2, 1, 1, 4).setValues([row]);
    return { name: name, price: row[1], img: row[2] };
  });
}

/** 品目名から受取時間・店舗などの末尾の（ ）を外す */
/* ---------- 卓上注文 ----------
 * プレイヤーは PIN を持たないので placeOrder と getOrder は PIN 不要。そのぶん次で守る。
 *   ・卓ごとのトークン（QRのURLに入れる。作り直せば過去のQRが一斉に失効する）
 *   ・卓・席が「卓」シートに実在すること
 *   ・大会が有効かつ「注文受付」が受付中、受付時間内であること
 *   ・個数・種類・回数の制限
 *   ・品目が本当に販売中かをサーバーで判定し直すこと（端末から来た値は信じない）
 * 注文は在庫を減らさない。依頼として記録するだけで、在庫切れの記録とは分ける。
 */

function bizDate_() { return Utilities.formatDate(new Date(new Date().getTime() - 5 * 3600000), TZ, 'yyyy-MM-dd'); }
function addDay_(d, n) { const x = new Date(d + 'T00:00:00+09:00'); x.setUTCDate(x.getUTCDate() + n); return Utilities.formatDate(x, TZ, 'yyyy-MM-dd'); }
/** 受取時刻の実際の日時。0〜4時台は開催日の翌日として扱う（画面側と同じ規則） */
/** 'H:MM' も 'HH:MM' も受ける。桁を揃えないと文字列比較が狂う（9:00 が 23:00 より後になる） */
function hm2_(v) { const m = str_(v).match(/^(\d{1,2}):(\d{2})$/); return m && +m[1] < 24 ? String(+m[1]).padStart(2, '0') + ':' + m[2] : ''; }
function recvAt_(date, hm) { const t = hm2_(hm); if (!t) return ''; return (Number(t.slice(0, 2)) < 5 ? addDay_(date, 1) : date) + ' ' + t; }
function rnd_(n, chars) {
  const cs = chars || 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < n; i++) s += cs.charAt(Math.floor(Math.random() * cs.length));
  return s;
}

/** いま注文できる品目（POPのまとめ名）。POPの「販売中」と同じ判定をサーバーでやり直す */
function onSale_(t, date) {
  const st = state_(t, date);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const map = {};
  st.items.forEach(function (it) {
    if (it.soldOut) return;
    const r = recvAt_(date, it.recv);
    if (r && r > now) return;              // まだ受け取っていない
    if (it.exp && !(now < it.exp)) return; // 賞味期限を過ぎた
    map[popBase_(it.name)] = true;
  });
  return map;
}

function tables_(tid) {
  return rows_(SH.TABLE).filter(function (r) { return str_(r[0]) === tid; }).map(function (r) {
    return { table: str_(r[1]), seats: Number(r[2]) || 0, token: str_(r[3]), active: str_(r[4]) !== '無効' };
  });
}

/** 受付時間（「11:00-23:00」形式。空なら時刻では絞らない）。終わりが始まりより小さいときは日をまたぐ */
function inHours_(hours) {
  const s = str_(hours).replace(/[！-～]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  if (!s) return true;                                   // 空欄は時刻で絞らない
  const m = s.match(/^(\d{1,2}:\d{2})\s*[-~〜]\s*(\d{1,2}:\d{2})$/);
  const a = m && hm2_(m[1]), b = m && hm2_(m[2]);
  if (!a || !b) return false;                            // 書式が違うときは開けない（書き損じで開きっぱなしにしない）
  const now = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  return a <= b ? (now >= a && now <= b) : (now >= a || now <= b);   // 桁を揃えてから比べる
}

function checkOrderOpen_(t) {
  if (!t.orderOpen) throw new Error('ただいま注文を受け付けていません');
  if (!inHours_(t.orderHours)) throw new Error('注文の受付時間外です（' + t.orderHours + '）');
}

/** 端末から来た品目を検証してそろえる。同じ品目は合算する */
function normItems_(items) {
  if (!items || !items.length) throw new Error('品目が選ばれていません');
  const map = {}, order = [];
  items.forEach(function (x) {
    const n = str_(x && x.name);
    const q = Number(x && x.qty);
    if (!n) throw new Error('品目名がありません');
    if (n.length > 60) throw new Error('品目名が長すぎます');
    if (!(q >= 1 && q <= ORD.MAX_QTY && q === Math.floor(q))) throw new Error('個数は1〜' + ORD.MAX_QTY + 'の整数で指定してください');
    if (!map[n]) { map[n] = 0; order.push(n); }
    map[n] += q;
  });
  if (order.length > ORD.MAX_KINDS) throw new Error('一度に選べるのは' + ORD.MAX_KINDS + '種類までです');
  let total = 0;
  const out = order.map(function (n) { total += map[n]; return { name: n, qty: map[n] }; });
  if (total > ORD.MAX_QTY) throw new Error('一度に注文できるのは合計' + ORD.MAX_QTY + '個までです');
  return out;
}

function orderRows_() { return rows_(SH.ORDER); }
function rowToOrder_(r) {
  return { id: str_(r[0]), no: Number(r[1]) || 0, tid: str_(r[2]), date: ds_(r[4]), table: str_(r[5]), seat: str_(r[6]),
    name: str_(r[7]), qty: Number(r[8]) || 0, ticket: str_(r[9]), state: str_(r[10]), reason: str_(r[11]),
    at: r[14] instanceof Date ? hm_(r[14]) : '', by: str_(r[15]), takenAt: r[17] instanceof Date ? hm_(r[17]) : '' };
}
/** リクエストIDは大会の中でだけ探す（他の大会の注文を引き当てて照会キーを渡さないため） */
function findOrderByReq_(tid, reqId) {
  const hit = orderRows_().filter(function (r) { return str_(r[2]) === tid && str_(r[13]) === reqId; });
  if (!hit.length) return null;
  return { id: str_(hit[0][0]), key: str_(hit[0][12]), lines: hit.map(rowToOrder_) };
}
/** 未完了（受付・配達中）の注文。新しい順に ORD.LIST_MAX 件まで */
function openOrders_(tid, date) {
  const out = [];
  orderRows_().forEach(function (r) {
    if (str_(r[2]) !== tid || ds_(r[4]) !== date) return;
    const s = str_(r[10]);
    if (s !== '受付' && s !== '配達中') return;
    out.push(rowToOrder_(r));
  });
  return out.slice(-ORD.LIST_MAX);
}

function limitCheck_(tid, date, table, seat) {
  const now = Date.now();
  const tableIds = {};   // 卓の制限は「注文の数」で数える。明細の数で数えると5品目の1回で上限に達してしまう
  let seatLast = 0, open = 0;
  orderRows_().forEach(function (r) {
    if (str_(r[2]) !== tid || ds_(r[4]) !== date) return;
    const s = str_(r[10]);
    if (s === '受付' || s === '配達中') open++;
    if (!(r[14] instanceof Date)) return;
    const ms = r[14].getTime();
    if (str_(r[5]) === table) {
      if (now - ms <= ORD.TABLE_SEC * 1000) tableIds[str_(r[0])] = 1;
      if (str_(r[6]) === seat && ms > seatLast) seatLast = ms;
    }
  });
  if (open >= ORD.OPEN_MAX) throw new Error('注文が混み合っています。少し待ってからお願いします');
  if (seatLast && now - seatLast < ORD.SEAT_SEC * 1000) throw new Error('続けての注文は' + ORD.SEAT_SEC + '秒あけてください');
  if (Object.keys(tableIds).length >= ORD.TABLE_MAX) throw new Error('この卓からの注文が続いています。少し待ってからお願いします');
}
/** 注文IDは重ならないものを選ぶ（同じ秒に乱数が重なると、別の注文が同じIDになる） */
function newOrderId_() {
  const used = {};
  orderRows_().forEach(function (r) { used[str_(r[0])] = 1; });
  for (let i = 0; i < 20; i++) {
    const id = 'O' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + rnd_(5);
    if (!used[id]) return id;
  }
  throw new Error('注文を記録できませんでした。もう一度お願いします');
}

/** 注文画面用。QRの正しさをここで確かめ、席数と受付状況も返す。PIN不要 */
function getOrderMenu(tid, table, token) {
  tid = str_(tid); table = str_(table); token = str_(token);
  const t = tournaments_().filter(function (x) { return x.id === tid && x.active; })[0];
  if (!t) throw new Error('この大会は注文を受け付けていません');
  const tb = tables_(tid).filter(function (x) { return x.table === table && x.active; })[0];
  if (!tb || !tb.token || tb.token !== token) throw new Error('このQRは使えません。スタッフにお知らせください');
  const st = getPopState(tid, '');
  st.seats = tb.seats;
  st.open = !!(t.orderOpen && inHours_(t.orderHours));
  st.hours = t.orderHours;
  return st;
}

/** 注文を受け付ける。PIN不要 */
function placeOrder(tid, table, token, seat, items, reqId) {
  tid = str_(tid); table = str_(table); token = str_(token); seat = str_(seat); reqId = str_(reqId);
  if (!reqId || reqId.length > 40) throw new Error('リクエストIDがありません');
  if (table.length > 8 || seat.length > 8) throw new Error('卓・席の指定が正しくありません');

  // QRの確認を先にやる。これより前に再送の判定を置くと、リクエストIDを言い当てるだけで
  // トークン無しに他人の注文（照会キーを含む）を引き出せてしまう
  const check = function () {
    const t = tournaments_().filter(function (x) { return x.id === tid && x.active; })[0];
    if (!t) throw new Error('この大会は注文を受け付けていません');
    const tb = tables_(tid).filter(function (x) { return x.table === table && x.active; })[0];
    if (!tb || !tb.token || tb.token !== token) throw new Error('このQRは使えません。スタッフにお知らせください');
    const sn = Number(seat);
    if (!(sn >= 1 && sn <= tb.seats && sn === Math.floor(sn))) throw new Error('席番号が正しくありません');
    checkOrderOpen_(t);
    return { t: t, seat: String(sn) };   // 席は正規化して持つ（"01" と "1" で間隔制限をすり抜けさせない）
  };
  check();
  const dup = findOrderByReq_(tid, reqId);
  if (dup) return dup;                                   // 保存できたのに応答が消えた場合の再送
  const list = normItems_(items);

  return withLock_(function () {
    // ロックを待っている間に、受付停止・トークン再発行・在庫切れ・期限到来が起きていることがある。
    // 保存の直前にもう一度ぜんぶ確かめる
    const c = check();
    const again = findOrderByReq_(tid, reqId);
    if (again) return again;
    const date = bizDate_();
    const sale = onSale_(c.t, date);
    list.forEach(function (x) { if (!sale[x.name]) throw new Error('「' + x.name + '」はただいま注文できません'); });
    limitCheck_(tid, date, table, c.seat);
    const price = popMaster_(false);
    const now = new Date();
    const id = newOrderId_();
    const key = rnd_(12);
    const rows = list.map(function (x, i) {
      return [id, i + 1, tid, c.t.name, date, table, c.seat, x.name, x.qty,
        price[x.name] ? str_(price[x.name].price) : '', '受付', '', key, reqId, now, '', '', '', '', ''];
    });
    append_(SH.ORDER, rows);
    const label = list.map(function (x) { return x.name + (x.qty > 1 ? '×' + x.qty : ''); }).join('、');
    notifyPush_('注文', table + '番テーブルの' + c.seat + '番シートのプレイヤーに ' + label + ' を持っていってください。');
    return { id: id, key: key, lines: rows.map(rowToOrder_) };
  });
}

/** プレイヤーが自分の注文の状態を見る。照会キーが合うものだけ返す。PIN不要 */
function getOrder(id, key) {
  id = str_(id); key = str_(key);
  if (!id || !key) throw new Error('注文が指定されていません');
  const hit = orderRows_().filter(function (r) { return str_(r[0]) === id && str_(r[12]) === key; });
  if (!hit.length) throw new Error('注文が見つかりません');
  return { id: id, lines: hit.map(rowToOrder_) };
}

/** スタッフが状態を変える。配達中 / 完了 / 取消。no を 0 にするとその注文の全明細 */
function setOrder(tid, pin, id, no, state, staff, device, reason) {
  const t = staff_(tid, pin);
  id = str_(id); state = str_(state); staff = str_(staff); device = str_(device); reason = str_(reason);
  if (['配達中', '完了', '取消'].indexOf(state) < 0) throw new Error('状態が正しくありません');
  if (!staff) throw new Error('名前がありません');
  const n = Number(no) || 0;
  return withLock_(function () {
    const sh = sheet_(SH.ORDER);
    const all = orderRows_();
    const now = new Date();
    let held = null, done = 0, already = 0;
    all.forEach(function (r, i) {
      if (str_(r[0]) !== id || str_(r[2]) !== tid) return;
      if (n && Number(r[1]) !== n) return;
      const cur = str_(r[10]);
      if (cur === '完了' || cur === '取消') return;                 // 終わった明細は動かさない
      if (state === '配達中' && cur === '配達中') {
        if (str_(r[16]) !== device) { held = str_(r[15]); return; } // 先に押した人が担当
        already++; return;                                          // 同じ端末からの再送。担当日時も通知もそのまま
      }
      const line = Utilities.formatDate(now, TZ, 'HH:mm') + ' ' + state + ' ' + staff;
      const hist = str_(r[19]) ? str_(r[19]) + ' / ' + line : line;
      const row = r.slice();
      row[10] = state;
      if (state === '取消') row[11] = reason;
      if (state === '配達中') { row[15] = staff; row[16] = device; row[17] = now; }
      if (state === '完了' || state === '取消') { row[18] = now; if (!str_(row[15])) { row[15] = staff; row[16] = device; } }
      row[19] = hist;
      sh.getRange(i + 2, 1, 1, HEAD[SH.ORDER].length).setValues([row]);
      done++;
    });
    if (!done && held) throw new Error(held + 'が配達中です');
    if (!done && already) return { id: id, orders: openOrders_(tid, bizDate_()) }; // 再送。何も変えずに成功を返す
    if (!done) throw new Error('変更できる注文がありません');
    if (state === '配達中') notifyPush_('担当', staff + 'が届けに行きます。');
    return { id: id, orders: openOrders_(tid, bizDate_()) };
  });
}

/* ---------- 卓の管理（管理画面） ---------- */

function adminTables(pin, tid) {
  admin_(pin);
  const t = tournaments_().filter(function (x) { return x.id === str_(tid); })[0];
  if (!t) throw new Error('大会が見つかりません');
  return { tables: tables_(str_(tid)), orderOpen: t.orderOpen, orderHours: t.orderHours };
}

/** 卓を追加・変更する。token を 'new' にするとトークンを作り直す（過去のQRが失効する） */
function adminSaveTable(pin, tid, table, seats, active, token) {
  admin_(pin);
  tid = str_(tid); table = str_(table);
  if (!table || table.length > 8) throw new Error('卓番号を8文字以内で入力してください');
  const n = Number(seats);
  if (!(n >= 1 && n <= 20 && n === Math.floor(n))) throw new Error('席数は1〜20の整数で入力してください');
  return withLock_(function () {
    const sh = sheet_(SH.TABLE);
    const all = rows_(SH.TABLE);
    let idx = -1;
    all.forEach(function (r, i) { if (str_(r[0]) === tid && str_(r[1]) === table) idx = i; });
    const old = idx >= 0 ? str_(all[idx][3]) : '';
    const tk = (str_(token) === 'new' || !old) ? rnd_(16) : old;
    const row = [tid, table, n, tk, str_(active) === '無効' ? '無効' : '有効', new Date()];
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, HEAD[SH.TABLE].length).setValues([row]);
    else append_(SH.TABLE, [row]);
    return tables_(tid);
  });
}

/** 注文受付の開閉と受付時間。緊急停止はこれ */
function adminOrderOpen(pin, tid, open, hours) {
  admin_(pin);
  tid = str_(tid);
  return withLock_(function () {
    const sh = sheet_(SH.TOURN);
    const all = rows_(SH.TOURN);
    let idx = -1;
    all.forEach(function (r, i) { if (str_(r[0]) === tid) idx = i; });
    if (idx < 0) throw new Error('大会が見つかりません');
    const row = all[idx].slice();
    row[5] = open ? '受付中' : '停止';
    if (hours !== undefined && hours !== null) row[6] = str_(hours);
    row[4] = new Date();
    sh.getRange(idx + 2, 1, 1, HEAD[SH.TOURN].length).setValues([row]);
    return { orderOpen: !!open, orderHours: str_(row[6]) };
  });
}

/* ---------- 外部通知 ----------
 * ロックの中で外部に通信すると全部の書き込みが詰まるので、「通知待ち」シートに積むだけにして、
 * 1分おきの時間主導トリガー sendPending() が送る。失敗は試行回数を増やして次回に回す。
 * 送り先はスクリプトプロパティで切り替える（トークンはコードに書かない）。
 *   NOTIFY_KIND = discord | slack | line | mail | （空＝送らない）
 *   DISCORD_WEBHOOK / SLACK_WEBHOOK / LINE_TOKEN + LINE_TO / NOTIFY_MAIL
 */

function notifyPush_(kind, text) {
  try {
    append_(SH.NOTIFY, [['N' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + rnd_(3, '0123456789'), kind, text, '未送信', new Date(), '', 0, '']]);
  } catch (e) { /* 通知の失敗で注文を失わない */ }
}

/** 1分おきのトリガーから呼ぶ。未送信をまとめて送る */
function sendPending() {
  const p = PropertiesService.getScriptProperties();
  const base = str_(p.getProperty('NOTIFY_KIND'));
  const sh = sheet_(SH.NOTIFY);
  // 1. ロックを取って送る分を「送信中」にする。こうしないと、1分おきのトリガーが重なったとき
  //    同じ通知を2回送る。また、注文の書き込みと行がぶつからない
  const jobs = withLock_(function () {
    const all = rows_(SH.NOTIFY), out = [];
    all.forEach(function (r, i) {
      if (str_(r[3]) !== '未送信' || out.length >= 20) return;
      // 種別ごとに送り先を変えられる。例：NOTIFY_KIND=discord、NOTIFY_KIND_期限=line
      // （LINEの無料枠は月200通なので、数の多い注文通知には足りない。期限通知だけなら収まる）
      const kind = str_(p.getProperty('NOTIFY_KIND_' + str_(r[1]))) || base;
      if (!kind) return;
      const tries = Number(r[6]) || 0;
      if (tries >= 5) { sh.getRange(i + 2, 4).setValue('失敗'); return; }
      sh.getRange(i + 2, 4).setValue('送信中');
      out.push({ i: i, kind: kind, text: str_(r[2]), tries: tries, row: r.slice() });
    });
    return out;
  });
  // 2. 送信そのものはロックの外でやる。外部通信で詰まっても注文の受付が止まらない
  jobs.forEach(function (j) {
    try { sendOne_(j.kind, p, j.text); j.okFlag = true; }
    catch (e) { j.okFlag = false; j.err = String(e && e.message ? e.message : e).slice(0, 200); }
  });
  // 3. 結果を書き戻す
  withLock_(function () {
    jobs.forEach(function (j) {
      const row = j.row;
      row[3] = j.okFlag ? '送信済' : '未送信';
      row[5] = j.okFlag ? new Date() : '';
      row[6] = j.tries + 1;
      row[7] = j.okFlag ? '' : j.err;
      sh.getRange(j.i + 2, 1, 1, HEAD[SH.NOTIFY].length).setValues([row]);
    });
  });
}

function sendOne_(kind, p, text) {
  if (kind === 'discord') {
    const url = str_(p.getProperty('DISCORD_WEBHOOK'));
    if (!url) throw new Error('DISCORD_WEBHOOK が未設定です');
    post_(url, { content: text });
  } else if (kind === 'slack') {
    const url = str_(p.getProperty('SLACK_WEBHOOK'));
    if (!url) throw new Error('SLACK_WEBHOOK が未設定です');
    post_(url, { text: text });
  } else if (kind === 'line') {
    const token = str_(p.getProperty('LINE_TOKEN')), to = str_(p.getProperty('LINE_TO'));
    if (!token || !to) throw new Error('LINE_TOKEN / LINE_TO が未設定です');
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text }] }),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() >= 300) throw new Error('LINE ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 120));
  } else if (kind === 'mail') {
    const to = str_(p.getProperty('NOTIFY_MAIL'));
    if (!to) throw new Error('NOTIFY_MAIL が未設定です');
    MailApp.sendEmail(to, '軽食アプリ', text);
  } else {
    throw new Error('NOTIFY_KIND が不明です：' + kind);
  }
}

function post_(url, obj) {
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(obj), muteHttpExceptions: true });
  if (res.getResponseCode() >= 300) throw new Error(res.getResponseCode() + ' ' + res.getContentText().slice(0, 120));
}

/** 賞味期限の予告をスタッフの外部通知にも流す。15分おきのトリガーから呼ぶ */
function notifyExpiry() { withLock_(notifyExpiry_); }   // 注文の書き込みと「通知待ち」の行がぶつからないように
function notifyExpiry_() {
  const warn = 20, sent = PropertiesService.getScriptProperties();
  const noted = JSON.parse(sent.getProperty('EXP_NOTED') || '{}');
  const now = new Date(), nowS = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
  const lim = Utilities.formatDate(new Date(now.getTime() + warn * 60000), TZ, 'yyyy-MM-dd HH:mm');
  let changed = false;
  tournaments_().filter(function (t) { return t.active; }).forEach(function (t) {
    const date = bizDate_();
    const st = state_(t, date);
    const byKey = {};
    st.items.forEach(function (it) {
      if (!it.exp || it.soldOut || it.remain != null) return;
      if (!(it.exp > nowS && it.exp <= lim)) return;
      const k = t.id + '|' + date + '|' + it.store + '|' + it.recv + '|' + it.exp;
      (byKey[k] = byKey[k] || []).push(popBase_(it.name));
    });
    Object.keys(byKey).forEach(function (k) {
      if (noted[k]) return;
      noted[k] = now.getTime(); changed = true;
      const ns = byKey[k];
      const label = ns.length === 1 ? ns[0] : ns.slice(0, 2).join('・') + ' ほか' + (ns.length - 2) + '品目';
      notifyPush_('期限', label + ' が期限切れになります。');
    });
  });
  if (changed) {
    const keep = {}, old = now.getTime() - 2 * 864e5;
    Object.keys(noted).forEach(function (k) { if (noted[k] > old) keep[k] = noted[k]; });
    sent.setProperty('EXP_NOTED', JSON.stringify(keep));
  }
}

/** メニューから1回押せば、通知のトリガーを作る */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (tg) {
    if (['sendPending', 'notifyExpiry'].indexOf(tg.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(tg);
  });
  ScriptApp.newTrigger('sendPending').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('notifyExpiry').timeBased().everyMinutes(15).create();
  try { SpreadsheetApp.getUi().alert('通知のトリガーを作りました（送信1分おき／期限の確認15分おき）'); } catch (e) {}
}

function popBase_(name) { return str_(name).replace(/(（[^）]*）)+$/, '').trim(); }

function popMaster_(withImg) {
  const m = {};
  let rows = [];
  try { rows = rows_(SH.POP); } catch (e) { return m; } // 初期設定前はPOP設定なしとして扱う
  rows.forEach(function (r) {
    const n = str_(r[0]);
    if (!n) return;
    m[n] = { price: r[1] === '' || r[1] == null ? '' : Number(r[1]), v: r[3] instanceof Date ? r[3].getTime() : 0, img: withImg ? str_(r[2]) : '' };
    if (!withImg) m[n].has = !!str_(r[2]);
  });
  return m;
}

/* ---------- 大会ごとの集計シート ---------- */

const TOURN_LAYOUT = '集計v3'; // レイアウトを変えたら値を変える（既存シートを作り直す）

/**
 * 大会ごとのシート「残数_大会名」を用意する（なければ作成、大会名が変わればシート名も変更）。
 * 中身は数式で「日次残数」「在庫設定」から引くため、期限切れ数が送信されるたびに自動で更新される。
 *   A〜G：その大会の日次残数の一覧（開催日 / 受取時間 / 品目（まとめ） / 品目 / 在庫数 / 期限切れ数 / 在庫切れ時刻）
 *   I〜N：品目ごとの集計（受取時間などの末尾の（ ）を外した品目名でまとめる）
 *   P〜V：品目×受取時間ごとの集計
 *   集計の列：記録回数 / 期限切れ回数（期限切れ数が1以上の回数）/ 期限切れ数合計 / 平均期限切れ数（記録1回あたり）/ 期限切れ率（期限切れ数÷在庫数）
 *   期限切れ率が50%を超える行は薄い赤。
 * シートは E1 の大会IDで識別し、F1 のレイアウト番号が古ければ作り直す。
 */
function tournSheet_(id, name) {
  const book = book_();
  if (!book.getSheets) return;
  const fixed = Object.keys(HEAD);
  let sh = book.getSheets().filter(function (x) {
    return fixed.indexOf(x.getName()) < 0 && str_(x.getRange('E1').getValue()) === id;
  })[0];
  const want = tournSheetName_(book, name, sh);
  if (!sh) sh = book.insertSheet(want);
  else if (sh.getName() !== want) sh.setName(want);
  if (str_(sh.getRange('F1').getValue()) !== TOURN_LAYOUT) buildTournSheet_(sh, id, name);
  if (str_(sh.getRange('B1').getValue()) !== name) sh.getRange('B1').setValue(name);
  return sh;
}

function buildTournSheet_(sh, id, name) {
  sh.clear();
  if (sh.setConditionalFormatRules) sh.setConditionalFormatRules([]);
  const c = "'" + SH.CLOSE + "'!", st = "'" + SH.STOCK + "'!";
  sh.getRange('A1:F1').setValues([['大会', name, '', '大会ID', id, TOURN_LAYOUT]]);
  sh.getRange('A1').setFontWeight('bold');
  sh.getRange('D1:F1').setFontColor('#9aa0a6');
  const stat = ['記録回数', '期限切れ回数', '期限切れ数合計', '平均期限切れ数', '期限切れ率'];
  sh.getRange('A3:G3').setValues([['開催日', '受取時間', '品目（まとめ）', '品目', '在庫数', '期限切れ数', '在庫切れ時刻']]);
  sh.getRange('I2').setValue('品目ごと');
  sh.getRange('I3:N3').setValues([['品目（まとめ）'].concat(stat)]);
  sh.getRange('P2').setValue('品目×受取時間ごと');
  sh.getRange('P3:V3').setValues([['品目（まとめ）', '受取時間'].concat(stat)]);
  ['A3:G3', 'I3:N3', 'P3:V3'].forEach(function (a) { sh.getRange(a).setFontWeight('bold').setBackground('#f1f3f4'); });
  sh.getRange('I2:P2').setFontWeight('bold');
  // 一覧：日次残数からこの大会の行を抽出し、受取時間は在庫設定から引く
  sh.getRange('A4').setFormula('=IFERROR(LET(' +
    'd, FILTER({' + c + 'C2:C, ' + c + 'D2:D, ' + c + 'E2:E, ' + c + 'F2:F, ' + c + 'H2:H}, ' + c + 'A2:A = $E$1), ' +
    'dt, INDEX(d, , 1), it, INDEX(d, , 2), ' +
    'rv, MAP(dt, it, LAMBDA(x, y, IFERROR(TO_TEXT(INDEX(FILTER(' + st + 'H2:H, ' + st + 'A2:A = $E$1, ' + st + 'C2:C = x, ' + st + 'E2:E = y), 1)), ""))), ' +
    'SORT(HSTACK(dt, rv, REGEXREPLACE(it, "(（[^）]*）)+$", ""), it, INDEX(d, , 3), INDEX(d, , 4), INDEX(d, , 5)), 1, TRUE, 2, TRUE, 4, TRUE)), "")');
  // 品目ごと（期限切れ回数の多い順）
  sh.getRange('I4').setFormula('=IFERROR(LET(n, UNIQUE(FILTER(C4:C, C4:C <> "")), ' +
    'SORT(HSTACK(n, ' +
    'MAP(n, LAMBDA(x, COUNTIF(C4:C, x))), ' +
    'MAP(n, LAMBDA(x, COUNTIFS(C4:C, x, F4:F, ">0"))), ' +
    'MAP(n, LAMBDA(x, SUMIF(C4:C, x, F4:F))), ' +
    'MAP(n, LAMBDA(x, AVERAGEIF(C4:C, x, F4:F))), ' +
    'MAP(n, LAMBDA(x, IFERROR(SUMIF(C4:C, x, F4:F) / SUMIF(C4:C, x, E4:E), 0)))), 3, FALSE, 4, FALSE)), "")');
  // 品目×受取時間ごと（品目名・受取時間の順）
  sh.getRange('P4').setFormula('=IFERROR(LET(k, UNIQUE(FILTER({C4:C, B4:B}, C4:C <> "")), n, INDEX(k, , 1), t, INDEX(k, , 2), ' +
    'SORT(HSTACK(n, t, ' +
    'MAP(n, t, LAMBDA(x, y, COUNTIFS(C4:C, x, B4:B, y))), ' +
    'MAP(n, t, LAMBDA(x, y, COUNTIFS(C4:C, x, B4:B, y, F4:F, ">0"))), ' +
    'MAP(n, t, LAMBDA(x, y, SUMIFS(F4:F, C4:C, x, B4:B, y))), ' +
    'MAP(n, t, LAMBDA(x, y, AVERAGEIFS(F4:F, C4:C, x, B4:B, y))), ' +
    'MAP(n, t, LAMBDA(x, y, IFERROR(SUMIFS(F4:F, C4:C, x, B4:B, y) / SUMIFS(E4:E, C4:C, x, B4:B, y), 0)))), 1, TRUE, 2, TRUE)), "")');
  ['E4:F', 'J4:L', 'R4:T'].forEach(function (a) { sh.getRange(a).setNumberFormat('0'); });
  sh.getRange('M4:M').setNumberFormat('0.0');
  sh.getRange('N4:N').setNumberFormat('0%');
  sh.getRange('U4:U').setNumberFormat('0.0');
  sh.getRange('V4:V').setNumberFormat('0%');
  sh.setFrozenRows(3);
  if (sh.setColumnWidth) {
    [[2, 70], [3, 160], [4, 200], [8, 24], [9, 160], [15, 24], [16, 160], [17, 70]].forEach(function (w) { sh.setColumnWidth(w[0], w[1]); });
  }
  if (sh.setConditionalFormatRules) {
    const rule = function (a, f, col) { return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(col).setRanges([sh.getRange(a)]).build(); };
    sh.setConditionalFormatRules([
      rule('F4:F', '=AND(ISNUMBER($E4),ISNUMBER($F4),$E4>0,$F4>=$E4/2)', '#f4cccc'), // 期限切れ数が在庫数の半分以上
      rule('F4:F', '=AND(ISNUMBER($F4),$F4=0)', '#cfe2f3'),                         // 期限切れ0
      rule('F4:F', '=ISNUMBER($F4)', '#d9ead3'),                                     // それ以外
      rule('I4:N', '=AND(ISNUMBER($N4),$N4>0.5)', '#f4cccc'),                        // 期限切れ率50%超
      rule('P4:V', '=AND(ISNUMBER($V4),$V4>0.5)', '#f4cccc'),
    ]);
  }
}

/** シート名に使えない文字を除き、他のシートと重ならない名前にする */
function tournSheetName_(book, name, self) {
  const base = ('残数_' + String(name).replace(/[\[\]\*\?\/\\:']/g, ' ').trim()).slice(0, 90);
  let n = base, k = 2;
  while (true) {
    const other = book.getSheetByName(n);
    if (!other || (self && other.getSheetId && self.getSheetId && other.getSheetId() === self.getSheetId())) return n;
    n = base + '（' + (k++) + '）';
  }
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
    // 注文受付は空欄なら「停止」とみなす（PIN不要の入口なので、明示したときだけ開ける）
    return { id: str_(r[0]), name: str_(r[1]), staffPin: str_(r[2]), active: str_(r[3]) !== '無効',
      orderOpen: str_(r[5]) === '受付中', orderHours: str_(r[6]) };
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
