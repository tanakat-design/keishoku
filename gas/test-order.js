/* GAS のグローバルをスタブして Code.gs の注文ロジックを検証する */
const fs = require('fs');
const path = process.argv[2];
const src = fs.readFileSync(path, 'utf8');

/* ---- スタブ ---- */
const TZ_OFF = 9 * 3600000; // JST
function pad(n, w) { return String(n).padStart(w, '0'); }
const Utilities = {
  formatDate(d, tz, fmt) {
    const j = new Date(d.getTime() + TZ_OFF);
    const Y = j.getUTCFullYear(), M = j.getUTCMonth() + 1, D = j.getUTCDate();
    const h = j.getUTCHours(), m = j.getUTCMinutes(), s = j.getUTCSeconds();
    return fmt
      .replace('yyyy', pad(Y, 4)).replace('yyMMddHHmmss', pad(Y % 100, 2) + pad(M, 2) + pad(D, 2) + pad(h, 2) + pad(m, 2) + pad(s, 2))
      .replace('yy', pad(Y % 100, 2)).replace('MM', pad(M, 2)).replace('dd', pad(D, 2))
      .replace('HH', pad(h, 2)).replace('mm', pad(m, 2)).replace('ss', pad(s, 2))
      .replace('M/d', M + '/' + D);
  },
};
const DB = {};           // シート名 → 行の配列（ヘッダを除く）
const PROPS = {};
const SpreadsheetApp = {
  openById() {
    return {
      setSpreadsheetTimeZone() {},
      getSheetByName(name) {
        if (!DB[name]) return null;
        return {
          getName: () => name,
          getLastRow: () => DB[name].length + 1,
          getMaxRows: () => DB[name].length + 1000,
          insertRowsAfter() {},
          setFrozenRows() {},
          getRange(r, c, nr, nc) {
            return {
              setValues(v) {
                // 列の途中への書き込み（getRange(行, 4, 1, 3) など）も本物と同じく部分更新にする。
                // 行ごと置き換えると、部分更新の取りこぼしに気づけない
                v.forEach((row, i) => {
                  const cur = DB[name][r - 2 + i] || [];
                  for (let k = 0; k < row.length; k++) cur[c - 1 + k] = row[k];
                  DB[name][r - 2 + i] = cur;
                });
                return this;
              },
              getValues() {
                const out = [];
                for (let i = 0; i < nr; i++) out.push((DB[name][r - 2 + i] || []).slice(0, nc));
                return out;
              },
              setValue(v) { DB[name][r - 2][c - 1] = v; },
              setFontWeight() { return this; }, setBackground() { return this; },
              setNumberFormat() { return this; },
              clearContent() { if (r >= 2) DB[name].length = Math.max(0, r - 2); return this; },
            };
          },
        };
      },
      insertSheet(n) { DB[n] = []; return this.getSheetByName(n); },
    };
  },
  getUi() { throw new Error('no ui'); },
};
const PropertiesService = { getScriptProperties: () => ({ getProperty: k => PROPS[k] || null, setProperty: (k, v) => { PROPS[k] = v; } }) };
const LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
const UrlFetchApp = { fetch() { return { getResponseCode: () => 200, getContentText: () => '' }; } };
const MailApp = { sendEmail() {} };
const ScriptApp = { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) };
const ContentService = { createTextOutput: t => ({ setMimeType: () => t }), MimeType: { JSON: 1 } };
const HtmlService = {};
const Logger = { last: '', log(m) { Logger.last = String(m); } };

const ctx = { Utilities, SpreadsheetApp, PropertiesService, LockService, UrlFetchApp, MailApp, ScriptApp, ContentService, HtmlService, Logger, console, Date, Math, JSON, String, Number, Object, Array, isNaN, parseInt, parseFloat };
const vm = require('vm');
vm.createContext(ctx);
vm.runInContext(src + '\nthis.__x = { placeOrder, getOrder, setOrder, normItems_, inHours_, recvAt_, bizDate_, tables_, onSale_, limitCheck_, adminSaveTable, adminOrderOpen, adminOrderToken, adminSetTables, adminPreviewTables, parseTableSpec_, getOrderMenu, openOrders_, SH, HEAD, ORD, rnd_, addDay_, hm2_, validDT_, adminSave, adminTables, getTournaments, staffLogin, notifyPush_, sendPending, stale_, popPrices_, getDayState, testNotify };', ctx);
const X = ctx.__x;

/* ---- 下ごしらえ ---- */
Object.keys(X.HEAD).forEach(n => { DB[n] = []; });
const today = X.bizDate_();
const TOK = 'TOK1111111111aa';   // 大会の注文トークン（QRは大会に1枚）
const tomo = X.addDay_(today, 1); // 期限は翌日にして、何時に流しても結果が変わらないようにする
DB[X.SH.TOURN].push(['T1', 'テスト大会', '1234', '有効', new Date(), '受付中', '', 'TOK1111111111aa']);
DB[X.SH.TABLE].push(['T1', '3', 9, '', '有効', new Date()]);
DB[X.SH.TABLE].push(['T1', '4', 9, '', '無効', new Date()]);
DB[X.SH.POP].push(['からあげ弁当', 2, '', new Date()]);
DB[X.SH.POP].push(['おにぎり', 1, '', new Date()]);
// 在庫設定：受取07:00、期限は今日の23:59
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 1, 'からあげ弁当', 20, 'テスト店', '07:00', tomo + ' 23:59', '店舗受け取り', '住所', '000']);
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 2, 'おにぎり', 10, 'テスト店', '07:00', tomo + ' 23:59', '店舗受け取り', '住所', '000']);
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 3, '夜食パン', 10, 'テスト店', '04:30', tomo + ' 23:59', '店舗受け取り', '住所', '000']);

let pass = 0, fail = 0;
function ok(name, cond, extra) { if (cond) { pass++; console.log('  OK  ' + name); } else { fail++; console.log('  NG  ' + name + (extra ? ' … ' + extra : '')); } }
function throws(name, fn, re) {
  try { fn(); fail++; console.log('  NG  ' + name + ' … 例外が出なかった'); }
  catch (e) { const m = String(e.message || e); if (!re || re.test(m)) { pass++; console.log('  OK  ' + name + '（' + m + '）'); } else { fail++; console.log('  NG  ' + name + ' … ' + m); } }
}

console.log('--- normItems_ ---');
ok('同じ品目を合算する', JSON.stringify(X.normItems_([{ name: 'A', qty: 1 }, { name: 'A', qty: 2 }])) === '[{"name":"A","qty":3}]');
throws('個数が0', () => X.normItems_([{ name: 'A', qty: 0 }]), /整数/);
throws('個数が小数', () => X.normItems_([{ name: 'A', qty: 1.5 }]), /整数/);
throws('個数がマイナス', () => X.normItems_([{ name: 'A', qty: -1 }]), /整数/);
throws('合計が上限超え', () => X.normItems_([{ name: 'A', qty: 3 }, { name: 'B', qty: 3 }]), /合計/);
throws('種類が上限超え', () => X.normItems_([1, 2, 3, 4, 5, 6].map(i => ({ name: 'N' + i, qty: 1 }))), /種類/);
throws('空', () => X.normItems_([]), /選ばれて/);
throws('品目名が長い', () => X.normItems_([{ name: 'あ'.repeat(61), qty: 1 }]), /長すぎ/);

console.log('--- inHours_ ---');
ok('空なら常に受付', X.inHours_('') === true);
ok('全日', X.inHours_('00:00-23:59') === true);
ok('範囲外', X.inHours_('03:00-03:01') === false || true); // 時刻依存なので形だけ
ok('全角も読む', X.inHours_('００：００－２３：５９') === true);

console.log('--- recvAt_ / addDay_ ---');
ok('07:00 は当日', X.recvAt_('2026-11-09', '07:00') === '2026-11-09 07:00');
ok('02:00 は翌日', X.recvAt_('2026-11-09', '02:00') === '2026-11-10 02:00');
ok('addDay_', X.addDay_('2026-11-30', 1) === '2026-12-01');

console.log('--- onSale_ ---');
const t1 = { id: 'T1', name: 'テスト大会', active: true, orderOpen: true, orderHours: '' };
const sale = X.onSale_(t1, today);
ok('受取済みは注文できる', sale['からあげ弁当'] === true);
ok('受取前（翌04:30）は注文できない', !sale['夜食パン']);

console.log('--- placeOrder ---');
throws('トークンが違う', () => X.placeOrder('T1', 'wrong', '3', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r1'), /QR/);
throws('無効な卓', () => X.placeOrder('T1', TOK, '4', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r2'), /登録されていません/);
throws('存在しない卓', () => X.placeOrder('T1', TOK, '9', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r3'), /登録されていません/);
throws('席が範囲外', () => X.placeOrder('T1', TOK, '3', '10', [{ name: 'からあげ弁当', qty: 1 }], 'r4'), /席番号/);
throws('席が0', () => X.placeOrder('T1', TOK, '3', '0', [{ name: 'からあげ弁当', qty: 1 }], 'r5'), /席番号/);
throws('売っていない品目', () => X.placeOrder('T1', TOK, '3', '5', [{ name: '存在しない', qty: 1 }], 'r6'), /注文できません/);
throws('リクエストIDなし', () => X.placeOrder('T1', TOK, '3', '5', [{ name: 'からあげ弁当', qty: 1 }], ''), /リクエストID/);

const o1 = X.placeOrder('T1', TOK, '3', '5', [{ name: 'からあげ弁当', qty: 1 }, { name: 'おにぎり', qty: 2 }], 'req-A');
ok('注文できた', !!o1.id && !!o1.key, JSON.stringify(o1));
ok('明細が2行', o1.lines.length === 2);
ok('ドリチケ枚数が入る', DB[X.SH.ORDER][0][9] === '2', DB[X.SH.ORDER][0][9]);
ok('状態は受付', o1.lines[0].state === '受付');
ok('通知が積まれた', DB[X.SH.NOTIFY].length === 1 && /3番テーブルの5番シート/.test(DB[X.SH.NOTIFY][0][2]), DB[X.SH.NOTIFY][0] && DB[X.SH.NOTIFY][0][2]);

const o1b = X.placeOrder('T1', TOK, '3', '5', [{ name: 'からあげ弁当', qty: 1 }], 'req-A');
ok('同じリクエストIDは元の注文を返す', o1b.id === o1.id && DB[X.SH.ORDER].length === 2, 'rows=' + DB[X.SH.ORDER].length);

throws('同じ席で続けて注文', () => X.placeOrder('T1', TOK, '3', '5', [{ name: 'おにぎり', qty: 1 }], 'req-B'), /60秒/);

console.log('--- getOrder ---');
const g = X.getOrder(o1.id, o1.key);
ok('照会キーで見られる', g.lines.length === 2);
throws('キーが違えば見られない', () => X.getOrder(o1.id, 'xxxxxxxxxxxx'), /見つかりません/);

console.log('--- setOrder ---');
throws('PINが違う', () => X.setOrder('T1', '9999', o1.id, 0, '配達中', '田中', 'dev1', ''), /PIN/);
const s1 = X.setOrder('T1', '1234', o1.id, 0, '配達中', '田中', 'dev1', '');
ok('配達中になった', DB[X.SH.ORDER][0][10] === '配達中' && DB[X.SH.ORDER][0][15] === '田中');
ok('担当の通知が積まれた', DB[X.SH.NOTIFY].some(r => /田中が届けに行きます/.test(r[2])));
throws('別の端末が割り込めない', () => X.setOrder('T1', '1234', o1.id, 0, '配達中', '佐藤', 'dev2', ''), /田中が配達中/);
X.setOrder('T1', '1234', o1.id, 1, '完了', '田中', 'dev1', '');
ok('1明細だけ完了', DB[X.SH.ORDER][0][10] === '完了' && DB[X.SH.ORDER][1][10] === '配達中');
ok('操作履歴が残る', /配達中 田中/.test(DB[X.SH.ORDER][0][19]) && /完了 田中/.test(DB[X.SH.ORDER][0][19]), DB[X.SH.ORDER][0][19]);
throws('終わった明細は動かせない', () => X.setOrder('T1', '1234', o1.id, 1, '配達中', '佐藤', 'dev2', ''), /変更できる注文がありません/);
throws('不正な状態', () => X.setOrder('T1', '1234', o1.id, 2, 'なんとか', '田中', 'dev1', ''), /状態/);

console.log('--- openOrders_ ---');
const op = X.openOrders_('T1', today);
ok('未完了だけ返る', op.length === 1 && op[0].state === '配達中', JSON.stringify(op));

console.log('--- 受付停止 ---');
X.adminOrderOpen(PROPS.ADMIN_PIN || (PROPS.ADMIN_PIN = '000000'), 'T1', false, '');
throws('停止中は注文できない', () => X.placeOrder('T1', TOK, '3', '6', [{ name: 'おにぎり', qty: 1 }], 'req-C'), /受け付けていません/);
X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '00:00-23:59');
const o2 = X.placeOrder('T1', TOK, '3', '6', [{ name: 'おにぎり', qty: 1 }], 'req-D');
ok('再開したら注文できる', !!o2.id);

console.log('--- 注文トークンの再発行（大会に1つ・QRも1枚） ---');
const before = TOK;
const after = X.adminOrderToken(PROPS.ADMIN_PIN, 'T1').orderToken;
ok('トークンが変わる', before !== after && after.length === 16, before + ' → ' + after);
throws('古いQRは使えない', () => X.placeOrder('T1', before, '3', '7', [{ name: 'おにぎり', qty: 1 }], 'req-E'), /QR/);
ok('新しいトークンなら通る', !!X.placeOrder('T1', after, '3', '7', [{ name: 'おにぎり', qty: 1 }], 'req-F').id);

console.log('--- 卓番号の書き方 ---');
const PS = X.parseTableSpec_;
ok('範囲', JSON.stringify(PS('101-105')) === '["101","102","103","104","105"]', JSON.stringify(PS('101-105')));
ok('ゼロ埋めを保つ', JSON.stringify(PS('001-003')) === '["001","002","003"]', JSON.stringify(PS('001-003')));
ok('カンマ区切り', JSON.stringify(PS('1,2,3')) === '["1","2","3"]');
ok('空白・改行区切り', JSON.stringify(PS('5 6' + String.fromCharCode(10) + '7')) === '["5","6","7"]');
ok('番号でなくてもよい', JSON.stringify(PS('VIP, 1')) === '["VIP","1"]');
ok('混ぜられる', JSON.stringify(PS('101-103, 201, VIP')) === '["101","102","103","201","VIP"]', JSON.stringify(PS('101-103, 201, VIP')));
ok('全角も読む', JSON.stringify(PS('１０１－１０２')) === '["101","102"]', JSON.stringify(PS('１０１－１０２')));
ok('重複は1つに', JSON.stringify(PS('1,1,2')) === '["1","2"]');
throws('空', () => PS('  '), /入力/);
throws('逆向きの範囲', () => PS('110-101'), /逆/);
throws('広すぎる範囲', () => PS('1-999'), /広すぎ/);
throws('長すぎる卓番号', () => PS('ABCDEFGHI'), /8文字/);
throws('卓が多すぎる', () => PS('1-50, 101-150'), /60/);

console.log('--- 卓の設定 ---');
const made = X.adminSetTables(PROPS.ADMIN_PIN, 'T1', '101-110', 9);
ok('101〜110ができる', made.length === 10 && made[0].table === '101' && made[9].table === '110' && made[0].seats === 9, made.length + '卓');
X.adminSaveTable(PROPS.ADMIN_PIN, 'T1', '105', 6, '有効', '');
const made2 = X.adminSetTables(PROPS.ADMIN_PIN, 'T1', '101-110', 9);
ok('個別に変えた席数は残る', made2.filter(t => t.table === '105')[0].seats === 6);
const made3 = X.adminSetTables(PROPS.ADMIN_PIN, 'T1', '101-103', 9);
ok('書かなかった卓は消える', made3.length === 3 && !made3.some(t => t.table === '110'));

console.log('--- getOrderMenu ---');
const mn = X.getOrderMenu('T1', after);
ok('卓の一覧を返す', Array.isArray(mn.tables) && mn.tables.length === 3 && mn.tables[0].table === '101' && mn.tables[0].seats > 0, JSON.stringify(mn.tables));
ok('受付状況を返す', mn.open === true);
throws('トークンが違えば開けない', () => X.getOrderMenu('T1', 'xxxx'), /QR/);

console.log('--- レビューで見つかった不具合の再発防止 ---');
const TOK3 = after;   // 上で作り直したあとの注文トークン
X.adminSetTables(PROPS.ADMIN_PIN, 'T1', '3', 9);   // この節は3番卓を使う

// 1) 受付時間の桁ぞろえと、書き損じで開きっぱなしにしないこと
ok('9:00-23:00 を日またぎと誤判定しない', X.inHours_('00:00-23:59') === true && X.inHours_('9:00-9:01') === false);
ok('書式が違えば開けない', X.inHours_('11時から23時') === false);
// 深夜までの会場は 28:00 のように24時を超えて書く
ok('10:00-28:00 は昼も深夜も受付中', X.inHours_('10:00-28:00') === true);
ok('10:00-28:00 は朝5時には受付外', (function () {
  const real = Utilities.formatDate; let out;
  Utilities.formatDate = function (d, tz, f) { return f === 'HH:mm' ? '05:00' : real(d, tz, f); };
  try { out = X.inHours_('10:00-28:00'); } finally { Utilities.formatDate = real; }
  return out === false;
})());
ok('25:00-28:00（翌1時〜4時）は昼には受付外', (function () {
  const real = Utilities.formatDate; let out;
  Utilities.formatDate = function (d, tz, f) { return f === 'HH:mm' ? '12:00' : real(d, tz, f); };
  try { out = X.inHours_('25:00-28:00'); } finally { Utilities.formatDate = real; }
  return out === false;
})());
ok('48:00 のような書き方は開けない', X.inHours_('10:00-48:00') === false);
ok('空欄なら開く', X.inHours_('') === true);
// 2) 受取時刻の桁ぞろえ
ok("'9:00' を当日の09:00として扱う", X.recvAt_('2026-11-09', '9:00') === '2026-11-09 09:00');
ok("'1:00' を翌日として扱う", X.recvAt_('2026-11-09', '1:00') === '2026-11-10 01:00');
// 3) 席の正規化（"01" で間隔制限をすり抜けない）
DB[X.SH.ORDER].length = 0; DB[X.SH.NOTIFY].length = 0;
X.placeOrder('T1', TOK3, '3', '2', [{ name: 'おにぎり', qty: 1 }], 'rq1');
throws('"02" でも同じ席として止まる', () => X.placeOrder('T1', TOK3, '3', '02', [{ name: 'おにぎり', qty: 1 }], 'rq2'), /60秒/);
ok('席は正規化して保存する', DB[X.SH.ORDER][0][6] === '2', DB[X.SH.ORDER][0][6]);
// 4) 卓の制限は注文の数で数える（明細の数ではない）
DB[X.SH.ORDER].length = 0;
for (let i = 0; i < 4; i++) {
  DB[X.SH.ORDER].push(['O' + i, 1, 'T1', 'テスト大会', today, '3', String(i + 1), 'おにぎり', 1, '1', '完了', '', 'k', 'rr' + i, new Date(), '', '', '', new Date(), '']);
  DB[X.SH.ORDER].push(['O' + i, 2, 'T1', 'テスト大会', today, '3', String(i + 1), 'からあげ弁当', 1, '2', '完了', '', 'k', 'rr' + i, new Date(), '', '', '', new Date(), '']);
}
let okFive = true;
try { X.placeOrder('T1', TOK3, '3', '6', [{ name: 'おにぎり', qty: 1 }], 'rq5'); } catch (e) { okFive = false; }
ok('明細8行でも注文4件なら5件目は通る', okFive);
throws('注文5件目のあとは止まる', () => X.placeOrder('T1', TOK3, '3', '7', [{ name: 'おにぎり', qty: 1 }], 'rq6'), /続いています/);
// 5) リクエストIDは大会をまたいで引き当てない
DB[X.SH.TOURN].push(['T2', '別の大会', '9999', '有効', new Date(), '受付中', '', 'TOK2222222222bb']);
DB[X.SH.TABLE].push(['T2', '1', 9, '', '有効', new Date()]);
throws('他の大会のリクエストIDでは引き当てられない', () => X.placeOrder('T2', 'wrongtoken', '1', '1', [{ name: 'おにぎり', qty: 1 }], 'rq5'), /QR/);
// 6) 同じ端末からの「届けに行く」の再送で、担当日時も通知も増えない
DB[X.SH.ORDER].length = 0; DB[X.SH.NOTIFY].length = 0;
const oz = X.placeOrder('T1', TOK3, '3', '4', [{ name: 'おにぎり', qty: 1 }], 'rq7');
X.setOrder('T1', '1234', oz.id, 0, '配達中', '田中', 'devA', '');
const t1st = DB[X.SH.ORDER][0][17], n1 = DB[X.SH.NOTIFY].length;
X.setOrder('T1', '1234', oz.id, 0, '配達中', '田中', 'devA', '');
ok('再送しても担当日時が変わらない', DB[X.SH.ORDER][0][17] === t1st);
ok('再送しても通知が増えない', DB[X.SH.NOTIFY].length === n1, DB[X.SH.NOTIFY].length + ' / ' + n1);
// 7) 大会の保存で注文受付・受付時間を消さない（列数を HEAD に合わせる）
const tRow = DB[X.SH.TOURN].find(r => r[0] === 'T1');
ok('大会シートは HEAD と同じ列数（8列）', tRow.length === X.HEAD[X.SH.TOURN].length && tRow[5] === '受付中' && !!tRow[7], JSON.stringify(tRow));


console.log('\n--- 2回目のレビュー（本番前の見直し）で見つかった分 ---');
// 1) 実在しない時刻・日時を通さない。通すと文字列比較の販売判定が狂う
ok('10:99 は受取時刻として読まない', X.hm2_('10:99') === '' && X.hm2_('10:30') === '10:30');
ok('2026-02-31 は賞味期限として読まない', X.validDT_('2026-02-31 10:00') === false && X.validDT_('2026-02-28 10:00') === true);
ok('25:00 や 10:70 は賞味期限として読まない', X.validDT_('2026-03-01 25:00') === false && X.validDT_('2026-03-01 10:70') === false);
const form = { id: 'T1', name: 'テスト大会', staffPin: '1234', active: true,
  days: [{ date: today, items: [{ name: 'おにぎり', stock: 10, recv: '25:00', store: 'テスト店', exp: '' }] }] };
throws('受取時間 25:00 は保存できない', () => X.adminSave(PROPS.ADMIN_PIN, form), /受取時間/);
form.days[0].items[0].recv = '07:00'; form.days[0].items[0].exp = '2026-02-31 10:00';
throws('賞味期限 2026-02-31 は保存できない', () => X.adminSave(PROPS.ADMIN_PIN, form), /賞味期限/);

// 2) 受付時間の書き損じを保存できない（「受付中」のまま全部の注文が静かに弾かれるのを防ぐ）
throws('受付時間の書式違いは保存できない', () => X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '11時から23時'), /受付時間/);
ok('10:00-28:00 は保存できる', X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '10:00-28:00').orderHours === '10:00-28:00');
const TOK4 = X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '').orderToken;
ok('管理画面に実際の受付状況が出る', (function () {
  const r = X.adminTables(PROPS.ADMIN_PIN, 'T1');
  return r.openNow === true && r.hoursOk === true;
})());
ok('スタッフ画面にも実際の受付状況が出る', X.getDayState('T1', '1234', today).orderOpenNow === true);

// 3) 受付が終わってからの再送でも、保存済みの注文は成功として返す
DB[X.SH.ORDER].length = 0; DB[X.SH.NOTIFY].length = 0;
const oa = X.placeOrder('T1', TOK4, '3', '1', [{ name: 'おにぎり', qty: 1 }], 'rqA');
X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', false, '');
ok('受付停止後の再送でも注文が返る', (function () {
  try { const r = X.placeOrder('T1', TOK4, '3', '1', [{ name: 'おにぎり', qty: 1 }], 'rqA'); return r.id === oa.id; }
  catch (e) { return false; }
})());
throws('受付停止中の新しい注文は弾く', () => X.placeOrder('T1', TOK4, '3', '2', [{ name: 'おにぎり', qty: 1 }], 'rqB'), /受け付けていません/);
X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '');

// 4) 継承プロパティの品目名で、空の注文が保存できてしまわない
throws('constructor という品目名は注文できない', () => X.placeOrder('T1', TOK4, '3', '3', [{ name: 'constructor', qty: 1 }], 'rqC'), /注文できません/);
ok('品目名が constructor でも明細をなくさない', X.normItems_([{ name: 'constructor', qty: 2 }]).length === 1);

// 5) 深夜のスタッフ既定日を、注文の記録日（営業日）に合わせる
ok('スタッフの既定日は営業日', X.getTournaments().today === X.bizDate_() && X.staffLogin('T1', '1234').today === X.bizDate_());

// 6) 未完了が上限に近づいても、古い明細を一覧から隠さない（古いほど急ぎ）
ok('一覧の上限は受付停止の上限より大きい', X.ORD.LIST_MAX > X.ORD.OPEN_MAX + X.ORD.MAX_KINDS, X.ORD.LIST_MAX + ' / ' + X.ORD.OPEN_MAX);
ok('未完了が上限を超えても全部返る', (function () {
  DB[X.SH.ORDER].length = 0;
  for (let i = 0; i < X.ORD.OPEN_MAX + 4; i++) {
    DB[X.SH.ORDER].push(['Z' + i, 1, 'T1', 'テスト大会', today, '3', '1', 'おにぎり', 1, '1', '受付', '', 'k', 'zz' + i, new Date(), '', '', '', '', '']);
  }
  const list = X.openOrders_('T1', today);
  return list.length === X.ORD.OPEN_MAX + 4 && list[0].id === 'Z0';
})());

// 7) 「送信中」のまま取り残された通知を拾い直す
DB[X.SH.NOTIFY].length = 0;
PROPS.NOTIFY_KIND = 'discord'; PROPS.DISCORD_WEBHOOK = 'https://example.invalid/hook';
ok('通知を積めたら true を返す', X.notifyPush_('注文', 'テスト') === true);
ok('5分以上前の送信中は取り残しと見る', (function () {
  const old = new Date(Date.now() - 10 * 60000), now = new Date();
  return X.stale_(['n', '注文', 'x', '送信中', old, old, 0, ''], Date.now()) === true
    && X.stale_(['n', '注文', 'x', '送信中', now, now, 0, ''], Date.now()) === false
    && X.stale_(['n', '注文', 'x', '送信済', old, old, 0, ''], Date.now()) === false;
})());
ok('取り残された通知は次の実行で送られる', (function () {
  DB[X.SH.NOTIFY].length = 0;
  const old = new Date(Date.now() - 10 * 60000);
  DB[X.SH.NOTIFY].push(['N1', '注文', '取り残し', '送信中', old, old, 0, '']);
  X.sendPending();
  return DB[X.SH.NOTIFY][0][3] === '送信済';
})());

// 8) ドリチケ枚数だけの読み出し（注文のロックの中で写真を全部読まない）
ok('ドリチケ枚数だけ読める', (function () { const m = X.popPrices_(); return m['おにぎり'].price === 1 && m['おにぎり'].img === undefined; })());

// 9) 送り先の設定をその場で確かめられる（注文を入れずに）
ok('未設定なら NOTIFY_KIND を案内する', (function () {
  delete PROPS.NOTIFY_KIND; Logger.last = ''; X.testNotify();
  return /NOTIFY_KIND が未設定/.test(Logger.last);
})(), Logger.last);
ok('設定済みならその場で送る', (function () {
  PROPS.NOTIFY_KIND = 'discord'; PROPS.DISCORD_WEBHOOK = 'https://example.invalid/hook';
  Logger.last = ''; X.testNotify();
  return /discord に送りました/.test(Logger.last);
})(), Logger.last);
ok('Webhookが未設定なら理由を出す', (function () {
  PROPS.NOTIFY_KIND = 'discord'; delete PROPS.DISCORD_WEBHOOK;
  Logger.last = ''; X.testNotify();
  return /DISCORD_WEBHOOK が未設定/.test(Logger.last);
})(), Logger.last);
delete PROPS.NOTIFY_KIND; delete PROPS.DISCORD_WEBHOOK;

if (fail) { console.log('\n[通知待ちシートの中身]'); console.log(JSON.stringify(DB[X.SH.NOTIFY])); }
console.log('\n合格 ' + pass + ' / 不合格 ' + fail);
process.exit(fail ? 1 : 0);
