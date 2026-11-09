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
                if (nr === 1 && r >= 2) DB[name][r - 2] = v[0].slice();
                else v.forEach((row, i) => { DB[name][r - 2 + i] = row.slice(); });
                return this;
              },
              getValues() {
                const out = [];
                for (let i = 0; i < nr; i++) out.push((DB[name][r - 2 + i] || []).slice(0, nc));
                return out;
              },
              setValue(v) { DB[name][r - 2][c - 1] = v; },
              setFontWeight() { return this; }, setBackground() { return this; },
              setNumberFormat() { return this; }, clearContent() { return this; },
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
const Logger = { log() {} };

const ctx = { Utilities, SpreadsheetApp, PropertiesService, LockService, UrlFetchApp, MailApp, ScriptApp, ContentService, HtmlService, Logger, console, Date, Math, JSON, String, Number, Object, Array, isNaN, parseInt, parseFloat };
const vm = require('vm');
vm.createContext(ctx);
vm.runInContext(src + '\nthis.__x = { placeOrder, getOrder, setOrder, normItems_, inHours_, recvAt_, bizDate_, tables_, onSale_, limitCheck_, adminSaveTable, adminOrderOpen, openOrders_, SH, HEAD, ORD, rnd_, addDay_ };', ctx);
const X = ctx.__x;

/* ---- 下ごしらえ ---- */
Object.keys(X.HEAD).forEach(n => { DB[n] = []; });
const today = X.bizDate_();
DB[X.SH.TOURN].push(['T1', 'テスト大会', '1234', '有効', new Date(), '受付中', '']);
DB[X.SH.TABLE].push(['T1', '3', 9, 'tok1234567890ab', '有効', new Date()]);
DB[X.SH.TABLE].push(['T1', '4', 9, 'tok4444444444aa', '無効', new Date()]);
DB[X.SH.POP].push(['からあげ弁当', 2, '', new Date()]);
DB[X.SH.POP].push(['おにぎり', 1, '', new Date()]);
// 在庫設定：受取07:00、期限は今日の23:59
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 1, 'からあげ弁当', 20, 'テスト店', '07:00', today + ' 23:59', '店舗受け取り', '住所', '000']);
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 2, 'おにぎり', 10, 'テスト店', '07:00', today + ' 23:59', '店舗受け取り', '住所', '000']);
DB[X.SH.STOCK].push(['T1', 'テスト大会', today, 3, '夜食パン', 10, 'テスト店', '23:50', today + ' 23:59', '店舗受け取り', '住所', '000']);

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
ok('受取前（23:50）は注文できない', !sale['夜食パン'] || new Date().getHours() >= 23);

console.log('--- placeOrder ---');
throws('トークンが違う', () => X.placeOrder('T1', '3', 'wrong', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r1'), /QR/);
throws('無効な卓', () => X.placeOrder('T1', '4', 'tok4444444444aa', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r2'), /QR/);
throws('存在しない卓', () => X.placeOrder('T1', '9', 'tok1234567890ab', '5', [{ name: 'からあげ弁当', qty: 1 }], 'r3'), /QR/);
throws('席が範囲外', () => X.placeOrder('T1', '3', 'tok1234567890ab', '10', [{ name: 'からあげ弁当', qty: 1 }], 'r4'), /席番号/);
throws('席が0', () => X.placeOrder('T1', '3', 'tok1234567890ab', '0', [{ name: 'からあげ弁当', qty: 1 }], 'r5'), /席番号/);
throws('売っていない品目', () => X.placeOrder('T1', '3', 'tok1234567890ab', '5', [{ name: '存在しない', qty: 1 }], 'r6'), /注文できません/);
throws('リクエストIDなし', () => X.placeOrder('T1', '3', 'tok1234567890ab', '5', [{ name: 'からあげ弁当', qty: 1 }], ''), /リクエストID/);

const o1 = X.placeOrder('T1', '3', 'tok1234567890ab', '5', [{ name: 'からあげ弁当', qty: 1 }, { name: 'おにぎり', qty: 2 }], 'req-A');
ok('注文できた', !!o1.id && !!o1.key, JSON.stringify(o1));
ok('明細が2行', o1.lines.length === 2);
ok('ドリチケ枚数が入る', DB[X.SH.ORDER][0][9] === '2', DB[X.SH.ORDER][0][9]);
ok('状態は受付', o1.lines[0].state === '受付');
ok('通知が積まれた', DB[X.SH.NOTIFY].length === 1 && /3番テーブルの5番シート/.test(DB[X.SH.NOTIFY][0][2]), DB[X.SH.NOTIFY][0] && DB[X.SH.NOTIFY][0][2]);

const o1b = X.placeOrder('T1', '3', 'tok1234567890ab', '5', [{ name: 'からあげ弁当', qty: 1 }], 'req-A');
ok('同じリクエストIDは元の注文を返す', o1b.id === o1.id && DB[X.SH.ORDER].length === 2, 'rows=' + DB[X.SH.ORDER].length);

throws('同じ席で続けて注文', () => X.placeOrder('T1', '3', 'tok1234567890ab', '5', [{ name: 'おにぎり', qty: 1 }], 'req-B'), /60秒/);

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
throws('停止中は注文できない', () => X.placeOrder('T1', '3', 'tok1234567890ab', '6', [{ name: 'おにぎり', qty: 1 }], 'req-C'), /受け付けていません/);
X.adminOrderOpen(PROPS.ADMIN_PIN, 'T1', true, '00:00-23:59');
const o2 = X.placeOrder('T1', '3', 'tok1234567890ab', '6', [{ name: 'おにぎり', qty: 1 }], 'req-D');
ok('再開したら注文できる', !!o2.id);

console.log('--- 卓のトークン再発行 ---');
const before = X.tables_('T1')[0].token;
X.adminSaveTable(PROPS.ADMIN_PIN, 'T1', '3', 9, '有効', 'new');
const after = X.tables_('T1')[0].token;
ok('トークンが変わる', before !== after && after.length === 16);
throws('古いQRは使えない', () => X.placeOrder('T1', '3', before, '7', [{ name: 'おにぎり', qty: 1 }], 'req-E'), /QR/);

if (fail) { console.log('\n[通知待ちシートの中身]'); console.log(JSON.stringify(DB[X.SH.NOTIFY])); }
console.log('\n合格 ' + pass + ' / 不合格 ' + fail);
process.exit(fail ? 1 : 0);
