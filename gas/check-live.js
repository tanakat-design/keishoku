#!/usr/bin/env node
/**
 * 本番の設定を外から点検する。
 *
 *   node gas/check-live.js                大会の一覧（IDつき）を出す
 *   node gas/check-live.js <大会ID>        その大会を点検する（今日の前後14日）
 *   node gas/check-live.js <大会ID> 30     前後30日を見る
 *
 * PINを使わない。getPopState を読むだけなので、何度流しても本番は変わらない。
 * 大会の設定漏れ（受取時間が空・賞味期限が受取より前・写真なし等）を機械的に拾う。
 * 要確認が1件でもあれば終了コード1で終わる。
 *
 * 非表示（無効）の大会は getPopState が返さないので、ここには出てこない。
 */
'use strict';
const fs = require('fs');
const path = require('path');

/** API_URL は index.html を正とする。ここに書き写すと、片方だけ直したときに食い違う */
function apiUrl() {
  const f = path.join(__dirname, '..', 'index.html');
  const m = fs.readFileSync(f, 'utf8').match(/const API_URL\s*=\s*'([^']+)'/);
  if (!m) throw new Error('index.html から API_URL を読み取れない');
  return m[1];
}
const URL_ = apiUrl();

async function call(fn, ...args) {
  const r = await fetch(URL_, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ fn, args }),
  });
  const j = await r.json();
  if (!j.ok) throw new Error(j.error);
  return j.result;
}

/** 同時に投げすぎるとGAS側で詰まるので、並列数を絞って回す */
async function mapLimit(list, n, f) {
  const out = new Array(list.length);
  let i = 0;
  await Promise.all(Array.from({ length: Math.min(n, list.length) }, async () => {
    while (i < list.length) { const k = i++; out[k] = await f(list[k]); }
  }));
  return out;
}

const ymd = (d) => d.toISOString().slice(0, 10);
const toMin = (hm) => { const m = /^(\d{1,2}):(\d{2})$/.exec(hm || ''); return m ? +m[1] * 60 + +m[2] : null; };
/** 'YYYY-MM-DD HH:MM' を分に直す。開催日からの経過分で見るので、日またぎも素直に比べられる */
function expMin(date, exp) {
  const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(exp || '');
  if (!m) return null;
  const d0 = Date.parse(date + 'T00:00:00Z');
  const d1 = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
  return Number.isNaN(d0) ? null : (d1 - d0) / 60000;
}

async function main() {
  const tid = process.argv[2];
  const span = Number(process.argv[3]) || 14;

  if (!tid) {
    const r = await call('getPopState');
    console.log('表示中の大会：');
    (r.tournaments || []).forEach((t) => console.log('  %s  %s', t.id, t.name));
    console.log('\n点検する：node gas/check-live.js <大会ID>');
    return 0;
  }

  // 日本時間の日付で並べる。toISOString は UTC なので、9時間足してからでないと朝9時前が1日ずれる
  const base = Date.now() + 9 * 3600000;
  const dates = Array.from({ length: span * 2 + 1 }, (_, k) => ymd(new Date(base + (k - span) * 86400000)));
  const got = await mapLimit(dates, 5, async (d) => {
    try { const r = await call('getPopState', tid, d); return r.items.length ? [d, r] : null; }
    catch (e) { return ['!' + d, e.message]; }
  });

  const bad = got.filter((x) => x && String(x[0])[0] === '!');
  if (bad.length) { console.error('取得に失敗した：' + bad[0][1]); return 1; }

  const days = got.filter(Boolean).sort((a, b) => (a[0] < b[0] ? -1 : 1));
  if (!days.length) {
    console.log('前後%d日に品目のある日が無い。大会IDか期間を確かめる。', span);
    return 1;
  }

  console.log('大会：%s', days[0][1].name);
  console.log('期間：%s 〜 %s（前後%d日を確認）\n', dates[0], dates[dates.length - 1], span);

  const warn = [];
  let noExpBlocks = 0, noImgNames = new Set(), noPrNames = new Set();
  for (const [date, st] of days) {
    const blocks = new Map();
    for (const it of st.items) {
      const key = [it.recv || '', it.store || '', it.exp || ''].join('|');
      if (!blocks.has(key)) blocks.set(key, { recv: it.recv || '', store: it.store || '', exp: it.exp || '', names: [] });
      blocks.get(key).names.push(it.base);
    }
    console.log('■ %s  %d品目', date, st.items.length);
    for (const b of [...blocks.values()].sort((x, y) => (x.recv + x.store < y.recv + y.store ? -1 : 1))) {
      console.log('    %s  %s  期限%s  %d品%s',
        (b.recv || '★受取なし').padEnd(6), (b.store || '★店舗なし').slice(0, 22).padEnd(22),
        (b.exp || '―').padEnd(17), b.names.length, b.exp ? '' : '  ← 賞味期限なし');

      if (!b.exp) noExpBlocks++;
      if (!b.recv) warn.push(`${date} ${b.store}：受取時間が空`);
      if (!b.store) warn.push(`${date} ${b.recv}：店舗が空`);
      if (new Set(b.names).size !== b.names.length) warn.push(`${date} ${b.store} ${b.recv}：同じブロックに同じ品目名がある`);
      if (b.exp) {
        const e = expMin(date, b.exp), r = toMin(b.recv);
        if (e === null) warn.push(`${date} ${b.store}：賞味期限の書き方が「YYYY-MM-DD HH:MM」でない（${b.exp}）`);
        else if (r !== null && e <= r) warn.push(`${date} ${b.store} ${b.recv}：賞味期限が受取時刻より前（${b.exp}）`);
        else if (e > 48 * 60) warn.push(`${date} ${b.store}：賞味期限が開催日の2日以上あと（${b.exp}）`);
      }
    }
    const names = [...new Set(st.items.map((i) => i.base))];
    const noImg = names.filter((n) => !(st.info[n] || {}).img);
    const noPr = names.filter((n) => { const p = (st.info[n] || {}).price; return p === '' || p == null; });
    noImg.forEach((n) => noImgNames.add(n));
    noPr.forEach((n) => noPrNames.add(n));
    if (noImg.length) console.log('    写真なし：%s', noImg.join('、'));
    if (noPr.length) console.log('    ドリチケ未設定：%s', noPr.join('、'));
    console.log('');
  }

  console.log('品目のある日：%s', days.map((d) => d[0]).join('、'));

  // 「賞味期限なし」「写真なし」は運用で承知のうえのことがあるので、要確認とは分けて数だけ出す
  const notes = [];
  if (noExpBlocks) notes.push(`賞味期限なしの受取ブロックが${noExpBlocks}件`);
  if (noImgNames.size) notes.push(`写真なしの品目が${noImgNames.size}件（${[...noImgNames].join('、')}）`);
  if (noPrNames.size) notes.push(`ドリチケ未設定の品目が${noPrNames.size}件（${[...noPrNames].join('、')}）`);
  if (notes.length) {
    console.log('\n【参考】意図したものなら直さなくてよい');
    notes.forEach((n) => console.log(' - %s', n));
  }

  if (warn.length) {
    console.log('\n【要確認】設定の間違いの可能性が高い');
    [...new Set(warn)].forEach((w) => console.log(' - %s', w));
    return 1;
  }
  console.log('\n要確認なし。');
  return 0;
}

// require されたときは走らせない。時刻の計算だけを取り出して試せるようにしてある
if (require.main === module) {
  main().then((c) => process.exit(c), (e) => { console.error('失敗：' + e.message); process.exit(1); });
}
module.exports = { ymd, toMin, expMin };
