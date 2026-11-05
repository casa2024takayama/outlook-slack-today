// 「予定をSlackへ」の画面テスト。1シナリオ＝1回のページ読み込み（状態はシナリオごとにまっさら）。
// setup() はアドイン本体より先に動き、開いている予定などを決める。run(t) は初期化のあとに動く。
(function () {
  'use strict';
  const URL_OK = 'https://hooks.slack.com/triggers/T000TEST/111/abcDEF';
  const pad = (n) => String(n).padStart(2, '0');
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const now = new Date();
  const at = (dayOffset, h, m = 0) => new Date(now.getFullYear(), now.getMonth(), now.getDate() + dayOffset, h, m);
  const today = ymd(now);
  const tomorrow = ymd(at(1, 0));
  const I = window.MockItems;
  const open = (item) => { Office.context.mailbox.item = item; };
  const withUrl = () => Office.context.roamingSettings.set('slackWebhookUrl', URL_OK);
  const MEETING = { id: 'AQMk+test/ID==', subject: '定例', start: at(0, 17), end: at(0, 17, 45), location: '3.Sirius（シリウス）_定員6名_本社' };
  // 送った本文の2行目（JSON）を取り出す
  const sentJson = (t, i = 0) => JSON.parse(JSON.parse(t.calls.fetch[i].opts.body).message.split('\n')[1]);

  window.SCENARIOS = {
    'no-url': {
      title: 'URL 未設定なら送らず、設定欄を開いて案内する',
      setup() { open(I.readAppt(MEETING)); },
      async run(t) {
        t.ok(t.$('#settings').open, '設定欄が最初から開いていない');
        t.click('#send');
        await t.sleep(50);
        t.eq(t.calls.fetch.length, 0, 'URL がないのに送信した');
        t.has(t.toast(), 'Webhook URL', '案内が出ない');
      },
    },
    'send-upsert': {
      title: '表示画面の予定を送る：送信先・形式・JSON・送信済み表示',
      setup() { withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.eq(t.$('#ev .t').textContent, '定例', '件名の表示');
        t.has(t.$('#ev .loc').textContent, 'シリウス', '場所の表示');
        t.ok(!t.visible(t.$('#notToday')), '今日の予定なのに「今日ではない」と出る');
        t.ok(!t.$('#settings').open, 'URL 設定済みなのに設定欄が開いている');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        const f = t.calls.fetch[0];
        t.eq(f.url, URL_OK, '送信先 URL');
        t.eq(f.opts.method, 'POST', 'メソッド');
        t.eq(f.opts.mode, 'no-cors', 'mode');
        const body = JSON.parse(f.opts.body);
        t.eq(Object.keys(body), ['message'], 'ワークフロー変数は message だけ');
        t.has(body.message.split('\n')[0], '📅 予定：定例（', '1行目（人が読む文）');
        t.eq(sentJson(t), { cw: 1, op: 'upsert', id: 'AQMk+test/ID==', title: '定例', start: `${today}T17:00:00+09:00`, end: `${today}T17:45:00+09:00`, allDay: false, location: '3.Sirius（シリウス）_定員6名_本社' }, 'JSON の中身');
        await t.until(() => t.visible(t.$('#sent')), '送信済み表示');
        t.has(t.$('#sent').textContent, '送信済み', '送信済みの文言');
        t.ok(JSON.parse(localStorage.getItem('olSlackToday.sent.v1'))['AQMk+test/ID=='], '送信記録が残っていない');
      },
    },
    'not-today': {
      title: '明日の予定は「今日ではない」注意を出す（送ることはできる）',
      setup() { withUrl(); open(I.readAppt({ ...MEETING, start: at(1, 10), end: at(1, 11) })); },
      async run(t) {
        t.ok(t.visible(t.$('#notToday')), '注意が出ない');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
      },
    },
    'allday-read': {
      title: '表示画面の終日（0:00〜翌0:00）は日付だけ・end は翌日',
      setup() { withUrl(); open(I.readAppt({ ...MEETING, subject: '休み', start: at(0, 0), end: at(1, 0), location: '' })); },
      async run(t) {
        t.has(t.$('#ev .when').textContent, '終日', '終日の表示');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        t.eq(sentJson(t), { cw: 1, op: 'upsert', id: 'AQMk+test/ID==', title: '休み', start: today, end: tomorrow, allDay: true }, '終日の JSON');
      },
    },
    'allday-multi': {
      title: '2日間の終日は end が3日目',
      setup() { withUrl(); open(I.readAppt({ ...MEETING, subject: '出張', start: at(0, 0), end: at(2, 0), location: '' })); },
      async run(t) {
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        const j = sentJson(t);
        t.eq([j.start, j.end, j.allDay], [today, ymd(at(2, 0)), true], '期間');
      },
    },
    'compose-saved': {
      title: '編集画面（主催）の予定：getAsync で読み、終日フラグを優先',
      setup() { withUrl(); open(I.composeAppt({ id: 'COMPOSE-1', subject: '社内打合せ', start: at(0, 0), end: at(1, 0), allDay: false, location: 'Teams' })); },
      async run(t) {
        await t.until(() => t.$('#ev .t'), '予定の表示');
        t.eq(t.$('#ev .t').textContent, '社内打合せ', '件名');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        const j = sentJson(t);
        t.eq([j.id, j.allDay, j.start], ['COMPOSE-1', false, `${today}T00:00:00+09:00`], 'Outlook の「終日ではない」を優先する');
      },
    },
    'compose-unsaved': {
      title: '未保存の予定は ID がないので、保存を案内して送らせない',
      setup() { withUrl(); open(I.composeAppt({ subject: '新しい予定', start: at(0, 9), end: at(0, 10) })); },
      async run(t) {
        await t.until(() => t.$('#ev').textContent.includes('保存'), '保存の案内');
        t.ok(!t.visible(t.$('#send')), '送信ボタンが見えている');
      },
    },
    'cancel-two-press': {
      title: '取り消しは2回押しで送る（1回目は送らない）',
      setup() { withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.click('#cancel');
        await t.sleep(50);
        t.eq(t.calls.fetch.length, 0, '1回目で送ってしまった');
        t.has(t.$('#cancel').textContent, 'もう一度押す', '確認の文言');
        t.click('#cancel');
        await t.until(() => t.calls.fetch.length === 1, '2回目で送信');
        t.eq(sentJson(t), { cw: 1, op: 'cancel', id: 'AQMk+test/ID==', title: '定例' }, 'cancel の JSON');
        t.has(JSON.parse(t.calls.fetch[0].opts.body).message, '🗑 予定の取り消し：定例', '1行目');
        await t.until(() => t.$('#sent').textContent.includes('取り消し送信済み'), '取り消し送信済み表示');
        t.eq(t.$('#cancel').textContent, '取り消しを送る', '文言が元に戻らない');
      },
    },
    'cancel-timeout': {
      title: '取り消しの1回目から5秒たつと元に戻る（うっかり2回目で送らない）',
      setup() { withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.click('#cancel');
        await t.sleep(5200);
        t.eq(t.$('#cancel').textContent, '取り消しを送る', '5秒後に戻らない');
        t.click('#cancel');
        await t.sleep(50);
        t.eq(t.calls.fetch.length, 0, '時間切れのあとの1回目で送ってしまった');
      },
    },
    'settings-save': {
      title: '設定：形の違う URL は保存しない・正しい URL は Outlook の設定に保存',
      setup() { open(I.readAppt(MEETING)); },
      async run(t) {
        t.$('#url').value = 'https://example.com/hook';
        t.click('#save');
        await t.sleep(30);
        t.eq(t.calls.roamingSaved, 0, '間違った URL を保存した');
        t.has(t.toast(), 'hooks.slack.com', '案内');
        t.$('#url').value = '  ' + URL_OK + '  ';
        t.click('#save');
        await t.until(() => t.calls.roamingSaved === 1, '保存');
        t.eq(Office.context.roamingSettings.get('slackWebhookUrl'), URL_OK, '前後の空白を除いて保存');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '保存した URL で送信');
      },
    },
    'test-send': {
      title: 'テスト送信は予定の JSON を含まない（ウィジェットに影響しない）',
      setup() { withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.click('#test');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        const m = JSON.parse(t.calls.fetch[0].opts.body).message;
        t.has(m, '🔧 テスト送信', '文言');
        t.ok(!m.includes('"cw"'), 'JSON が入っている');
        t.ok(!t.visible(t.$('#sent')), 'テストなのに送信済みになった');
      },
    },
    'item-changed': {
      title: 'ピン留め中に別の予定を選ぶと読み直す',
      setup() { withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.click('#send');
        await t.until(() => t.visible(t.$('#sent')), '1件目の送信済み');
        const h = t.calls.handlers['olkItemSelectedChanged'];
        t.ok(h, 'ItemChanged に登録していない');
        Office.context.mailbox.item = I.readAppt({ id: 'OTHER', subject: '面談', start: at(0, 15), end: at(0, 15, 30) });
        h();
        await t.until(() => t.$('#ev .t') && t.$('#ev .t').textContent === '面談', '2件目の表示');
        t.ok(!t.visible(t.$('#sent')), '別の予定なのに送信済みと出る');
      },
    },
    'fetch-fail': {
      title: '送信に失敗したら知らせ、送信済みにしない',
      setup() { window.__MOCK = { fetchFail: true }; withUrl(); open(I.readAppt(MEETING)); },
      async run(t) {
        t.click('#send');
        await t.until(() => t.toast().includes('送信できませんでした'), '失敗の案内');
        t.ok(!t.visible(t.$('#sent')), '失敗したのに送信済み');
        t.ok(!t.$('#send').disabled, 'ボタンが押せないまま');
      },
    },
    'no-location': {
      title: '場所が空なら location を送らない',
      setup() { withUrl(); open(I.readAppt({ ...MEETING, location: '' })); },
      async run(t) {
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        t.ok(!('location' in sentJson(t)), 'location が入っている');
      },
    },
    'special-chars': {
      title: '件名の記号（& " < > 絵文字）がそのまま JSON に入る',
      setup() { withUrl(); open(I.readAppt({ ...MEETING, subject: 'R&D "定例" <臨時> 🎉' })); },
      async run(t) {
        t.eq(t.$('#ev .t').textContent, 'R&D "定例" <臨時> 🎉', '画面の表示（HTML として解釈しない）');
        t.click('#send');
        await t.until(() => t.calls.fetch.length === 1, '送信');
        t.eq(sentJson(t).title, 'R&D "定例" <臨時> 🎉', 'JSON の件名');
      },
    },
    'already-sent': {
      title: '前に送った予定を開くと送信済みと出る',
      setup() {
        localStorage.setItem('olSlackToday.sent.v1', JSON.stringify({ 'AQMk+test/ID==': { op: 'upsert', at: new Date().toISOString() } }));
        withUrl(); open(I.readAppt(MEETING));
      },
      async run(t) { t.ok(t.visible(t.$('#sent')), '送信済みが出ない'); },
    },
    'no-item': {
      title: '予定を開いていないときは案内だけ',
      setup() { open(undefined); },
      async run(t) {
        t.has(t.$('#ev').textContent, '予定を開くと', '案内');
        t.ok(!t.visible(t.$('#send')), '送信ボタンが見えている');
      },
    },
    'outside-outlook': {
      title: 'Outlook の外で開くと試し画面になる',
      setup() { window.__MOCK = { host: null }; },
      async run(t) {
        await t.until(() => t.visible(t.$('#demo')), '試し画面');
        t.has(t.$('#ev .t').textContent, '定例ミーティング', 'サンプル予定');
      },
    },
  };

  // シナリオごとにまっさらにしてから setup
  const name = decodeURIComponent(location.hash.slice(1));
  try { localStorage.clear(); } catch (e) { /* なし */ }
  window.__MOCK = {};
  if (window.SCENARIOS[name]) window.SCENARIOS[name].setup();
})();
