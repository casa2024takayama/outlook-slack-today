// Outlook アドイン：開いている予定を、commute-widget が読む cw:1 形式で Slack ワークフロー（Webhook）に送る。
// 送るのはボタンを押したときだけ。Webhook URL はこの Outlook アカウントの設定（roamingSettings）に保存する。
(function () {
  'use strict';
  const VERSION = '0.1.0';
  const URL_KEY = 'slackWebhookUrl';
  const SENT_KEY = 'olSlackToday.sent.v1';
  const URL_RE = /^https:\/\/hooks\.slack\.com\/triggers\/[A-Za-z0-9/_-]+$/;
  const C = window.CwPayload;
  const $ = (id) => document.getElementById(id);
  const pad = (n) => String(n).padStart(2, '0');

  let inOutlook = false;
  let ev = null; // { id, title, start, end, allDay?, location }

  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 5000);
  }

  // ---------- 設定（Webhook URL） ----------
  function getUrl() {
    try { if (inOutlook) return Office.context.roamingSettings.get(URL_KEY) || ''; } catch (e) { /* 下へ */ }
    try { return localStorage.getItem(URL_KEY) || ''; } catch (e) { return ''; }
  }
  function setUrl(u) {
    return new Promise((resolve) => {
      if (inOutlook) {
        try {
          Office.context.roamingSettings.set(URL_KEY, u);
          Office.context.roamingSettings.saveAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded));
          return;
        } catch (e) { /* 下へ */ }
      }
      try { localStorage.setItem(URL_KEY, u); resolve(true); } catch (e) { resolve(false); }
    });
  }

  // ---------- 送信の記録（予定ごと・このPC） ----------
  const loadSent = () => { try { return JSON.parse(localStorage.getItem(SENT_KEY)) || {}; } catch (e) { return {}; } };
  function markSent(id, op) {
    try {
      const s = loadSent();
      s[id] = { op, at: new Date().toISOString() };
      // 古い記録は30日で消す
      const lim = Date.now() - 30 * 864e5;
      Object.keys(s).forEach((k) => { if (new Date(s[k].at) < lim) delete s[k]; });
      localStorage.setItem(SENT_KEY, JSON.stringify(s));
    } catch (e) { /* 記録できなくても送信は済んでいる */ }
  }

  async function post(message) {
    const url = getUrl();
    if (!URL_RE.test(url)) { $('settings').open = true; toast('先に設定で Webhook URL を保存してください'); return false; }
    try {
      // Slack は CORS ヘッダーを返さないので no-cors。応答は読めない（届いたかは Slack で確認）
      await fetch(url, { method: 'POST', mode: 'no-cors', body: JSON.stringify({ message }) });
      return true;
    } catch (e) {
      toast('送信できませんでした（ネットワーク・VPN・URL を確認）：' + e.message);
      return false;
    }
  }

  // ---------- 画面 ----------
  const fmt = (d) => `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  function render() {
    const box = $('ev');
    if (!ev) { box.innerHTML = '<p class="empty">予定を開くと、ここに表示します</p>'; $('actions').hidden = true; return; }
    const o = C.upsert(ev);
    box.innerHTML = '<div class="t"></div><div class="m when"></div><div class="m loc"></div>';
    box.querySelector('.t').textContent = ev.title || '（件名なし）';
    box.querySelector('.when').textContent = o.allDay ? `${o.start} 終日（〜${o.end} の前日まで）` : `${fmt(ev.start)} 〜 ${fmt(ev.end)}`;
    box.querySelector('.loc').textContent = ev.location ? `📍 ${ev.location}` : '';
    $('actions').hidden = false;
    $('notToday').hidden = C.touchesToday(o);
    $('preview').textContent = C.message(o);
    const s = loadSent()[ev.id];
    const sent = $('sent');
    sent.hidden = !s;
    if (s) {
      const d = new Date(s.at);
      sent.innerHTML = '<span class="badge"></span><span class="at"></span>';
      sent.querySelector('.badge').textContent = s.op === 'cancel' ? '取り消し送信済み' : '送信済み';
      sent.querySelector('.at').textContent = `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}（内容を変えたらもう一度送ってください）`;
    }
  }

  // ---------- 予定の読み取り ----------
  const getAsync = (prop) => new Promise((resolve, reject) => prop.getAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))));
  async function readItem() {
    const item = Office.context.mailbox.item;
    if (!item) { ev = null; return render(); }
    try {
      if (item.subject && typeof item.subject.getAsync === 'function') {
        // 編集画面（自分が主催者の予定）
        const [title, start, end, location] = await Promise.all([getAsync(item.subject), getAsync(item.start), getAsync(item.end), getAsync(item.location)]);
        let allDay;
        if (item.isAllDayEvent && item.isAllDayEvent.getAsync) { try { allDay = await getAsync(item.isAllDayEvent); } catch (e) { /* 推定に任せる */ } }
        const id = await new Promise((resolve) => {
          if (!item.getItemIdAsync) return resolve('');
          item.getItemIdAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ''));
        });
        if (!id) {
          ev = null;
          $('ev').innerHTML = '<p class="empty">まだ保存されていない予定です。<br>一度保存してから開き直して、もう一度押してください（予定 ID が付きます）。</p>';
          $('actions').hidden = true;
          return;
        }
        ev = { id, title, start, end, allDay, location };
      } else {
        // 表示画面（招待された予定・保存済みの予定）
        ev = { id: item.itemId, title: item.subject, start: new Date(item.start), end: new Date(item.end), location: item.location || '' };
      }
      render();
    } catch (e) {
      ev = null;
      $('ev').innerHTML = '<p class="empty"></p>';
      $('ev').firstChild.textContent = '予定を読み取れませんでした：' + (e && e.message ? e.message : e);
      $('actions').hidden = true;
    }
  }

  // ---------- ボタン ----------
  $('send').onclick = async () => {
    if (!ev) return;
    const o = C.upsert(ev);
    $('send').disabled = true;
    if (await post(C.message(o))) {
      markSent(ev.id, 'upsert');
      render();
      toast('送信しました。Slack の「Workflow Builder」とのDMに届いたか確認してください');
    }
    $('send').disabled = false;
  };
  // アドインの中では confirm() が使えないので、2回押しで確かめる
  let armTimer;
  $('cancel').onclick = async () => {
    if (!ev) return;
    const btn = $('cancel');
    if (!btn.dataset.armed) {
      btn.dataset.armed = '1';
      btn.textContent = 'もう一度押すと取り消しを送ります（Outlook の予定は消しません）';
      armTimer = setTimeout(() => { delete btn.dataset.armed; btn.textContent = '取り消しを送る'; }, 5000);
      return;
    }
    clearTimeout(armTimer);
    delete btn.dataset.armed;
    btn.textContent = '取り消しを送る';
    if (await post(C.message(C.cancel(ev)))) {
      markSent(ev.id, 'cancel');
      render();
      toast('取り消しを送信しました');
    }
  };
  $('save').onclick = async () => {
    const u = $('url').value.trim();
    if (u && !URL_RE.test(u)) return toast('https://hooks.slack.com/triggers/… の形の URL を貼ってください');
    toast((await setUrl(u)) ? (u ? '保存しました' : 'URL を消しました') : '保存できませんでした');
  };
  $('test').onclick = async () => {
    // 予定の JSON を含まないので、ウィジェットの表示には影響しない
    if (await post(`🔧 テスト送信（予定をSlackへ v${VERSION}）${new Date().toLocaleString()}`)) toast('テストを送信しました。Slack に届いたか確認してください');
  };

  function startDemo() {
    $('demo').hidden = false;
    const now = new Date();
    const loc = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    const s = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 17, 0);
    $('dTitle').value = '定例ミーティング（試し）';
    $('dStart').value = loc(s);
    $('dEnd').value = loc(new Date(s.getTime() + 45 * 6e4));
    $('dLoc').value = '3.Sirius（シリウス）_定員6名_本社';
    $('dRun').onclick = () => {
      ev = { id: 'demo-' + $('dTitle').value, title: $('dTitle').value, start: new Date($('dStart').value), end: new Date($('dEnd').value), location: $('dLoc').value };
      render();
    };
    $('dRun').click();
    $('url').value = getUrl();
  }

  // ---------- 起動 ----------
  $('ver').textContent = 'v' + VERSION;
  let started = false;
  const fallback = () => { if (!started) { started = true; startDemo(); } };
  if (window.Office && Office.onReady) {
    setTimeout(fallback, 3000);
    Office.onReady((info) => {
      if (started) return;
      started = true;
      inOutlook = !!(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox);
      if (!inOutlook) return startDemo();
      $('url').value = getUrl();
      if (!getUrl()) $('settings').open = true;
      readItem();
      try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, readItem); } catch (e) { /* 古い Outlook では無視 */ }
    });
  } else fallback();
})();
