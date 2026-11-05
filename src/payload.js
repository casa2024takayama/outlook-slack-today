// 予定 → commute-widget が読む cw:1 形式の JSON と、Slack に送る本文を作る
// ブラウザでは window.CwPayload、Node では module.exports
(function (global) {
  'use strict';
  const pad = (n) => String(n).padStart(2, '0');
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const hm = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  // 端末のタイムゾーンで +09:00 のような表記にする
  const offset = (d) => {
    const m = -d.getTimezoneOffset();
    const s = m >= 0 ? '+' : '-';
    return `${s}${pad(Math.floor(Math.abs(m) / 60))}:${pad(Math.abs(m) % 60)}`;
  };
  const iso = (d) => `${ymd(d)}T${hm(d)}:${pad(d.getSeconds())}${offset(d)}`;

  // 終日の判定：Outlook が教えてくれない環境では「0:00 開始・0:00 終了・24時間の倍数」で判断
  const looksAllDay = (s, e) => s.getHours() === 0 && s.getMinutes() === 0 && e.getHours() === 0 && e.getMinutes() === 0 && e - s >= 864e5 - 36e5;

  /**
   * @param {{id:string,title:string,start:Date,end:Date,allDay?:boolean,location?:string}} ev
   */
  function upsert(ev) {
    const allDay = ev.allDay != null ? !!ev.allDay : looksAllDay(ev.start, ev.end);
    const o = { cw: 1, op: 'upsert', id: ev.id, title: ev.title || '（件名なし）' };
    if (allDay) {
      o.start = ymd(ev.start);
      // 終わりは含まない日付。0:00 終了ならその日、それ以外は翌日
      const e = new Date(ev.end);
      if (e.getHours() || e.getMinutes()) e.setDate(e.getDate() + 1);
      o.end = ymd(e <= ev.start ? new Date(ev.start.getFullYear(), ev.start.getMonth(), ev.start.getDate() + 1) : e);
      o.allDay = true;
    } else {
      o.start = iso(ev.start);
      o.end = iso(ev.end);
      o.allDay = false;
    }
    if (ev.location) o.location = ev.location;
    return o;
  }
  const cancel = (ev) => ({ cw: 1, op: 'cancel', id: ev.id, title: ev.title || '' });

  // 人が読む1行目＋JSON（コードブロックなし。ウィジェットは JSON だけでも読める）
  function message(o) {
    const when = o.allDay ? (o.start === prevDay(o.end) ? `${md(o.start)} 終日` : `${md(o.start)}〜${md(prevDay(o.end))} 終日`)
      : o.start ? `${md(o.start.slice(0, 10))} ${o.start.slice(11, 16)}〜${o.end.slice(11, 16)}` : '';
    const head = o.op === 'cancel' ? `🗑 予定の取り消し：${o.title}` : `📅 予定：${o.title}${when ? '（' + when + '）' : ''}`;
    return `${head}\n${JSON.stringify(o)}`;
  }
  const md = (s) => { const [, m, d] = s.split('-').map(Number); return `${m}/${d}`; };
  const prevDay = (s) => { const [y, m, d] = s.split('-').map(Number); return ymd(new Date(y, m - 1, d - 1)); };

  // 今日にかかるか（ウィジェットは今日の分だけ表示する）
  function touchesToday(o, now = new Date()) {
    const t = ymd(now);
    if (o.allDay) return o.start <= t && t < o.end;
    return o.start.slice(0, 10) <= t && t <= o.end.slice(0, 10);
  }

  const api = { upsert, cancel, message, touchesToday, iso, ymd, looksAllDay };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.CwPayload = api;
})(typeof window !== 'undefined' ? window : globalThis);
