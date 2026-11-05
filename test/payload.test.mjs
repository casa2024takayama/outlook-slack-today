import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const C = createRequire(import.meta.url)('../src/payload.js');
// 時刻つき
const a = C.upsert({ id: 'AAMk1', title: '定例', start: new Date(2026,10,5,17,0), end: new Date(2026,10,5,17,45), location: '3.Sirius（シリウス）_定員6名_本社' });
assert.deepEqual(a, { cw:1, op:'upsert', id:'AAMk1', title:'定例', start:'2026-11-05T17:00:00+09:00', end:'2026-11-05T17:45:00+09:00', allDay:false, location:'3.Sirius（シリウス）_定員6名_本社' });
// 終日（0:00〜翌0:00 で推定）
const b = C.upsert({ id: 'x', title: '休み', start: new Date(2026,10,5), end: new Date(2026,10,6) });
assert.equal(b.start, '2026-11-05'); assert.equal(b.end, '2026-11-06'); assert.equal(b.allDay, true);
// 終日 2日間（明示）
const c = C.upsert({ id: 'y', title: '出張', start: new Date(2026,10,5), end: new Date(2026,10,7), allDay: true });
assert.equal(c.end, '2026-11-07');
assert.ok(C.touchesToday(c, new Date(2026,10,6,12)));
assert.ok(!C.touchesToday(c, new Date(2026,10,7,1)));
assert.ok(!C.touchesToday(a, new Date(2026,10,6)));
// 0:00〜0:30 は終日ではない
assert.equal(C.upsert({ id:'z', title:'深夜', start:new Date(2026,10,5,0,0), end:new Date(2026,10,5,0,30) }).allDay, false);
// cancel と本文
const m = C.message(C.cancel({ id:'AAMk1', title:'定例' }));
assert.equal(m, '🗑 予定の取り消し：定例\n{"cw":1,"op":"cancel","id":"AAMk1","title":"定例"}');
console.log(C.message(a)); console.log(C.message(c));
console.log('all ok');
