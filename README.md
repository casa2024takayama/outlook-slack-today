# outlook-slack-today（予定をSlackへ）

Outlook の予定を開いてボタンを押すと、その予定を Slack ワークフロー（Webhook）経由で自分宛ての DM に送るアドイン。
iPhone の通勤ウィジェットが、その DM を読んで「今日の予定」に反映する（Slack の Outlook Calendar アプリは、当日に追加した当日の予定を通知しないため、その穴埋め）。

- 送るのはボタンを押したときだけ。Webhook URL はアドインの設定（Outlook の roamingSettings）に保存し、このリポジトリには入れない
- 予定の中身はアドインから Slack の Webhook にだけ送る

## 送る形式（cw:1）

1投稿＝1件。人が読む1行目と、JSON 1行。

```
📅 予定：定例（11/5 17:00〜17:45）
{"cw":1,"op":"upsert","id":"<itemId>","title":"定例","start":"2026-11-05T17:00:00+09:00","end":"2026-11-05T17:45:00+09:00","allDay":false,"location":"…"}
```

- `upsert`：「Slack に送る（追加・変更）」。時間や場所を変えたら、もう一度送る（同じ `id` は後の投稿が勝つ）
- `cancel`：「取り消しを送る」（2回押し）。Outlook で予定を消しても自動では送られないので、消す前に押す
- 終日：`start` は日付だけ、`end` は翌日（終わりを含まない）。編集画面では Outlook の終日フラグ、表示画面では「0:00開始・0:00終了・24時間以上」で判定
- 送らないもの
  - `snapshot`：アドインは開いている1件しか読めない（全件読むには Graph の管理者承認が要る）
  - `going`：Office.js で自分の出欠が取れない
- 今日にかからない予定は、ウィジェットが無視する旨を画面に表示（送ることはできる）

## Slack ワークフローの作り方（一度だけ）

1. Slack →「ワークフロー」→ 新規 →「Webhook から開始」
2. 変数を1つ追加：キー `message`、型「テキスト」
3. ステップ「メッセージを送信」→ 送信先に自分の名前を指定 → 本文に変数 `message` だけを入れる
4. 公開して、表示された Webhook URL をアドインの「設定」に貼って保存 →「テスト送信」
5. 届いた DM（Workflow Builder とのDM）のチャンネル ID（`D…`）を、通勤ウィジェットの `addinChannel` に設定する

## 開発

```bash
node build.mjs https://casa2024takayama.github.io/outlook-slack-today   # src/ と assets/ → docs/
node test/run.mjs                                                         # 自動テスト（下の「自動テスト」）
```

- ブラウザで `docs/taskpane.html` を直接開くと「試し画面」になる（予定を手で入れて送れる。**送信は本当に Slack に届く**）
- アドイン ID：`3e3e14fb-eaa1-4f20-abf5-a2b83b9b6281`。直したら `build.mjs` の `VERSION` を上げる
- 公開リポジトリなので、社内の名前・会議室名・Webhook URL・チャンネル ID を入れないこと

## 自動テスト

```bash
node test/run.mjs              # 単体テスト＋画面テスト（ヘッドレス Chrome）
node test/run.mjs --manifest   # ＋ manifest 検証（Microsoft の検証サービスに送る）
node test/run.mjs send-upsert  # シナリオを指定して一部だけ
```

- 単体テスト：`test/*.test.mjs`（node だけで動く）
- 画面テスト：`test/harness/`。`src/taskpane.html` から本物の Office.js を外し、偽物（`office-mock.js`）を差し込んだページを `test/.out/` に作って、Chrome を画面なしで動かす。Outlook に渡した値（予定画面・送信内容）を記録して確かめる
  - シナリオは `test/harness/scenarios.js`。1シナリオ＝1回のページ読み込みで、状態は毎回まっさら
  - 表示されているかは hidden 属性ではなく CSS の結果で判定する（Outlook の中で試し入力欄が見えていた不具合を見逃さないため）
- 実際の Outlook・Slack には何も送らない。Outlook 実機での見た目と、本物の Office.js との違いは確かめられない
- 直したら、push する前に `node test/run.mjs --manifest` を通す（manifest のバージョンが 1.0 未満だとインストールで弾かれる）

## 導入

Outlook →「アドインを取得」→「個人用アドイン」→「カスタムアドインを追加」→「ファイルから追加」で `docs/manifest.xml` を選ぶ。予定を開くと、ツールバーか「…」に **予定をSlackへ** が出る。
