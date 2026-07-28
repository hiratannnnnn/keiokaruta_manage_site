# フォーム送信トリガー

大会フォーム回答後の業務処理は、このApps Scriptプロジェクトの
`registerFormResponseToDatabase`だけが実行する。

## 本番設定

Apps Scriptの「トリガー」画面から、次のインストール型トリガーを手動で1件だけ設定する。
コードからトリガーを作成・削除しない。

- 実行する関数: `registerFormResponseToDatabase`
- イベントのソース: スプレッドシートから
- イベントの種類: フォーム送信時
- 対象: `MAIN_SPREADSHEET_ID`のスプレッドシート

設定後、`diagnoseFormSubmitTrigger()`を手動実行し、`ok: true`、
`matching_count: 1`、`handler_count: 1`であることを確認する。

`taikai_manage`側に`memoForm`トリガーが残っている間は、同じ回答に二つの
Apps Scriptプロジェクトが反応する。切替は回答受付を一時停止した保守時間に、次の順で
連続して行う。

1. 管理サイト側のコードを`clasp push`し、通知先設定を確認する
2. `taikai_manage`側の`memoForm`トリガーを停止する
3. 管理サイト側のフォーム送信トリガーを1件だけ設定する
4. `diagnoseFormSubmitTrigger()`が`ok: true`であることを確認する
5. テストフォームを一件送信し、DB・v2・名簿・通知が各一回だけ処理されたことを確認する

両トリガーを同時に有効にした状態ではテスト回答を送信しない。テストに失敗した場合は
管理サイト側トリガーを停止してから、旧トリガーを戻すか原因を修正する。

## 再実行

非表示の「フォーム送信処理」シートは、回答ごとに次の状態を保持する。

- DB登録
- 大会シートv2書戻し
- 名簿更新
- 追加申込通知

途中失敗した回答を再処理すると、`done`または`sent`の処理は再実行せず、
未完了処理だけを続行する。行を削除したり状態を手作業で変更したりしない。
手動再処理は`retryFormResponseProcessing("大会回答シート名", 回答行番号)`を実行する。

追加申込通知は送信前に一意トークンとGmail draft IDを保存する。状態が
`delivery_unknown`の場合、自動再送は行わない。同じ通知トークンを含む送信済みメールを
確認し、見つかった状態で同じ回答を再処理して`sent`へ回復させる。

通知先はScript Propertiesの`FORM_RESPONSE_NOTIFICATION_TO`へ設定する。
