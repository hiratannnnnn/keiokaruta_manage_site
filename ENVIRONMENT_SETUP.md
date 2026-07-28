# 環境設定

外部リソースIDと実行環境別の値はソースコードへ記載せず、Apps Scriptの
Script Propertiesから取得します。

1. `.env.example`を参考に、ローカルの`.env`へ値を記入する
2. `.env`の各キーを、同じ名前でApps ScriptのScript Propertiesへ登録する
3. 管理サイトの設定ページにある「環境設定診断」で登録状態を確認する
4. 未設定がないことを確認してからデプロイする

`.env`はGitとclaspの対象外です。設定診断にも値そのものは表示されません。

フォーム送信後の追加申込通知先は`FORM_RESPONSE_NOTIFICATION_TO`へ登録します。
宛先をソースコードへ直接記載しません。

LINE連携を使用する場合は、推測困難で互いに異なる
`LINE_LINK_WEBHOOK_SECRET`と`LINE_LINK_BINDING_SECRET`も登録します。
前者はtaikai_manageからGAS Web Appを呼ぶ共有秘密、後者はLINEユーザーIDを
API送信用のbinding hashへ変換するGAS内限定秘密です。
