# ショートカットメールツール 仕様

## 概要
Excel で定型のメールアドレスや文面をすばやくコピーするための VBA マクロを定義する。

- `ShortcutMailTool.bas` の `ShowShortcutMailMenu` マクロを実行するとメニューが表示される。
- メニューは `sample_data.csv` から読み込んだ項目の一覧を表示する。
- 番号を指定すると該当する内容がクリップボードにコピーされる。
- コピー後は確認のメッセージを表示する。

## CSV フォーマット
- ファイル名: `sample_data.csv`
- 文字コード: UTF-8
- 1 行目はヘッダーとして `Label,Content` を配置する。
- 2 行目以降にコピーしたい項目を 1 行ずつ記述する。

### 例
```
Label,Content
サポート,support@example.com
会議案内,明日は10時から会議があります。
```

## 参照設定
クリップボード操作のため `Microsoft Forms 2.0 Object Library` を参照設定しておくこと。
