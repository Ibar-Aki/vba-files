# Excel結合処理システム

[![Version](https://img.shields.io/badge/version-2.0-blue.svg)](https://github.com)
[![VBA](https://img.shields.io/badge/VBA-Excel-green.svg)](https://docs.microsoft.com/ja-jp/office/vba/api/overview/excel)

2つのExcelファイルを識別コード（キー列）で結合するVBAマクロシステムです。

---

## ✨ 特徴

- **ドラッグ＆ドロップ操作**: バッチファイルに2つのExcelファイルをD&Dするだけで実行
- **複数形式対応**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb` に対応
- **柔軟な設定**: 設定ファイルでヘッダー行数、識別コード列などをカスタマイズ可能
- **詳細なログ**: 処理結果をログシートに出力
- **エンコーディング対応**: UTF-8 / Shift_JIS 両方のソースコードを提供

---

## 📁 フォルダ構成

```
ファイル結合システム/
├── src/                              # ソースコード
│   ├── UTF-8/                        # UTF-8版（編集用）
│   │   ├── ThisWorkbook.cls          # ワークブックイベント
│   │   ├── modConstants.bas          # 定数定義
│   │   ├── modMain.bas               # メイン処理
│   │   ├── modFileHandler.bas        # ファイル処理
│   │   ├── modDataProcessor.bas      # データ結合処理
│   │   ├── modLogger.bas             # ログ処理
│   │   ├── modConfig.bas             # 設定管理
│   │   └── modValidator.bas          # 検証処理
│   └── Shift_JIS/                    # Shift_JIS版（VBAインポート用）
│       └── (上記と同一内容)
├── batch/                            # バッチファイル
│   └── Start.bat                     # 起動用バッチ
├── docs/                             # ドキュメント
│   ├── 設計書.md                     # 詳細設計書
│   └── 実装手順書.md                 # セットアップ手順
├── Config/                           # 設定ファイル格納
│   └── MergeConfig.xlsx              # 設定ファイル
├── Output/                           # 出力ファイル格納
├── Logs/                             # ログファイル格納
├── ExcelMergeEngine.xlsm             # マクロ実行エンジン（要作成）
└── README.md                         # このファイル
```

---

## 🚀 セットアップ手順

### 1. ExcelMergeEngine.xlsm の作成

1. Excelを起動し、新規ブックを作成
2. **ファイル → 名前を付けて保存** で `ExcelMergeEngine.xlsm`（マクロ有効ブック）として保存
3. **Alt + F11** でVBAエディタを開く
4. **ツール → 参照設定** で以下を有効化:
   - Microsoft Scripting Runtime

### 2. VBAモジュールのインポート

1. VBAエディタで **ファイル → ファイルのインポート**
2. `src/Shift_JIS/` フォルダから以下のファイルをすべてインポート:
   - `modConstants.bas`
   - `modMain.bas`
   - `modFileHandler.bas`
   - `modDataProcessor.bas`
   - `modLogger.bas`
   - `modConfig.bas`
   - `modValidator.bas`
3. `ThisWorkbook.cls` の内容を ThisWorkbook にコピー

### 3. 設定ファイルの作成

`Config/MergeConfig.xlsx` を作成し、`Config` シートに以下を設定:

| 設定項目 | 値 | 説明 |
|----------|-----|------|
| Excel1_HeaderRows | 3 | Excel1のヘッダー行数 |
| Excel1_DataStartRow | 4 | Excel1のデータ開始行 |
| Excel1_IDColumn | B | Excel1の識別コード列 |
| Excel2_HeaderRows | 2 | Excel2のヘッダー行数 |
| Excel2_DataStartRow | 3 | Excel2のデータ開始行 |
| Excel2_IDColumn | A | Excel2の識別コード列 |
| Output_FileNameFormat | 結合データ_[DATE].xlsx | 出力ファイル名 |
| Output_IncludeLogSheet | TRUE | ログシート出力 |

---

## 📖 使用方法

### 方法1: ドラッグ＆ドロップ

1. 結合したい2つのExcelファイルを選択
2. `batch/Start.bat` にドラッグ＆ドロップ
3. 処理完了後、`Output/` フォルダを確認

### 方法2: コマンドライン

```batch
cd batch
Start.bat "C:\path\to\Excel1.xlsx" "C:\path\to\Excel2.xlsx"
```

### 方法3: VBAから直接実行

```vba
Sub Test()
    Call ExecuteMerge("C:\path\to\Excel1.xlsx", "C:\path\to\Excel2.xlsx")
End Sub
```

---

## ⚙️ 設定のカスタマイズ

### 識別コード列の変更

設定ファイル `MergeConfig.xlsx` の `Excel1_IDColumn` / `Excel2_IDColumn` を変更します。

- 列文字（A, B, C...）または列番号（1, 2, 3...）で指定可能

### ヘッダー行数の変更

複数行ヘッダーに対応しています。`Excel1_HeaderRows` / `Excel2_HeaderRows` で設定します。

---

## 🔧 トラブルシューティング

| 症状 | 原因 | 対処方法 |
|------|------|----------|
| マクロが実行されない | セキュリティ設定 | マクロを有効化する |
| 「参照設定が見つかりません」エラー | 参照設定不足 | Microsoft Scripting Runtime を有効化 |
| 文字化け | エンコーディング | Shift_JIS版のモジュールを使用 |
| ファイルが見つからない | パス問題 | ファイルパスに日本語や特殊文字を含めない |
| メモリエラー | データ量過多 | 64bit版Excelを使用、またはデータを分割 |

---

## 📊 処理仕様

### 結合ロジック

1. Excel1とExcel2を識別コードで照合
2. 一致するデータを横方向に結合
3. 片方にしかないデータも出力（空欄補完）

### 出力ファイル

- **結合データシート**: マージされたデータ
- **処理ログシート**: 処理統計と詳細ログ

---

## 📝 バージョン履歴

| バージョン | 日付 | 変更内容 |
|------------|------|----------|
| 2.0 | 2026/01/07 | 完全リファクタリング、モジュール分離、エンコーディング対応 |
| 1.0 | 2025/07/16 | 初版 |

---

## 📄 ライセンス

MIT License

---

## 👤 作成者

Excel結合処理システム開発チーム
