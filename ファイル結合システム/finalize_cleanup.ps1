$ErrorActionPreference = "Stop"

# ルートディレクトリへ移動
Set-Location ..

# README.md の内容
$readmeContent = @"
# VBA Files Repository

このリポジトリは、様々なVBAマクロ、ツール、ドキュメントを管理しています。

## 📁 ディレクトリ構成

### 🏗️ Projects (projects/)
各ツールやシステムごとのソースコードとリソースが含まれています。

- **[ExcelMergeSystem](projects/ExcelMergeSystem)**: Excelファイル結合システム
- **[LSEntry](projects/LSEntry)**: LS入力システム
- **[ShortcutMailTool](projects/ShortcutMailTool)**: メール作成ショートカットツール
- **[WBSGenerator](projects/WBSGenerator)**: WBS自動作成ツール
- **[TranscriptionSystem](projects/TranscriptionSystem)**: 転記システム
- **[UsefulItems](projects/UsefulItems)**: 便利なアイテム集
- **[QuickTextAccess](projects/QuickTextAccess)**: クイックテキストアクセスツール

### 📚 Docs (docs/)
ドキュメントやプロンプト集です。

- **[Prompts](docs/Prompts)**: AI用プロンプト集
- **[General](docs/General)**: 一般ドキュメント

### 🛠️ Utils (utils/)
ユーティリティやバッチファイルです。

- **[BatchFiles](utils/BatchFiles)**: 各種バッチファイル

### 📦 Misc (misc/)
- **Temp**: 一時ファイルなど

## 🚀 更新履歴

- **2026/01/08**: リポジトリ全体の構成を整理しました。

"@

# README.md 作成
[System.IO.File]::WriteAllText("README.md", $readmeContent, [System.Text.Encoding]::UTF8)
Write-Host "Created README.md"

# 空になったフォルダの削除（ファイル結合システム）
$oldDir = "ファイル結合システム"
if (Test-Path $oldDir) {
    # 中身がまだあるか確認
    $remaining = Get-ChildItem -Path $oldDir
    if ($remaining.Count -eq 0 -or ($remaining.Count -eq 1 -and $remaining[0].Name -eq "finalize_cleanup.ps1")) {
        # 空（またはこのスクリプトだけ）なら削除
        # ただしカレントディレクトリにいると削除できないので注意が必要だが、
        # 今は .. に移動しているので大丈夫なはず
        
        Write-Host "Removing empty directory: $oldDir"
        # git clean -fd で消えるはずだが、明示的に消す
        Remove-Item -Path $oldDir -Recurse -Force
    } else {
        Write-Host "Directory $oldDir is not empty, skipping removal."
        $remaining | ForEach-Object { Write-Host ("- " + $_.Name) }
    }
}

# Git コミット & プッシュ
git add .
git commit -m "chore: Reorganize repository structure"
git push origin main

Write-Host "Repository cleanup completed."
