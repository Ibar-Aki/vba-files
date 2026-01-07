$ErrorActionPreference = "Stop"

# ルートディレクトリへ移動
Set-Location ..

# Git管理下に追加（未追跡ファイルも含めるため）
git add .

# ディレクトリ作成
if (-not (Test-Path "projects")) { New-Item -ItemType Directory -Path "projects" }
if (-not (Test-Path "docs")) { New-Item -ItemType Directory -Path "docs" }
if (-not (Test-Path "utils")) { New-Item -ItemType Directory -Path "utils" }
if (-not (Test-Path "misc")) { New-Item -ItemType Directory -Path "misc" }

# ファイル・フォルダ移動関数
function Move-GitItem {
    param (
        [string]$Source,
        [string]$Dest
    )
    if (Test-Path $Source) {
        Write-Host "Moving $Source to $Dest"
        # フォルダの場合は親フォルダが存在することを確認
        $parent = Split-Path $Dest -Parent
        if (-not (Test-Path $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        
        # git mv 実行
        # Windowsのパス区切り文字に対応するため、パスを調整
        git mv "$Source" "$Dest"
    } else {
        Write-Host "Skipping $Source (Not found)"
    }
}

# プロジェクト群の移動
Move-GitItem "ファイル結合システム" "projects/ExcelMergeSystem"
Move-GitItem "LS入力" "projects/LSEntry"
Move-GitItem "shortcut-mail-tool" "projects/ShortcutMailTool"
Move-GitItem "WBS作成" "projects/WBSGenerator"
Move-GitItem "転記システム" "projects/TranscriptionSystem"
Move-GitItem "Useful Item" "projects/UsefulItems"

# 単体ファイルのプロジェクト化
if (Test-Path "Quick Text Access.xlsm") {
    $destDir = "projects/QuickTextAccess"
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    git mv "Quick Text Access.xlsm" "$destDir/Quick Text Access.xlsm"
}

# ドキュメント群の移動
Move-GitItem "Gem, プロンプト" "docs/Prompts"
Move-GitItem "各種ドキュメント" "docs/General"

# ユーティリティ群の移動
Move-GitItem "各種バッチファイル" "utils/BatchFiles"

# その他
Move-GitItem "一時メモ" "misc/Temp"

Write-Host "Reorganization complete."
