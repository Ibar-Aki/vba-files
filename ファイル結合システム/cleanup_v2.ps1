$ErrorActionPreference = "Stop"

# ルートディレクトリへ移動
Set-Location ..

# 他のフォルダは移動済みなので省略

# ファイル結合システムの移動（中身移動戦略）
$sourceDir = "ファイル結合システム"
$destDir = "projects/ExcelMergeSystem"

if (Test-Path $sourceDir) {
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    # 中身を移動
    Get-ChildItem -Path $sourceDir | ForEach-Object {
        $itemPath = $_.FullName
        # スクリプト自体と.gitは除外
        if ($_.Name -ne "cleanup.ps1" -and $_.Name -ne ".git") {
            git mv "$itemPath" "$destDir/"
        }
    }
    
    Write-Host "Moved contents of $sourceDir to $destDir"
}

Write-Host "Reorganization complete."
