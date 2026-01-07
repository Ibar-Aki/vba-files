@echo off
setlocal enabledelayedexpansion

:: =============================================================
:: ファイル連番振り直しバッチ
:: 概要: 既存の番号を無視し、ファイルを名前順に並べて連番を振り直す
:: =============================================================

:: --- 設定項目 (ここを環境に合わせて変更してください) ---

REM 対象とするファイル名のパターンを指定します (ワイルドカード * が使えます)
set "FILE_PATTERN=フォーマット*.txt"

REM 新しいファイル名の接頭辞(先頭部分)を指定します
set "NEW_PREFIX=フォーマット_整理済_"

:: ---------------------------------------------------------

echo 以下の設定でファイル名の変更処理を開始します。
echo.
echo   対象ファイル: %FILE_PATTERN%
echo   新しい名前　: %NEW_PREFIX%[連番].拡張子
echo.
pause

set "count=0"

echo.
echo --- 処理開始 ---

rem /on でファイルを名前順にソートしてから処理する
for /f "delims=" %%F in ('dir /b /on "%FILE_PATTERN%"') do (
    rem 連番を1増やす
    set /a count+=1

    rem 3桁のゼロ埋め番号を作成 (例: 1 -> 001, 12 -> 012)
    set "num=000!count!"
    set "num=!num:~-3!"

    rem 新しいファイル名を生成 (接頭辞 + ゼロ埋め番号 + 元の拡張子)
    set "newName=!NEW_PREFIX!!num!%%~xF"

    rem ファイル名を変更
    ren "%%F" "!newName!"

    rem 処理結果を画面に表示
    echo [変更前]: %%F  ==^>  [変更後]: !newName!
)

echo.
echo --- 処理完了 ---
echo %count%個のファイルを処理しました。
pause