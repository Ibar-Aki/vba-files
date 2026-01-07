@echo off
setlocal enabledelayedexpansion
chcp 932 >nul

:: ==========================================================
:: シンプルファイル名一括置換バッチ
:: 同一ディレクトリ内のファイルを対象
:: ==========================================================

echo.
echo =====================================
echo  ファイル名一括置換ツール
echo =====================================
echo.

:: 現在のディレクトリ内のファイル一覧表示
echo 現在のディレクトリ: %CD%
echo.
echo ファイル一覧:
dir /b *.* 2>nul | findstr /v /c:"%~nx0"
echo.

:: 検索パターン入力
:INPUT_SEARCH
set /p "SEARCH_TEXT=置換前の文字列を入力してください: "
if "%SEARCH_TEXT%"=="" (
    echo エラー: 検索文字列を入力してください
    goto INPUT_SEARCH
)

:: 置換パターン入力
set /p "REPLACE_TEXT=置換後の文字列を入力してください（空白も可）: "

:: ファイルフィルタ入力（オプション）
echo.
echo ファイルフィルタ（例: *.txt, *.jpg, 何も入力しなければ全ファイル）
set /p "FILE_FILTER=フィルタ: "
if "%FILE_FILTER%"=="" set "FILE_FILTER=*.*"

echo.
echo =====================================
echo プレビュー（実際の変更は行いません）
echo =====================================
echo.

set "COUNT=0"
set "PREVIEW_LIST="

:: プレビュー処理
for /f "delims=" %%F in ('dir /b "%FILE_FILTER%" 2^>nul') do (
    if not "%%F"=="%~nx0" (
        set "ORIGINAL=%%F"
        set "NEW_NAME=!ORIGINAL:%SEARCH_TEXT%=%REPLACE_TEXT%!"
        
        if not "!ORIGINAL!"=="!NEW_NAME!" (
            set /a COUNT+=1
            echo !COUNT!. !ORIGINAL! --^> !NEW_NAME!
            set "PREVIEW_LIST=!PREVIEW_LIST! "%%F""
        )
    )
)

if %COUNT%==0 (
    echo 該当するファイルが見つかりませんでした。
    echo.
    pause
    exit /b 0
)

echo.
echo 合計 %COUNT% 個のファイルが対象です。
echo.

:: 実行確認
echo 実際にリネームを実行しますか？
echo.
echo [Enter] 実行する
echo [その他のキー + Enter] キャンセル
echo.
set /p "CONFIRM=> "
if "%CONFIRM%"=="" goto EXECUTE_RENAME

echo キャンセルしました。
pause
exit /b 0

:EXECUTE_RENAME
echo.
echo =====================================
echo リネーム実行中...
echo =====================================
echo.

set "SUCCESS_COUNT=0"
set "ERROR_COUNT=0"

:: 実際のリネーム処理
for /f "delims=" %%F in ('dir /b "%FILE_FILTER%" 2^>nul') do (
    if not "%%F"=="%~nx0" (
        set "ORIGINAL=%%F"
        set "NEW_NAME=!ORIGINAL:%SEARCH_TEXT%=%REPLACE_TEXT%!"
        
        if not "!ORIGINAL!"=="!NEW_NAME!" (
            ren "!ORIGINAL!" "!NEW_NAME!" 2>nul
            if !errorlevel!==0 (
                set /a SUCCESS_COUNT+=1
                echo 成功: !ORIGINAL! --^> !NEW_NAME!
            ) else (
                set /a ERROR_COUNT+=1
                echo エラー: !ORIGINAL! のリネームに失敗しました
            )
        )
    )
)

echo.
echo =====================================
echo 処理完了
echo =====================================
echo 成功: %SUCCESS_COUNT% 件
echo エラー: %ERROR_COUNT% 件
echo.

pause
exit /b 0