@echo off
setlocal enabledelayedexpansion

:: -------------------------------------------------
:: 高機能ファイルリネームバッチ (Ver.3)
:: 機能: 日付、時刻、カスタム文字をファイル名に追加
:: -------------------------------------------------

:: --- 初期設定 ---
set "TARGET_FOLDER=%cd%"  REM 対象フォルダ (規定値: このバッチファイルがある場所)
set "FILE_PATTERN=*.*"    REM 対象ファイル (規定値: 全てのファイル)
set "RECURSIVE=/R"        REM サブフォルダを対象にするか (/R) 空欄なら対象外

:MENU
cls
echo =================================================================
echo  高機能 ファイル名一括変更ツール
echo =================================================================
echo.
echo   現在の設定:
echo     対象フォルダ: %TARGET_FOLDER%
echo     対象ファイル: %FILE_PATTERN%
echo     サブフォルダ: %RECURSIVE% ( /R=含む )
echo.
echo -----------------------------------------------------------------
echo.
echo  [ 追加するものを選択してください ]
echo.
echo    1. ファイル名の【先頭】に [YYYYMMDD_] を追加
echo    2. ファイル名の【末尾】に [_YYYYMMDD] を追加
echo    3. ファイル名の【末尾】に [_YYYYMMDD-HHMMSS] を追加
echo.
echo    4. ファイル名の【先頭】に [カスタム文字_] を追加
echo    5. ファイル名の【末尾】に [_カスタム文字] を追加 
echo.
echo    8. [テスト実行(Dry Run)] を行う (名前の変更はしません)
echo    9. [設定の変更]
echo.
echo    Q. 終了
echo.
echo -----------------------------------------------------------------

set /p "CHOICE=番号を選択してください: "

if /i "%CHOICE%"=="1" call :ProcessFiles "PREFIX_DATE" && goto MENU
if /i "%CHOICE%"=="2" call :ProcessFiles "SUFFIX_DATE" && goto MENU
if /i "%CHOICE%"=="3" call :ProcessFiles "SUFFIX_DATETIME" && goto MENU
if /i "%CHOICE%"=="4" call :ProcessFiles "PREFIX_CUSTOM" && goto MENU
if /i "%CHOICE%"=="5" call :ProcessFiles "SUFFIX_CUSTOM" && goto MENU
if /i "%CHOICE%"=="8" call :ProcessFiles "DRY_RUN" && goto MENU
if /i "%CHOICE%"=="9" goto SETTINGS
if /i "%CHOICE%"=="Q" goto EOF

echo 無効な選択です。
pause
goto MENU


:SETTINGS
cls
echo.
echo --- 設定の変更 ---
echo   現在の設定値が表示されています。変更しない場合はそのままEnterを押してください。
echo.
set /p "FILE_PATTERN=対象ファイルのパターンを入力 (例: *.jpg, *.pptx, *.*) [%FILE_PATTERN%]: "
set /p "TARGET_FOLDER=対象フォルダのパスを入力してください [%TARGET_FOLDER%]: "
set /p "RECURSIVE_INPUT=サブフォルダも対象にしますか？ (Y/N) : "
if /i "%RECURSIVE_INPUT%"=="Y" (set "RECURSIVE=/R") else (set "RECURSIVE=")
goto MENU


:ProcessFiles
cls
set "MODE=%~1"

rem --- 日付と時刻の書式設定 (YYYYMMDD と HHMMSS) ---
set "DATE_STR=%date:~0,4%%date:~5,2%%date:~8,2%"
set "TIME_STR=%time:~0,2%%time:~3,2%%time:~6,2%"
set "TIME_STR=%TIME_STR: =0%"

set "TIMESTAMP_DATE=%DATE_STR%"
set "TIMESTAMP_DATETIME=%DATE_STR%-%TIME_STR%"

rem --- カスタム文字の入力 ---
if "%MODE%"=="PREFIX_CUSTOM" or "%MODE%"=="SUFFIX_CUSTOM" (
    if "%MODE%"=="PREFIX_CUSTOM" set "POS=先頭"
    if "%MODE%"=="SUFFIX_CUSTOM" set "POS=末尾"
    set /p "CUSTOM_TEXT=ファイル名の!POS!に追加する文字を入力してください: "
    if not defined CUSTOM_TEXT (
        echo 文字が入力されませんでした。メニューに戻ります。
        pause
        exit /b
    )
)

echo.
echo --- 処理を開始します ---
echo モード: %MODE%
echo.

set "COUNT=0"

if "%MODE%"=="DRY_RUN" (
    echo ★★★ テスト実行モードです ★★★
    echo どのように名前が変更されるかを表示します。実際の変更は行われません。
    echo.
    echo [変更前] ==^> [変更後]
    echo ----------------------------------
)

for %RECURSIVE% %%F in ("%TARGET_FOLDER%\%FILE_PATTERN%") do (
    set "FILE_NAME=%%~nF"
    set "FILE_EXT=%%~xF"
    set "NEW_NAME="

    rem --- モードに応じて新しいファイル名を生成 ---
    if "%MODE%"=="PREFIX_DATE" set "NEW_NAME=%TIMESTAMP_DATE%_!FILE_NAME!!FILE_EXT!"
    if "%MODE%"=="SUFFIX_DATE" set "NEW_NAME=!FILE_NAME!_%TIMESTAMP_DATE%!FILE_EXT!"
    if "%MODE%"=="SUFFIX_DATETIME" set "NEW_NAME=!FILE_NAME!_%TIMESTAMP_DATETIME%!FILE_EXT!"
    if "%MODE%"=="PREFIX_CUSTOM" set "NEW_NAME=%CUSTOM_TEXT%_!FILE_NAME!!FILE_EXT!"
    if "%MODE%"=="SUFFIX_CUSTOM" set "NEW_NAME=!FILE_NAME!_%CUSTOM_TEXT%!FILE_EXT!"

    if defined NEW_NAME (
        if "%MODE%"=="DRY_RUN" (
            echo "%%~nxF" ==^> "!NEW_NAME!"
        ) else (
            if not "%%~nxF"=="!NEW_NAME!" (
                ren "%%F" "!NEW_NAME!"
                echo RENAMED: "%%~nxF" --^> "!NEW_NAME!"
            )
        )
        set /a COUNT+=1
    )
)

echo.
echo ----------------------------------
echo 処理が完了しました: %COUNT% 件のファイルを処理しました。
echo.
pause
exit /b

:EOF
endlocal
exit