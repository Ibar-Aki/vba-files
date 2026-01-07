@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

:: ==========================================================
:: 高機能ファイル名一括置換バッチ
:: 機能: 正規表現対応、プレビュー、ログ出力、安全性チェック
:: ==========================================================

set "VERSION=1.0"
set "SCRIPT_NAME=FileRenamer"

:: カラー設定
set "C_RESET=[0m"
set "C_RED=[91m"
set "C_GREEN=[92m"
set "C_YELLOW=[93m"
set "C_BLUE=[94m"
set "C_MAGENTA=[95m"
set "C_CYAN=[96m"
set "C_WHITE=[97m"

:: 初期設定
set "TARGET_DIR=%CD%"
set "SEARCH_PATTERN="
set "REPLACE_PATTERN="
set "PREVIEW_MODE=1"
set "USE_REGEX=0"
set "CASE_SENSITIVE=0"
set "INCLUDE_SUBDIRS=0"
set "FILE_FILTER=*.*"
set "LOG_FILE="
set "BACKUP_MODE=0"

:: メイン処理開始
:MAIN
cls
call :SHOW_HEADER
echo.

:: パラメータがある場合は直接実行モード
if not "%~1"=="" (
    call :PARSE_ARGS %*
    goto :EXECUTE_RENAME
)

:: インタラクティブモード
call :INTERACTIVE_MODE
goto :END

:: ========== サブルーチン ==========

:SHOW_HEADER
echo %C_CYAN%===============================================%C_RESET%
echo %C_CYAN%  %SCRIPT_NAME% v%VERSION%  %C_RESET%
echo %C_CYAN%  高機能ファイル名一括置換ツール%C_RESET%
echo %C_CYAN%===============================================%C_RESET%
goto :EOF

:INTERACTIVE_MODE
echo %C_WHITE%現在の設定:%C_RESET%
echo   対象ディレクトリ: %C_YELLOW%!TARGET_DIR!%C_RESET%
echo   検索パターン: %C_YELLOW%!SEARCH_PATTERN!%C_RESET%
echo   置換パターン: %C_YELLOW%!REPLACE_PATTERN!%C_RESET%
echo   ファイルフィルタ: %C_YELLOW%!FILE_FILTER!%C_RESET%
echo   プレビューモード: %C_YELLOW%!PREVIEW_MODE!%C_RESET%
echo   正規表現使用: %C_YELLOW%!USE_REGEX!%C_RESET%
echo   大文字小文字区別: %C_YELLOW%!CASE_SENSITIVE!%C_RESET%
echo   サブディレクトリ含む: %C_YELLOW%!INCLUDE_SUBDIRS!%C_RESET%
echo   バックアップモード: %C_YELLOW%!BACKUP_MODE!%C_RESET%
echo.

echo %C_GREEN%メニュー:%C_RESET%
echo   %C_WHITE%1%C_RESET% - 対象ディレクトリを設定
echo   %C_WHITE%2%C_RESET% - 検索パターンを設定
echo   %C_WHITE%3%C_RESET% - 置換パターンを設定
echo   %C_WHITE%4%C_RESET% - ファイルフィルタを設定
echo   %C_WHITE%5%C_RESET% - オプション設定
echo   %C_WHITE%6%C_RESET% - プレビュー実行
echo   %C_WHITE%7%C_RESET% - 実行（実際にリネーム）
echo   %C_WHITE%8%C_RESET% - ヘルプ表示
echo   %C_WHITE%0%C_RESET% - 終了
echo.

set /p "CHOICE=選択してください (0-8): "

if "%CHOICE%"=="1" call :SET_TARGET_DIR
if "%CHOICE%"=="2" call :SET_SEARCH_PATTERN
if "%CHOICE%"=="3" call :SET_REPLACE_PATTERN
if "%CHOICE%"=="4" call :SET_FILE_FILTER
if "%CHOICE%"=="5" call :SET_OPTIONS
if "%CHOICE%"=="6" call :PREVIEW_RENAME
if "%CHOICE%"=="7" call :EXECUTE_RENAME
if "%CHOICE%"=="8" call :SHOW_HELP
if "%CHOICE%"=="0" goto :END

goto :INTERACTIVE_MODE

:SET_TARGET_DIR
echo.
echo %C_CYAN%対象ディレクトリの設定%C_RESET%
set /p "NEW_DIR=ディレクトリパスを入力 (現在: !TARGET_DIR!): "
if not "!NEW_DIR!"=="" (
    if exist "!NEW_DIR!" (
        set "TARGET_DIR=!NEW_DIR!"
        echo %C_GREEN%設定完了: !TARGET_DIR!%C_RESET%
    ) else (
        echo %C_RED%エラー: 指定されたディレクトリが存在しません%C_RESET%
    )
)
pause
goto :EOF

:SET_SEARCH_PATTERN
echo.
echo %C_CYAN%検索パターンの設定%C_RESET%
set /p "NEW_PATTERN=検索パターンを入力 (現在: !SEARCH_PATTERN!): "
if not "!NEW_PATTERN!"=="" set "SEARCH_PATTERN=!NEW_PATTERN!"
echo %C_GREEN%設定完了: !SEARCH_PATTERN!%C_RESET%
pause
goto :EOF

:SET_REPLACE_PATTERN
echo.
echo %C_CYAN%置換パターンの設定%C_RESET%
set /p "NEW_PATTERN=置換パターンを入力 (現在: !REPLACE_PATTERN!): "
set "REPLACE_PATTERN=!NEW_PATTERN!"
echo %C_GREEN%設定完了: !REPLACE_PATTERN!%C_RESET%
pause
goto :EOF

:SET_FILE_FILTER
echo.
echo %C_CYAN%ファイルフィルタの設定%C_RESET%
echo 例: *.txt, *.jpg, *.*, IMG_*.jpg
set /p "NEW_FILTER=フィルタを入力 (現在: !FILE_FILTER!): "
if not "!NEW_FILTER!"=="" set "FILE_FILTER=!NEW_FILTER!"
echo %C_GREEN%設定完了: !FILE_FILTER!%C_RESET%
pause
goto :EOF

:SET_OPTIONS
echo.
echo %C_CYAN%オプション設定%C_RESET%
echo   %C_WHITE%1%C_RESET% - 正規表現使用 (現在: !USE_REGEX!)
echo   %C_WHITE%2%C_RESET% - 大文字小文字区別 (現在: !CASE_SENSITIVE!)
echo   %C_WHITE%3%C_RESET% - サブディレクトリ含む (現在: !INCLUDE_SUBDIRS!)
echo   %C_WHITE%4%C_RESET% - バックアップモード (現在: !BACKUP_MODE!)
echo   %C_WHITE%5%C_RESET% - ログファイル設定 (現在: !LOG_FILE!)
echo   %C_WHITE%0%C_RESET% - 戻る
echo.

set /p "OPT_CHOICE=選択してください (0-5): "

if "%OPT_CHOICE%"=="1" call :TOGGLE_OPTION USE_REGEX
if "%OPT_CHOICE%"=="2" call :TOGGLE_OPTION CASE_SENSITIVE
if "%OPT_CHOICE%"=="3" call :TOGGLE_OPTION INCLUDE_SUBDIRS
if "%OPT_CHOICE%"=="4" call :TOGGLE_OPTION BACKUP_MODE
if "%OPT_CHOICE%"=="5" call :SET_LOG_FILE
if "%OPT_CHOICE%"=="0" goto :EOF

goto :SET_OPTIONS

:TOGGLE_OPTION
set "VAR_NAME=%1"
if "!!VAR_NAME!!"=="0" (
    set "%VAR_NAME%=1"
    echo %C_GREEN%!VAR_NAME! を有効にしました%C_RESET%
) else (
    set "%VAR_NAME%=0"
    echo %C_YELLOW%!VAR_NAME! を無効にしました%C_RESET%
)
pause
goto :EOF

:SET_LOG_FILE
echo.
set /p "NEW_LOG=ログファイルパスを入力 (空白で無効): "
set "LOG_FILE=!NEW_LOG!"
if "!LOG_FILE!"=="" (
    echo %C_YELLOW%ログ出力を無効にしました%C_RESET%
) else (
    echo %C_GREEN%ログファイル: !LOG_FILE!%C_RESET%
)
pause
goto :EOF

:PREVIEW_RENAME
if "!SEARCH_PATTERN!"=="" (
    echo %C_RED%エラー: 検索パターンが設定されていません%C_RESET%
    pause
    goto :EOF
)

echo.
echo %C_CYAN%プレビュー実行中...%C_RESET%
echo.

set "RENAME_COUNT=0"
call :PROCESS_FILES 1

echo.
echo %C_GREEN%プレビュー完了: !RENAME_COUNT! 件のファイルがリネーム対象です%C_RESET%
pause
goto :EOF

:EXECUTE_RENAME
if "!SEARCH_PATTERN!"=="" (
    echo %C_RED%エラー: 検索パターンが設定されていません%C_RESET%
    pause
    goto :EOF
)

echo.
echo %C_YELLOW%警告: 実際にファイルをリネームします%C_RESET%
set /p "CONFIRM=続行しますか？ (y/N): "
if /i not "!CONFIRM!"=="y" goto :EOF

echo.
echo %C_CYAN%リネーム実行中...%C_RESET%
echo.

if not "!LOG_FILE!"=="" (
    echo [%DATE% %TIME%] リネーム処理開始 > "!LOG_FILE!"
)

set "RENAME_COUNT=0"
set "ERROR_COUNT=0"
call :PROCESS_FILES 0

echo.
echo %C_GREEN%処理完了: !RENAME_COUNT! 件のファイルをリネームしました%C_RESET%
if !ERROR_COUNT! gtr 0 echo %C_RED%エラー: !ERROR_COUNT! 件の処理に失敗しました%C_RESET%

if not "!LOG_FILE!"=="" (
    echo [%DATE% %TIME%] 処理完了: !RENAME_COUNT! 件成功, !ERROR_COUNT! 件失敗 >> "!LOG_FILE!"
)

pause
goto :EOF

:PROCESS_FILES
set "IS_PREVIEW=%1"

if "!INCLUDE_SUBDIRS!"=="1" (
    set "SEARCH_OPTION=/s"
) else (
    set "SEARCH_OPTION="
)

pushd "!TARGET_DIR!"

for /f "delims=" %%F in ('dir /b !SEARCH_OPTION! "!FILE_FILTER!" 2^>nul') do (
    set "ORIGINAL_NAME=%%~nxF"
    set "NEW_NAME=!ORIGINAL_NAME!"
    
    :: 置換処理
    if "!USE_REGEX!"=="1" (
        :: 正規表現モード（簡易実装）
        call :REGEX_REPLACE "!NEW_NAME!" "!SEARCH_PATTERN!" "!REPLACE_PATTERN!" NEW_NAME
    ) else (
        :: 通常の文字列置換
        if "!CASE_SENSITIVE!"=="1" (
            set "NEW_NAME=!NEW_NAME:%%SEARCH_PATTERN%%=%%REPLACE_PATTERN%%!"
        ) else (
            call :CASE_INSENSITIVE_REPLACE "!NEW_NAME!" "!SEARCH_PATTERN!" "!REPLACE_PATTERN!" NEW_NAME
        )
    )
    
    :: 変更があるかチェック
    if not "!ORIGINAL_NAME!"=="!NEW_NAME!" (
        set /a RENAME_COUNT+=1
        
        if "!IS_PREVIEW!"=="1" (
            echo %C_WHITE%!ORIGINAL_NAME!%C_RESET% %C_MAGENTA%→%C_RESET% %C_GREEN%!NEW_NAME!%C_RESET%
        ) else (
            :: バックアップ作成
            if "!BACKUP_MODE!"=="1" (
                copy "!ORIGINAL_NAME!" "!ORIGINAL_NAME!.bak" >nul 2>&1
            )
            
            :: 実際のリネーム
            ren "!ORIGINAL_NAME!" "!NEW_NAME!" 2>nul
            if errorlevel 1 (
                set /a ERROR_COUNT+=1
                echo %C_RED%エラー: !ORIGINAL_NAME! のリネームに失敗%C_RESET%
                if not "!LOG_FILE!"=="" (
                    echo [%DATE% %TIME%] エラー: !ORIGINAL_NAME! → !NEW_NAME! >> "!LOG_FILE!"
                )
            ) else (
                echo %C_WHITE%!ORIGINAL_NAME!%C_RESET% %C_MAGENTA%→%C_RESET% %C_GREEN%!NEW_NAME!%C_RESET%
                if not "!LOG_FILE!"=="" (
                    echo [%DATE% %TIME%] 成功: !ORIGINAL_NAME! → !NEW_NAME! >> "!LOG_FILE!"
                )
            )
        )
    )
)

popd
goto :EOF

:CASE_INSENSITIVE_REPLACE
set "STR=%~1"
set "SEARCH=%~2"
set "REPLACE=%~3"
set "RESULT_VAR=%~4"

:: 大文字小文字を区別しない置換（簡易実装）
set "TEMP_STR=!STR!"
for %%A in (a b c d e f g h i j k l m n o p q r s t u v w x y z) do (
    for %%B in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        set "TEMP_STR=!TEMP_STR:%%A=%%B!"
    )
)
set "TEMP_SEARCH=!SEARCH!"
for %%A in (a b c d e f g h i j k l m n o p q r s t u v w x y z) do (
    for %%B in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        set "TEMP_SEARCH=!TEMP_SEARCH:%%A=%%B!"
    )
)

if not "!TEMP_STR!"=="!TEMP_STR:%TEMP_SEARCH%=%REPLACE%!" (
    set "%RESULT_VAR%=!STR:%SEARCH%=%REPLACE%!"
) else (
    set "%RESULT_VAR%=!STR!"
)
goto :EOF

:REGEX_REPLACE
:: 簡易正規表現実装（基本的なパターンのみ）
set "STR=%~1"
set "PATTERN=%~2"
set "REPLACE=%~3"
set "RESULT_VAR=%~4"

:: ここでは基本的な置換のみ実装
set "%RESULT_VAR%=!STR:%PATTERN%=%REPLACE%!"
goto :EOF

:SHOW_HELP
cls
call :SHOW_HEADER
echo.
echo %C_GREEN%使用方法:%C_RESET%
echo   %C_WHITE%インタラクティブモード:%C_RESET%
echo     %SCRIPT_NAME%.bat
echo.
echo   %C_WHITE%コマンドラインモード:%C_RESET%
echo     %SCRIPT_NAME%.bat -d "ディレクトリ" -s "検索" -r "置換" [オプション]
echo.
echo %C_GREEN%オプション:%C_RESET%
echo   %C_WHITE%-d DIR%C_RESET%      対象ディレクトリ
echo   %C_WHITE%-s PATTERN%C_RESET%  検索パターン
echo   %C_WHITE%-r PATTERN%C_RESET%  置換パターン
echo   %C_WHITE%-f FILTER%C_RESET%   ファイルフィルタ (デフォルト: *.*)
echo   %C_WHITE%-p%C_RESET%          プレビューモードで実行
echo   %C_WHITE%-x%C_RESET%          正規表現を使用
echo   %C_WHITE%-c%C_RESET%          大文字小文字を区別
echo   %C_WHITE%-sub%C_RESET%        サブディレクトリも対象
echo   %C_WHITE%-b%C_RESET%          バックアップを作成
echo   %C_WHITE%-l FILE%C_RESET%     ログファイルを指定
echo.
echo %C_GREEN%例:%C_RESET%
echo   %C_YELLOW%IMG を Photo に置換:%C_RESET%
echo     %SCRIPT_NAME%.bat -s "IMG" -r "Photo" -p
echo.
echo   %C_YELLOW%拡張子を変更:%C_RESET%
echo     %SCRIPT_NAME%.bat -f "*.jpeg" -s ".jpeg" -r ".jpg"
echo.
pause
goto :EOF

:PARSE_ARGS
:ARG_LOOP
if "%~1"=="" goto :EOF
if "%~1"=="-d" (
    set "TARGET_DIR=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-s" (
    set "SEARCH_PATTERN=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-r" (
    set "REPLACE_PATTERN=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-f" (
    set "FILE_FILTER=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-l" (
    set "LOG_FILE=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-p" (
    set "PREVIEW_MODE=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-x" (
    set "USE_REGEX=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-c" (
    set "CASE_SENSITIVE=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-sub" (
    set "INCLUDE_SUBDIRS=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-b" (
    set "BACKUP_MODE=1"
    shift
    goto :ARG_LOOP
)
shift
goto :ARG_LOOP

:END
echo.
echo %C_CYAN%ご利用ありがとうございました。%C_RESET%
pause >nul
exit /b 0
