@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

:: ==================================================
:: Excel結合処理システム 起動バッチ
:: Version: 2.0
:: Date: 2026/01/07
:: ==================================================

:: 設定
set ENGINE_FILE=%~dp0..\ExcelMergeEngine.xlsm
set LOG_DIR=%~dp0..\Logs
set OUTPUT_DIR=%~dp0..\Output

:: カラー設定
color 0A

:: ディレクトリ作成
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: タイトル表示
cls
echo.
echo ============================================
echo    Excel結合処理システム v2.0
echo ============================================
echo.

:: 引数チェック
if "%~1"=="" goto USAGE
if "%~2"=="" goto USAGE

:: ファイル存在確認
if not exist "%~1" (
    color 0C
    echo [エラー] ファイルが見つかりません:
    echo          %~1
    goto ERROR_EXIT
)

if not exist "%~2" (
    color 0C
    echo [エラー] ファイルが見つかりません:
    echo          %~2
    goto ERROR_EXIT
)

:: 拡張子確認（複数形式対応）
set EXT1=%~x1
set EXT2=%~x2

call :CHECK_EXTENSION "%EXT1%" EXT1_OK
if "%EXT1_OK%"=="0" (
    color 0C
    echo [エラー] Excel1は対応形式ではありません
    echo          対応形式: .xlsx, .xls, .xlsm, .xlsb
    echo          指定されたファイル: %~nx1
    goto ERROR_EXIT
)

call :CHECK_EXTENSION "%EXT2%" EXT2_OK
if "%EXT2_OK%"=="0" (
    color 0C
    echo [エラー] Excel2は対応形式ではありません
    echo          対応形式: .xlsx, .xls, .xlsm, .xlsb
    echo          指定されたファイル: %~nx2
    goto ERROR_EXIT
)

:: エンジンファイル確認
if not exist "%ENGINE_FILE%" (
    :: 親ディレクトリも確認
    set ENGINE_FILE=%~dp0ExcelMergeEngine.xlsm
    if not exist "!ENGINE_FILE!" (
        color 0C
        echo [エラー] 処理エンジンが見つかりません
        echo          確認してください: ExcelMergeEngine.xlsm
        goto ERROR_EXIT
    )
)

:: 処理実行表示
echo 入力ファイル:
echo   Excel1: %~nx1
echo   Excel2: %~nx2
echo.
echo 処理を開始します...
echo.

:: VBScriptでExcelマクロ実行
set TEMP_VBS=%TEMP%\ExcelMerge_%RANDOM%.vbs

:: VBScriptファイル作成
(
    echo ' Excel結合処理VBScript
    echo ' Version: 2.0
    echo Option Explicit
    echo.
    echo Dim objExcel, objWorkbook
    echo Dim strEngine, strFile1, strFile2
    echo Dim bVisible
    echo.
    echo ' 引数取得
    echo strEngine = "%ENGINE_FILE%"
    echo strFile1 = "%~f1"
    echo strFile2 = "%~f2"
    echo bVisible = False
    echo.
    echo ' Excel起動
    echo On Error Resume Next
    echo Set objExcel = CreateObject^("Excel.Application"^)
    echo If Err.Number ^<^> 0 Then
    echo     WScript.Echo "エラー: Excelを起動できません"
    echo     WScript.Quit 1
    echo End If
    echo On Error GoTo 0
    echo.
    echo objExcel.Visible = bVisible
    echo objExcel.DisplayAlerts = False
    echo.
    echo On Error Resume Next
    echo.
    echo ' エンジンファイルを開く
    echo Set objWorkbook = objExcel.Workbooks.Open^(strEngine^)
    echo.
    echo If Err.Number ^<^> 0 Then
    echo     WScript.Echo "エラー: エンジンファイルを開けません - " ^& Err.Description
    echo     objExcel.Quit
    echo     WScript.Quit 1
    echo End If
    echo.
    echo ' マクロ実行
    echo objExcel.Run "ExecuteMerge", strFile1, strFile2
    echo.
    echo If Err.Number ^<^> 0 Then
    echo     WScript.Echo "エラー: マクロ実行エラー - " ^& Err.Description
    echo     On Error Resume Next
    echo     objWorkbook.Close False
    echo     objExcel.Quit
    echo     WScript.Quit 1
    echo End If
    echo.
    echo ' クリーンアップ
    echo On Error Resume Next
    echo objWorkbook.Close False
    echo objExcel.Quit
    echo.
    echo Set objWorkbook = Nothing
    echo Set objExcel = Nothing
    echo.
    echo WScript.Quit 0
) > "%TEMP_VBS%"

:: VBScript実行
cscript //nologo "%TEMP_VBS%"
set RESULT=%ERRORLEVEL%

:: 一時ファイル削除
del "%TEMP_VBS%" 2>nul

:: 結果確認
if %RESULT% equ 0 (
    color 0A
    echo.
    echo ============================================
    echo [成功] 処理が完了しました
    echo.
    echo 出力フォルダ: %OUTPUT_DIR%
    echo ============================================
) else (
    color 0C
    echo.
    echo ============================================
    echo [エラー] 処理中にエラーが発生しました
    echo.
    echo ログフォルダを確認してください:
    echo %LOG_DIR%
    echo ============================================
)

echo.
pause
exit /b %RESULT%

:USAGE
echo 使用方法:
echo.
echo   1. 2つのExcelファイルを選択
echo   2. このバッチファイルにドラッグ＆ドロップ
echo.
echo   または、コマンドラインから実行:
echo   %~n0 Excel1.xlsx Excel2.xlsx
echo.
echo 対応形式:
echo   .xlsx, .xls, .xlsm, .xlsb
echo.
echo 注意事項:
echo   - 2つのファイルを同時に指定してください
echo   - ファイルパスに特殊文字が含まれないようにしてください
echo.
pause
exit /b 1

:ERROR_EXIT
echo.
pause
exit /b 1

:CHECK_EXTENSION
:: 拡張子チェックサブルーチン
set "ext=%~1"
set "%2=0"

if /i "%ext%"==".xlsx" set "%2=1"
if /i "%ext%"==".xls" set "%2=1"
if /i "%ext%"==".xlsm" set "%2=1"
if /i "%ext%"==".xlsb" set "%2=1"

goto :eof
