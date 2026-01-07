'==================================================
' modConstants - 定数定義モジュール
' システム全体で使用する定数を一元管理
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' ログレベル定数
'--------------------------------------------------
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARNING As String = "WARNING"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_DEBUG As String = "DEBUG"

'--------------------------------------------------
' デフォルト設定値
'--------------------------------------------------
Public Const DEFAULT_EXCEL1_HEADER_ROWS As Long = 3
Public Const DEFAULT_EXCEL1_DATA_START_ROW As Long = 4
Public Const DEFAULT_EXCEL1_ID_COLUMN As String = "B"

Public Const DEFAULT_EXCEL2_HEADER_ROWS As Long = 2
Public Const DEFAULT_EXCEL2_DATA_START_ROW As Long = 3
Public Const DEFAULT_EXCEL2_ID_COLUMN As String = "A"

Public Const DEFAULT_OUTPUT_FILENAME_FORMAT As String = "結合データ_[DATE].xlsx"
Public Const DEFAULT_INCLUDE_LOG_SHEET As Boolean = True

'--------------------------------------------------
' 表示・処理制限
'--------------------------------------------------
Public Const MAX_DISPLAY_IDS As Long = 20           ' ログシートに表示する識別コードの最大件数
Public Const MAX_PROCESSING_ROWS As Long = 100000   ' 最大処理行数

'--------------------------------------------------
' タイムスタンプフォーマット
'--------------------------------------------------
Public Const TIMESTAMP_FORMAT_FULL As String = "yyyy/mm/dd hh:mm:ss"
Public Const TIMESTAMP_FORMAT_TIME As String = "hh:mm:ss"
Public Const TIMESTAMP_FORMAT_FILE As String = "yyyymmdd_hhmmss"

'--------------------------------------------------
' 設定ファイル関連
'--------------------------------------------------
Public Const CONFIG_FOLDER_NAME As String = "Config"
Public Const CONFIG_FILE_NAME As String = "MergeConfig.xlsx"
Public Const CONFIG_SHEET_NAME As String = "Config"

Public Const OUTPUT_FOLDER_NAME As String = "Output"
Public Const LOGS_FOLDER_NAME As String = "Logs"

'--------------------------------------------------
' 設定項目キー名
'--------------------------------------------------
Public Const CFG_EXCEL1_HEADER_ROWS As String = "Excel1_HeaderRows"
Public Const CFG_EXCEL1_DATA_START_ROW As String = "Excel1_DataStartRow"
Public Const CFG_EXCEL1_ID_COLUMN As String = "Excel1_IDColumn"

Public Const CFG_EXCEL2_HEADER_ROWS As String = "Excel2_HeaderRows"
Public Const CFG_EXCEL2_DATA_START_ROW As String = "Excel2_DataStartRow"
Public Const CFG_EXCEL2_ID_COLUMN As String = "Excel2_IDColumn"

Public Const CFG_OUTPUT_FILENAME_FORMAT As String = "Output_FileNameFormat"
Public Const CFG_INCLUDE_LOG_SHEET As String = "Output_IncludeLogSheet"
Public Const CFG_MAX_ROWS As String = "Processing_MaxRows"

'--------------------------------------------------
' 対応ファイル拡張子
'--------------------------------------------------
Public Const SUPPORTED_EXTENSIONS As String = ".xlsx,.xls,.xlsm,.xlsb"

'--------------------------------------------------
' エラーコード
'--------------------------------------------------
Public Const ERR_FILE_NOT_FOUND As String = "E001"
Public Const ERR_FILE_ACCESS As String = "E002"
Public Const ERR_COLUMN_NOT_FOUND As String = "E003"
Public Const ERR_MEMORY_INSUFFICIENT As String = "E004"
Public Const ERR_OUTPUT_ACCESS As String = "E005"
Public Const ERR_INVALID_FORMAT As String = "E006"
Public Const ERR_SAME_FILE As String = "E007"

'--------------------------------------------------
' UI関連
'--------------------------------------------------
Public Const APP_TITLE As String = "Excel結合処理システム"
Public Const APP_VERSION As String = "2.0"

'--------------------------------------------------
' 色定数
'--------------------------------------------------
Public Const COLOR_HEADER_BG As Long = 14277081     ' RGB(217, 217, 217)
Public Const COLOR_ERROR_TEXT As Long = 255         ' RGB(255, 0, 0)
Public Const COLOR_WARNING_TEXT As Long = 36095     ' RGB(255, 140, 0)
