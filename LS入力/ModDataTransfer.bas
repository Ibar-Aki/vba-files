'===============================================================================
' モジュール名: ModDataTransfer
'
' 【概要】      「データ登録」シートから「月次データ」シートへ、日々の作業時間データを
'               転記・集計するための機能を提供します。
' 【作成】      「JJ-07」2025/08
' 【対象環境】  Excel 2016+ / Windows
' 【主要機能】
' ・指定した日付の作業データを、「作業コード」と「作番」の組み合わせで集計
' ・集計結果を「月次データ」シートの対応する日付行・作業列に転記
' ・クリップボードに作業データをコピー
'===============================================================================
Option Explicit

'===============================================================================
' 【WinAPI宣言セクション】
' クリップボード操作用のWindows API関数を定義します。
' ※VBA7 (Office 2010以降の64bit版) とそれ以前の32bit版の両方に対応
'===============================================================================
#If VBA7 Then
    ' --- 64ビット版Office用API宣言 ---
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
#Else
    ' --- 32ビット版Office用API宣言 ---
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If

' --- WinAPI関連定数 ---
Private Const GMEM_MOVEABLE As Long = &H2           ' メモリブロックが移動可能であることを示すフラグ
Private Const CF_UNICODETEXT As Long = 13           ' クリップボード形式：Unicode文字列

'===============================================================================
' 【定数宣言セクション】
' モジュール全体の動作を制御する定数を定義します。
' ※シート構成やレイアウトが変更された場合は、このセクションを修正してください。
'===============================================================================

' --- シート名定数 ---
Private Const DATA_SHEET_NAME        As String = "データ登録"    ' 転記元となるデータ入力シート名
Private Const MONTHLY_SHEET_NAME     As String = "月次データ"    ' 転記先となる月次集計シート名

' --- 重要セル位置定数 ---
Private Const DATE_CELL_PRIORITY As String = "D4"  ' 転記日付セル（最優先でこのセルの日付を使用）
Private Const DATE_CELL_NORMAL   As String = "D3"  ' 転記日付セル（D4が空の場合に使用）
 ' エラーメッセージ表示セルは共通定数 ERR_CELL_ADDR（ModAppConfig.bas）を使用

' --- 行・列番号定数（データ登録シート）---
Private Const DATA_START_ROW     As Long = 8       ' データ入力範囲の開始行
Private Const COL_WORKNO   As Long = 3   ' C列：作番
Private Const COL_CATEGORY As Long = 4   ' D列：作業コード
Private Const COL_TIME     As Long = 5   ' E列：作業時間

' --- 行・列番号定数（月次データシート）---
Private Const COL_DATE     As Long = 2   ' B列：日付列
Private Const MONTHLY_WORKNO_ROW As Long = 10       ' 作番が記載されている行
Private Const MONTHLY_HEADER_ROW As Long = 11      ' 作業コードが記載されている行
Private Const MONTHLY_DATA_START_ROW As Long = 12  ' データ記録範囲の開始行
Private Const MONTHLY_MIN_COL    As Long = 3       ' データ記録範囲の最小列（C列）

' --- 時間計算関連定数 ---
Private Const MINUTES_PER_HOUR   As Double = 60#      ' 1時間の分数 (60分)
Private Const MINUTES_PER_DAY    As Double = 1440#    ' 1日の分数 (24時間 * 60分)
Private Const MAX_MINUTES_PER_HOUR As Long = 60       ' 1時間あたりの最大分数（時刻形式の妥当性チェック用）

' --- 動作設定・書式定数 ---
Private Const KEY_SEPARATOR As String = "|"              ' 内部処理で「作業コード」と「作番」を連結する際の区切り文字
Private Const TIME_FORMAT As String = "[hh]:mm"          ' Excelセルに設定する時間書式（24時間以上表示対応）
Private Const DATE_FORMAT As String = "yyyy/mm/dd(aaa)"  ' メッセージ表示用の日付書式
Private Const PREVIEW_TAB As String = vbTab              ' 確認ダイアログのプレビュー表示で使用するタブ文字
Private Const DUP_HIGHLIGHT_COLOR As Long = vbYellow       ' 重複データを検知した際にセルを塗りつぶす色

' --- 列追加ポリシー定数（新規列の追加方法を制御）---
Private Const AddPolicy_Prompt As Long = 0  ' ユーザーに確認してから列を追加
Private Const AddPolicy_Auto   As Long = 1  ' 確認なしで自動的に列を追加
Private Const AddPolicy_Reject As Long = 2  ' 新規列の追加を許可しない

' --- 開発・デバッグ用定数 ---
Private Const AUTO_ADD_POLICY As Long = AddPolicy_Prompt   ' 通常運用時の列追加ポリシー
Private Const DRY_RUN         As Boolean = False           ' Trueにすると、実際の書き込みを行わないテストモードで実行

'===============================================================================
' 【カスタムエラー定数セクション】
' このモジュール固有のエラーコードを定義します。
' ※VBA標準エラーとの衝突を避けるため、vbObjectErrorに数値を加算して使用
'===============================================================================
Private Const ERR_SHEET_NOT_FOUND   As Long = vbObjectError + 1  ' 指定されたシートが見つからないエラー
Private Const ERR_INVALID_DATE      As Long = vbObjectError + 2  ' 有効な日付が取得できないエラー
Private Const ERR_NO_DATA           As Long = vbObjectError + 3  ' 転記対象のデータが1件もないエラー
Private Const ERR_DATE_NOT_FOUND    As Long = vbObjectError + 4  ' 転記先日付が月次シートに見つからないエラー
Private Const ERR_PROTECTION_FAILED As Long = vbObjectError + 5  ' シート保護の解除に失敗したエラー

'===============================================================================
' 【データ構造（Type宣言セクション）】
' 処理に必要な情報をまとめて管理するためのカスタムデータ型を定義します。
'===============================================================================

' --- アプリケーション状態保存用 ---
' ※マクロ実行前のExcel設定を保存し、処理後に復元するために使用
Private Type ApplicationState
    ScreenUpdating As Boolean    ' 画面更新の状態 (True/False)
    EnableEvents   As Boolean    ' イベント発生の状態 (True/False)
    Calculation    As Long       ' 計算モードの状態 (xlCalculationAutomatic など)
End Type

' --- シート保護情報保存用 ---
' ※シート保護を一時的に解除し、処理後に元の状態に戻すために使用
Private Type SheetProtectionInfo
    IsProtected As Boolean       ' 元の保護状態 (True/False)
    Password    As String        ' 解除に使用したパスワード
End Type

' --- 転記処理設定用 ---
' ※転記処理全体で共有する設定情報を格納
Private Type TransferConfig
    targetDate     As Date       ' 転記対象の日付
    targetRow      As Long       ' 月次シート上の転記対象行番号
    DryRun         As Boolean    ' ドライランモード (True/False)
    AddPolicy      As Long       ' 新規列の追加ポリシー
End Type

' --- 処理結果情報用 ---
' ※処理完了後にユーザーへ表示する結果サマリーを格納
Private Type ProcessResult
    ProcessedCount  As Long      ' 正常に処理された件数
    DuplicateCount  As Long      ' 重複が検知された件数
    ErrorCount      As Long      ' エラーが発生した件数
    NewColumnsAdded As Long      ' 新規に追加された列数
    Messages        As String    ' ユーザーへの通知メッセージ
    Success         As Boolean   ' 処理全体の成功フラグ (True/False)
End Type

'===============================================================================
' 【メイン処理】
' このモジュールのエントリーポイント。ユーザーが直接実行するプロシージャです。
'===============================================================================
Public Sub TransferDataToMonthlySheet()
    ' --- 変数宣言 ---
    Dim prevState As ApplicationState         ' Excelアプリケーションの実行前状態
    Dim config As TransferConfig              ' 転記処理の各種設定
    Dim result As ProcessResult               ' 転記処理の結果
    Dim protectionInfo As SheetProtectionInfo ' 月次シートの保護情報
    Dim wsData As Worksheet                   ' 「データ登録」シートオブジェクト
    Dim wsMonthly As Worksheet                ' 「月次データ」シートオブジェクト

    ' --- ステップ1：Excel状態の保存と高速化設定 ---
    ' ※処理中のパフォーマンス向上のため、画面更新や自動計算を一時的に停止
    SaveAndSetApplicationState prevState

    ' --- エラーハンドリング設定 ---
    ' ※以降でエラーが発生した場合は「ErrorHandler」セクションにジャンプ
    On Error GoTo ErrorHandler

    ' --- 事前準備：前回のエラーメッセージをクリア ---
    ClearErrorCellOnMonthlySheet

    ' --- ステップ2：初期化と事前検証 ---
    ' ※シートの存在確認、日付の取得、転記先行の特定などを行う
    If Not InitializeTransferConfig(config, protectionInfo, wsData, wsMonthly) Then
        GoTo CleanUp ' 初期化に失敗した場合は、後処理へスキップ
    End If

    ' --- ステップ3：メインのデータ転記処理を実行 ---
    ExecuteDataTransfer config, wsData, wsMonthly, result

    ' --- ステップ4：処理結果をダイアログで表示 ---
    ShowTransferResults result

CleanUp:
    ' --- 最終処理：シート保護とExcel状態を元に戻す ---
    ' ※エラー発生時も必ず実行される
    RestoreSheetProtection wsMonthly, protectionInfo
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    ' --- エラー発生時の処理 ---
    Dim emsg As String
    ' エラー情報を分かりやすいメッセージに変換
    emsg = GetErrorDetails(Err.Number, Err.description)
    ' 月次シートの指定セルにエラーメッセージを表示
    ReportErrorToMonthlySheet emsg
    ' ★★★ 変更点：エラー内容をメッセージボックスでも表示 ★★★
    MsgBox emsg, vbCritical, "エラー"
    ' 最終処理へ
    Resume CleanUp
End Sub

'===============================================================================
' 【初期化・設定プロシージャ群】
' メイン処理の実行に必要な準備と検証を行います。
'===============================================================================

'===============================================================================
' 【機能名】転記処理の初期化と設定
' 【概要】  データ転記に必要な各種設定（シートオブジェクト、対象日付、対象行など）を
'           初期化し、事前検証を行います。
' 【引数】  config: 初期化された設定を格納するTransferConfig構造体 (出力)
'           protInfo: シート保護情報を格納するSheetProtectionInfo構造体 (出力)
'           wsData: 「データ登録」シートオブジェクト (出力)
'           wsMonthly: 「月次データ」シートオブジェクト (出力)
' 【戻り値】Boolean: 初期化に成功した場合はTrue、失敗した場合はFalse
'===============================================================================
Private Function InitializeTransferConfig( _
    ByRef config As TransferConfig, _
    ByRef protInfo As SheetProtectionInfo, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet) As Boolean

    ' --- 戻り値の初期化 ---
    InitializeTransferConfig = False

    ' --- ステップ1：ワークシートの取得と存在検証 ---
    If Not GetAndValidateWorksheets(wsData, wsMonthly) Then Exit Function

    ' --- ステップ2：月次シートの保護を一時的に解除 ---
    ' ※保護されている場合、必要に応じてパスワードの入力を求めます
    If Not UnprotectSheetIfNeeded(wsMonthly, protInfo) Then Exit Function

    ' --- ステップ3：データ登録シートから転記対象の日付を決定 ---
    If Not DetermineTargetDate(wsData, config.targetDate) Then Exit Function

    ' --- ステップ4：月次シートから対象日付と一致する行を検索 ---
    config.targetRow = FindMatchingDateRow(wsMonthly, config.targetDate)
    If config.targetRow = 0 Then
        ' 一致する日付が見つからない場合はカスタムエラーを発生
        RaiseCustomError ERR_DATE_NOT_FOUND, Format$(config.targetDate, DATE_FORMAT)
        Exit Function
    End If

    ' --- ステップ5：定数から動作設定を読み込み ---
    config.DryRun = DRY_RUN
    config.AddPolicy = AUTO_ADD_POLICY

    ' --- 全ての初期化処理が成功 ---
    InitializeTransferConfig = True
End Function

'===============================================================================
' 【機能名】ワークシートの取得と検証
' 【概要】  処理に必要な「データ登録」「月次データ」シートを取得し、
'           存在するかどうかを確認します。
' 【引数】  wsData: 「データ登録」シートオブジェクト (出力)
'           wsMonthly: 「月次データ」シートオブジェクト (出力)
' 【戻り値】Boolean: 両方のシートが正常に取得できた場合はTrue
'===============================================================================
Private Function GetAndValidateWorksheets(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    On Error GoTo SheetError

    ' --- 定義されたシート名でシートオブジェクトを取得 ---
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)

    ' --- シートの基本構造を検証 ---
    If Not ValidateSheetStructure(wsData, wsMonthly) Then
        GetAndValidateWorksheets = False
        Exit Function
    End If

    ' --- 正常終了 ---
    GetAndValidateWorksheets = True
    Exit Function

SheetError:
    ' --- シート取得エラー時の処理 ---
    ' ※存在しないシート名が指定された場合にこのエラーが発生
    RaiseCustomError ERR_SHEET_NOT_FOUND, "シート: " & DATA_SHEET_NAME & ", " & MONTHLY_SHEET_NAME
    GetAndValidateWorksheets = False
End Function

'===============================================================================
' 【機能名】シート構造の検証
' 【概要】  各シートが処理の前提となる最小限の構造を持っているかチェックします。
' 【引数】  wsData: 「データ登録」シートオブジェクト
'           wsMonthly: 「月次データ」シートオブジェクト
' 【戻り値】Boolean: 構造が有効な場合はTrue
'===============================================================================
Private Function ValidateSheetStructure(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    ' --- 戻り値の初期化 ---
    ValidateSheetStructure = False
    
    ' --- 「データ登録」シートの構造チェック ---
    ' ※データ開始行より下に何らかの作番データが存在するかを確認
    If wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row < DATA_START_ROW Then
        Exit Function
    End If

    ' --- 「月次データ」シートの構造チェック ---
    ' ※ヘッダ行に最小列（C列）までのデータが存在するかを確認
    If wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column < MONTHLY_MIN_COL Then
        Exit Function
    End If

    ' --- 検証成功 ---
    ValidateSheetStructure = True
End Function


'===============================================================================
' 【機能名】対象日付の決定
' 【概要】  「データ登録」シートから転記対象日を取得します。
'           優先セル(D4)を最初に見に行き、空の場合は通常セル(D3)を見に行きます。
' 【引数】  wsData: 「データ登録」シートオブジェクト
'           targetDate: 取得した日付を格納する変数 (出力)
' 【戻り値】Boolean: 有効な日付が取得できた場合はTrue
'===============================================================================
Private Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    ' --- 戻り値の初期化 ---
    DetermineTargetDate = False

    ' --- 優先セル（D4）から日付取得を試行 ---
    If IsDate(wsData.Range(DATE_CELL_PRIORITY).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_PRIORITY).Value)
        DetermineTargetDate = True

    ' --- 優先セルが空の場合、通常セル（D3）から日付取得を試行 ---
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).Value)
        DetermineTargetDate = True

    ' --- 両方のセルから有効な日付が取得できない場合はエラー ---
    Else
        RaiseCustomError ERR_INVALID_DATE, "セル " & DATE_CELL_NORMAL & " または " & DATE_CELL_PRIORITY
    End If
End Function

'===============================================================================
' 【データ転記コアロジック群】
' データの収集、集計、書き込みに関する中心的な処理を行います。
'===============================================================================

'===============================================================================
' 【機能名】データ転記処理の実行
' 【概要】  データ収集、集計、プレビュー、書き込みという一連の転記処理を統括します。
' 【引数】  config: 転記処理の設定情報
'           wsData: 「データ登録」シートオブジェクト
'           wsMonthly: 「月次データ」シートオブジェクト
'           result: 処理結果を格納するProcessResult構造体 (出力)
'===============================================================================
Private Sub ExecuteDataTransfer( _
    ByRef config As TransferConfig, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet, _
    ByRef result As ProcessResult)

    ' --- 変数宣言 ---
    Dim items As collection              ' 収集データ：各要素は Array(WorkNo, Category, Minutes, RowIndex)
    Dim aggregated As Object             ' 集計データ：Scripting.Dictionary (Key="Category|WorkNo", Value=合計分数)
    Dim mapDict As Object                ' 列マッピング：Scripting.Dictionary (Key="Category|WorkNo", Value=列番号)
    Dim lastCol As Long                  ' 月次シートの最終列番号

    ' --- ステップ1：「データ登録」シートから有効な時間データを収集 ---
    Set items = CollectTimeDataFromSheet(wsData)
    If items.Count = 0 Then
        RaiseCustomError ERR_NO_DATA, "有効な時間データが見つかりません"
        Exit Sub
    End If

    ' --- ステップ2：同一キー（作業コード+作番）のデータを集計 ---
    Set aggregated = AggregateTimeData(items)

    ' --- ステップ3：「月次データ」シートの列構成を解析し、マッピング情報を構築 ---
    Set mapDict = CreateObject("Scripting.Dictionary")
    BuildColumnMapping wsMonthly, lastCol, mapDict

    ' --- ステップ4：集計結果をプレビュー表示し、ユーザーに実行確認を求める ---
    If Not ShowPreviewAndConfirm(config.targetDate, aggregated) Then
        result.Success = False ' ユーザーがキャンセルした場合
        Exit Sub
    End If

    ' --- ステップ5：ドライランモードの場合は、書き込みせずに終了 ---
    If config.DryRun Then
        result.Messages = "ドライラン完了（実際の書き込みは実行されませんでした）"
        result.Success = True
        Exit Sub
    End If

    ' --- ステップ6：収集したデータをクリップボードにコピー（ユーザーの再利用のため） ---
    CopyDataToClipboard items, wsData

    ' --- ステップ7：集計データを月次シートに書き込み ---
    ' ※この中で、重複チェックや列の新規作成も行われる
    WriteAggregatedDataToSheet config, wsMonthly, aggregated, mapDict, lastCol, result

    ' --- 処理成功フラグを設定 ---
    result.Success = True
End Sub

'===============================================================================
' 【機能名】時間データの収集
' 【概要】  「データ登録」シートをスキャンし、有効な（必須項目が入力されている）
'           作業時間データを収集します。
' 【引数】  wsData: 「データ登録」シートオブジェクト
' 【戻り値】Collection: 収集したデータのコレクション。
'           各要素は Array(作番, 作業コード, 分数, 元の行番号) の形式。
'===============================================================================
Private Function CollectTimeDataFromSheet(ByRef wsData As Worksheet) As collection
    ' --- 変数宣言 ---
    Dim col As New collection           ' 収集結果を格納するコレクション
    Dim lastRow As Long, r As Long      ' ループ用の行変数
    Dim workNo As String                ' 作番
    Dim category As String              ' 作業コード
    Dim minutes As Double               ' 作業時間（分数に変換後）
    Dim arr(1 To 4) As Variant          ' 1行分のデータを格納する配列

    ' --- データが入力されている最終行を取得 ---
    lastRow = wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row

    ' --- データ開始行から最終行までループ ---
    For r = DATA_START_ROW To lastRow
        ' --- 各列の値を取得し、不要な空白を除去 ---
        workNo = Trim$(CStr(wsData.Cells(r, COL_WORKNO).Value))
        category = Trim$(CStr(wsData.Cells(r, COL_CATEGORY).Value))
        minutes = ConvertToMinutesEx(wsData.Cells(r, COL_TIME).Value) ' 様々な時間形式を分数に統一

        ' --- 有効性チェック：作番、作業コードが入力され、時間が0より大きい場合のみ対象 ---
        If (workNo <> "") And (category <> "") And (minutes > 0) Then
            ' --- 配列にデータを格納 ---
            arr(1) = workNo
            arr(2) = category
            arr(3) = minutes
            arr(4) = r          ' 元の行番号を保持（クリップボードコピー用）

            ' --- コレクションに追加 ---
            col.Add arr
        End If
    Next

    ' --- 収集結果を返す ---
    Set CollectTimeDataFromSheet = col
End Function

'===============================================================================
' 【機能名】時間データの集計
' 【概要】  収集したデータリストを元に、同一の「作業コード＋作番」を持つデータを
'           合算（集計）します。
' 【引数】  items: 収集されたデータのコレクション
' 【戻り値】Object(Scripting.Dictionary): 集計結果のディクショナリ
'           (Key="作業コード|作番", Value=合計分数)
'===============================================================================
Private Function AggregateTimeData(ByRef items As collection) As Object
    ' --- 変数宣言 ---
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Long                       ' ループカウンタ
    Dim key As String                   ' ディクショナリのキー
    Dim v As Variant                    ' コレクションの各要素（配列）

    ' --- 収集したデータ件数分ループ ---
    For i = 1 To items.Count
        v = items(i) ' 配列 [WorkNo, Category, Minutes, RowIndex] を取得

        ' --- キーを「作業コード|作番」の形式で生成 ---
        key = CStr(v(2)) & KEY_SEPARATOR & CStr(v(1))

        ' --- キーの存在に応じて、分数を加算または新規追加 ---
        If dic.Exists(key) Then
            dic(key) = dic(key) + CDbl(v(3)) ' 既存キー：加算
        Else
            dic.Add key, CDbl(v(3))          ' 新規キー：追加
        End If
    Next

    ' --- 集計結果のディクショナリを返す ---
    Set AggregateTimeData = dic
End Function

'===============================================================================
' 【機能名】列マッピングの構築
' 【概要】  「月次データ」シートのヘッダを解析し、「作業コード＋作番」と
'           列番号の対応表（ディクショナリ）を作成します。
' 【引数】  wsMonthly: 「月次データ」シートオブジェクト
'           lastColOut: 最終列番号を格納する変数 (出力)
'           mapDict: 作成したマッピング情報を格納するディクショナリ (出力)
'===============================================================================
Private Sub BuildColumnMapping(ByRef wsMonthly As Worksheet, ByRef lastColOut As Long, ByRef mapDict As Object)
    ' --- 変数宣言 ---
    Dim lastCol As Long                 ' 月次シートの最終列番号
    Dim c As Long                       ' ループ用の列変数
    Dim categoryName As String          ' ヘッダから読み取った作業コード名
    Dim workNoName As String            ' ヘッダから読み取った作番名
    Dim key As String                   ' ディクショナリ用のキー

    ' --- データが入力されている最終列を取得 ---
    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    lastColOut = lastCol

    ' --- データ開始列から最終列までループ ---
    For c = MONTHLY_MIN_COL To lastCol
        ' --- ヘッダ情報（作業コードと作番）を取得 ---
categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).Value))
workNoName = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).Value))

        ' --- 有効な列のみマッピングに登録 ---
        ' ※作業コードが空の列は無視する
        If categoryName <> "" Then
            key = categoryName & KEY_SEPARATOR & workNoName ' キー生成

            If Not mapDict.Exists(key) Then
                mapDict.Add key, c ' ディクショナリにキーと列番号を登録
            End If
        End If
    Next
End Sub

'===============================================================================
' 【機能名】転記内容のプレビュー表示と確認
' 【概要】  集計結果を整形してメッセージボックスに表示し、ユーザーに転記を
'           実行するかどうかの最終確認を求めます。
' 【引数】  targetDate: 転記対象の日付
'           aggregatedData: 集計結果のディクショナリ
' 【戻り値】Boolean: ユーザーが「はい」を選択した場合はTrue、「いいえ」の場合はFalse
'===============================================================================
Private Function ShowPreviewAndConfirm(ByVal targetDate As Date, ByRef aggregatedData As Object) As Boolean
    ' --- 変数宣言 ---
    Dim msg As String                   ' ダイアログに表示するメッセージ文字列
    Dim key As Variant                  ' ディクショナリのキーを巡回するための変数
    Dim n As Long                       ' 表示件数カウンタ
    Dim MAX_LINES As Long: MAX_LINES = 50 ' プレビューで詳細表示する最大行数

    ' --- メッセージのヘッダ部分を作成 ---
    msg = "以下の内容で転記します。よろしいですか？" & vbCrLf & vbCrLf & _
          "対象日付: " & Format$(targetDate, DATE_FORMAT) & vbCrLf & _
          String(50, "-") & vbCrLf & _
          "作番" & PREVIEW_TAB & " | 作業コード" & " | 時間" & vbCrLf & _
          String(50, "-") & vbCrLf

    ' --- 集計データの内容をメッセージに追加 ---
    For Each key In aggregatedData.Keys
        n = n + 1
        If n <= MAX_LINES Then
            ' --- 最大表示行数までは詳細を表示 ---
            Dim parts() As String
            parts = Split(CStr(key), KEY_SEPARATOR) ' キーを「作業コード」と「作番」に分割
            If UBound(parts) >= 1 Then
                ' 作番、作業コード、時間をタブ区切りで追加
                msg = msg & parts(1) & PREVIEW_TAB & " | " & parts(0) & PREVIEW_TAB & _
                      " | " & MinutesToHHMMString(aggregatedData(key)) & vbCrLf
            End If
        Else
            ' --- 最大表示行数を超えた場合は、残り件数のみ表示 ---
            msg = msg & "...ほか " & (aggregatedData.Count - MAX_LINES) & " 件" & vbCrLf
            Exit For
        End If
    Next

    ' --- 確認ダイアログを表示し、ユーザーの選択結果を返す ---
    ShowPreviewAndConfirm = (MsgBox(msg, vbYesNo + vbQuestion, "転記内容の確認") = vbYes)
End Function


'===============================================================================
' 【機能名】集計データの書き込み
' 【概要】  集計されたデータを月次シートの適切なセルに書き込みます。
'           この中で、列の取得・新規作成や重複処理の呼び出しを行います。
' 【引数】  config: 転記処理の設定情報
'           wsMonthly: 「月次データ」シートオブジェクト
'           aggregatedData: 集計結果のディクショナリ
'           mapDict: 列マッピング情報
'           lastCol: 最終列番号（新規作成時に更新される可能性あり）
'           result: 処理結果のサマリー (出力)
'===============================================================================
Private Sub WriteAggregatedDataToSheet( _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef aggregatedData As Object, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult)

    ' --- 変数宣言 ---
    Dim key As Variant                  ' ディクショナリのキー
    Dim parts() As String               ' キーを分割した結果の配列 (0=作業コード, 1=作番)
    Dim targetCol As Long               ' 書き込み対象の列番号

    ' --- 結果カウンタの初期化 ---
    result.ProcessedCount = 0
    result.DuplicateCount = 0
    result.NewColumnsAdded = 0

    ' --- 集計データの全項目をループ処理 ---
    For Each key In aggregatedData.Keys
        parts = Split(CStr(key), KEY_SEPARATOR)
        If UBound(parts) >= 1 Then
            ' --- 書き込み対象の列を取得（存在しない場合は設定に応じて新規作成）---
            targetCol = GetOrCreateColumn(parts(0), parts(1), config, wsMonthly, mapDict, lastCol, result)

            ' --- 有効な列が取得できた場合のみ書き込み実行 ---
            If targetCol > 0 Then
                ' ★★★ 変更点: 引数に config.targetDate を追加 ★★★
                WriteTimeDataToCell wsMonthly, config.targetRow, targetCol, aggregatedData(key), result, config.targetDate
                result.ProcessedCount = result.ProcessedCount + 1
            End If
        End If
    Next
End Sub


'===============================================================================
' 【列管理プロシージャ群】
' 月次シートの列を取得、または新規に作成・設定する処理を行います。
'===============================================================================

'===============================================================================
' 【機能名】列の取得または新規作成
' 【概要】  指定された「作業コード＋作番」に対応する列番号を取得します。
'           存在しない場合は、設定されたポリシーに基づき新規作成を試みます。
' 【引数】  category: 作業コード
'           workNo: 作番
'           config: 転記処理の設定情報
'           wsMonthly: 「月次データ」シートオブジェクト
'           mapDict: 列マッピング情報（新規作成時に更新される）
'           lastCol: 最終列番号（新規作成時に更新される）
'           result: 処理結果サマリー（新規作成時に更新される）
' 【戻り値】Long: 対象の列番号。作成が拒否された場合は0を返す。
'===============================================================================
Private Function GetOrCreateColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult) As Long

    ' --- 変数宣言 ---
    Dim key As String: key = category & KEY_SEPARATOR & workNo
    Dim newCol As Long

    ' --- マッピングにキーが存在すれば、既存の列番号を返す ---
    If mapDict.Exists(key) Then
        GetOrCreateColumn = mapDict(key)
        Exit Function
    End If

    ' --- 既存列がない場合、列追加ポリシーに応じて処理を分岐 ---
    Select Case config.AddPolicy
        Case AddPolicy_Reject
            ' ポリシーが「拒否」の場合：0を返して処理をスキップ
            GetOrCreateColumn = 0

        Case AddPolicy_Auto
            ' ポリシーが「自動」の場合：確認なしで新規列を作成
            newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
            If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
            GetOrCreateColumn = newCol

        Case Else ' AddPolicy_Prompt (デフォルト)
            ' ポリシーが「確認」の場合：ユーザーに確認
            If ConfirmColumnCreation(category, workNo) Then
                newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
                If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
                GetOrCreateColumn = newCol
            Else
                GetOrCreateColumn = 0 ' ユーザーがキャンセル
            End If
    End Select
End Function


'===============================================================================
' 【機能名】新規列の作成
' 【概要】  月次シートの末尾に新しい列を追加し、ヘッダ情報と書式を設定します。
' 【引数】  category: 新しい作業コード
'           workNo: 新しい作番
'           wsMonthly: 「月次データ」シートオブジェクト
'           mapDict: 列マッピング情報（この中で更新）
'           lastCol: 最終列番号（この中で更新）
' 【戻り値】Long: 新しく作成された列の番号
'===============================================================================
Private Function CreateNewColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long) As Long

    ' --- 新しい列番号を決定（最終列の次）---
    Dim newCol As Long
    newCol = lastCol + 1

    ' --- ヘッダ情報（作業コードと作番）を設定 ---
wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).Value = category
wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).Value = workNo

    ' --- 既存の列から書式（列幅、色、配置など）をコピー ---
    ApplyColumnFormatting wsMonthly, newCol, IIf(lastCol >= MONTHLY_MIN_COL, lastCol, MONTHLY_MIN_COL)

    ' --- データ入力部分のセルに時間書式を設定 ---
    SetDataColumnFormat wsMonthly, newCol

    ' --- マッピング情報と最終列番号を更新 ---
    mapDict.Add category & KEY_SEPARATOR & workNo, newCol
    lastCol = newCol

    ' --- 作成した列番号を返す ---
    CreateNewColumn = newCol
End Function

'===============================================================================
' 【機能名】列書式の適用
' 【概要】  新規作成した列に、既存の列から書式（列幅、配置、色など）をコピーします。
' 【引数】  wsMonthly: 「月次データ」シートオブジェクト
'           newCol: 書式設定対象の新しい列番号
'           sourceCol: 書式のコピー元となる列番号
'===============================================================================
Private Sub ApplyColumnFormatting(ByRef wsMonthly As Worksheet, ByVal newCol As Long, ByVal sourceCol As Long)
    On Error Resume Next ' 書式設定でエラーが発生しても処理を続行

    ' --- 列幅をコピー ---
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(sourceCol).ColumnWidth

    ' --- ヘッダ行（作業コード）の書式をコピー ---
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).WrapText
    End With

    ' --- 作番行の書式をコピー ---
    With wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).WrapText
    End With

    On Error GoTo 0
End Sub

'===============================================================================
' 【機能名】データ列の書式設定
' 【概要】  新規作成した列のデータ入力範囲に、時間表示書式 ([hh]:mm) を適用します。
' 【引数】  wsMonthly: 「月次データ」シートオブジェクト
'           col: 書式設定対象の列番号
'===============================================================================
Private Sub SetDataColumnFormat(ByRef wsMonthly As Worksheet, ByVal col As Long)
    ' --- データ範囲の最終行を取得（データがない場合はデフォルト行数まで設定）---
    Dim lastRow As Long
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then lastRow = MONTHLY_DATA_START_ROW + 31 ' デフォルト31日分

    ' --- データ範囲全体に時間書式を適用 ---
    With wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, col), wsMonthly.Cells(lastRow, col))
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

'===============================================================================
' 【機能名】列作成の確認ダイアログ
' 【概要】  新しい列を追加する前に、ユーザーに実行の可否を確認します。
' 【引数】  category: 新しい作業コード
'           workNo: 新しい作番
' 【戻り値】Boolean: ユーザーが「はい」を選択した場合はTrue
'===============================================================================
Private Function ConfirmColumnCreation(ByVal category As String, ByVal workNo As String) As Boolean
    ConfirmColumnCreation = (MsgBox( _
        "作業コード【" & category & "】+作番【" & workNo & "】の列がありません。" & vbCrLf & _
        "月次データシートに新しい列を追加しますか？", _
        vbYesNo + vbQuestion, "列の追加確認") = vbYes)
End Function

'===============================================================================
' 【セル書き込み・重複処理プロシージャ群】
' 個別のセルへのデータ書き込みと、その際の重複処理を担当します。
'===============================================================================

'===============================================================================
' 【機能名】セルへの時間データ書き込み
' 【概要】  指定されたセルに時間データを書き込みます。
'           ※重要：既存値がある場合は重複とみなし、上書きします。
' 【引数】  wsMonthly: 「月次データ」シートオブジェクト
'           targetRow: 書き込み対象の行番号
'           targetCol: 書き込み対象の列番号
'           minutes: 書き込む時間（分数）
'           result: 処理結果サマリー（重複時に更新される）
'           targetDate: 転記対象の日付（重複メッセージ記録用）
'===============================================================================
Private Sub WriteTimeDataToCell( _
    ByRef wsMonthly As Worksheet, _
    ByVal targetRow As Long, ByVal targetCol As Long, _
    ByVal minutes As Double, _
    ByRef result As ProcessResult, _
    ByVal targetDate As Date)

    ' --- 変数宣言 ---
    Dim existingValue As Double         ' セルに既に入力されている値（Excelシリアル値）
    Dim newValue As Double              ' これから書き込む新しい値（Excelシリアル値）
    Dim isDup As Boolean                ' 重複フラグ (True/False)
    Dim dupMsg As String                ' 重複メッセージ

    ' --- 既存値のチェック ---
    existingValue = NzD(wsMonthly.Cells(targetRow, targetCol).Value, 0#) ' Nullやエラーを安全に0に変換
    newValue = MinutesToSerial(minutes) ' 分数をExcelのシリアル値に変換
    isDup = (existingValue <> 0#)       ' 0以外の値が既にあれば重複と判断

    ' --- 重複時の処理 ---
    If isDup Then
        result.DuplicateCount = result.DuplicateCount + 1 ' 重複カウンタをインクリメント

        ' セルを黄色でハイライト
        HighlightDuplicateCell wsMonthly.Cells(targetRow, targetCol)

        ' 記録する重複メッセージを生成
        dupMsg = "登録日: " & Format$(targetDate, DATE_FORMAT) & " | 既存値検出: [" & _
                 CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, targetCol).Value) & "|" & _
                 CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, targetCol).Value) & "] 旧=" & _
                 SerialToHHMMString(existingValue)
        
        ReportErrorToMonthlySheet dupMsg, True
    End If

    ' --- セルへの値書き込み（常に上書き）---
    With wsMonthly.Cells(targetRow, targetCol)
        .Value = newValue
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

'===============================================================================
' 【機能名】重複セルのハイライト
' 【概要】  重複が検知されたセルを、指定された色で塗りつぶします。
' 【引数】  cell: ハイライト対象のRangeオブジェクト
'===============================================================================
Private Sub HighlightDuplicateCell(ByRef cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .Color = DUP_HIGHLIGHT_COLOR ' 定数で定義された色（黄色）
    End With
End Sub

'===============================================================================
' 【クリップボード操作プロシージャ群】
' 収集したデータをユーザーが再利用しやすいようにクリップボードにコピーします。
'===============================================================================

'===============================================================================
' 【機能名】クリップボードへのデータコピー
' 【概要】  収集したデータをタブ区切りのテキスト形式に整形し、
'           クリップボードにコピーします。
' 【引数】  items: 収集されたデータのコレクション
'           wsData: 「データ登録」シートオブジェクト（元のテキスト形式取得用）
'===============================================================================
Private Sub CopyDataToClipboard(ByRef items As collection, ByRef wsData As Worksheet)
    ' --- 変数宣言 ---
    Dim sb As String                    ' クリップボードにコピーする文字列を構築するためのバッファ
    Dim i As Long                       ' ループカウンタ
    Dim v As Variant                    ' コレクションの各要素（配列）

    ' --- 収集したデータ件数分ループ ---
    For i = 1 To items.Count
        v = items(i) ' 配列 [WorkNo, Category, Minutes, RowIndex] を取得

        ' --- タブ区切り形式の文字列を生成 ---
        ' ※重要：Excelに貼り付けた際の体裁を整えるため、意図的にタブを挿入
        sb = sb & CStr(v(1)) & vbTab & CStr(v(2)) & vbTab & vbTab & _
                 CStr(wsData.Cells(CLng(v(4)), COL_TIME).text) & vbCrLf
    Next

    ' --- 文字列が生成された場合のみ、クリップボードにコピー ---
    If Len(sb) > 0 Then CopyTextToClipboardSafe sb
End Sub

'===============================================================================
' 【機能名】安全なクリップボードコピー
' 【概要】  クリップボードへのテキストコピーを試みます。
'           まず互換性の高い `Forms.DataObject` を使用し、失敗した場合は
'           より低レベルな `WinAPI` を使用するフォールバック方式を採ります。
' 【引数】  textToCopy: クリップボードにコピーする文字列
'===============================================================================
Private Sub CopyTextToClipboardSafe(ByVal textToCopy As String)
    On Error GoTo APIFallback ' DataObjectでのエラー発生時はAPIFallbackへ

    ' --- 方法1：Forms.DataObjectを使用 (参照設定不要だが、環境によっては失敗する) ---
    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject")
    dataObject.SetText textToCopy
    dataObject.PutInClipboard
    Exit Sub

APIFallback:
    ' --- 方法2：WinAPIを直接呼び出し (より確実な方法) ---
    CopyTextToClipboardWinAPI textToCopy
End Sub

'===============================================================================
' 【機能名】WinAPIによるクリップボードコピー
' 【概要】  Windows APIを直接呼び出して、Unicodeテキストをクリップボードにコピーします。
' 【引数】  textToCopy: クリップボードにコピーする文字列
'===============================================================================
Private Sub CopyTextToClipboardWinAPI(ByVal textToCopy As String)
    ' --- 変数宣言（64/32ビット両対応）---
#If VBA7 Then
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
#Else
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
#End If
    Dim bytesNeeded As Long             ' 確保するメモリサイズ
    Dim retryCount As Long              ' クリップボードオープンリトライ用カウンタ

    ' --- 空文字列の場合は何もしない ---
    If Len(textToCopy) = 0 Then Exit Sub

    ' --- Unicode文字列を格納するためのグローバルメモリを確保 ---
    bytesNeeded = (Len(textToCopy) + 1) * 2 ' Unicodeは1文字2バイト + Null終端文字
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, bytesNeeded)
    If hGlobalMemory = 0 Then Exit Sub ' メモリ確保失敗

    ' --- メモリをロックしてポインタを取得し、文字列をコピー ---
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        lstrcpyW lpGlobalMemory, StrPtr(textToCopy) ' Unicode文字列をメモリにコピー
        GlobalUnlock hGlobalMemory

        ' --- クリップボードのオープンを試行（他プロセスが使用中の場合があるためリトライ処理）---
        For retryCount = 1 To 5
            If OpenClipboard(0) <> 0 Then Exit For ' オープン成功
            DoEvents
        Next retryCount

        If retryCount <= 5 Then
            ' --- クリップボード操作 ---
            EmptyClipboard ' クリップボードを空にする
            If SetClipboardData(CF_UNICODETEXT, hGlobalMemory) = 0 Then
                ' データのセットに失敗した場合は、確保したメモリを解放
                GlobalFree hGlobalMemory
            End If
            CloseClipboard ' クリップボードを閉じる
        Else
            ' リトライしてもオープンできなかった場合はメモリを解放
            GlobalFree hGlobalMemory
        End If
    Else
        ' メモリロック失敗時はメモリ解放
        GlobalFree hGlobalMemory
    End If
End Sub
'===============================================================================
' 【ユーティリティ関数・プロシージャ群】
' 時間変換、検索、メッセージ処理など、モジュール内で共通して使用される
' 汎用的な機能を提供します。
'===============================================================================

'===============================================================================
' 【機能名】拡張時間変換（分数へ）
' 【概要】  様々な形式（Date型、シリアル値、"HHMM"形式の数値/文字列など）で
'           与えられた時間データを、全て「分数」に統一して変換します。
' 【引数】  timeValue: 変換元の時間データ (Variant型)
' 【戻り値】Double: 変換後の分数。変換不可の場合は0を返す。
'===============================================================================
Private Function ConvertToMinutesEx(ByVal timeValue As Variant) As Double
    ' --- 変数宣言 ---
    Dim s As String

    ' --- 戻り値の初期化 ---
    ConvertToMinutesEx = 0

    ' --- 空値チェック ---
    If IsEmpty(timeValue) Then Exit Function

    ' --- Date型（時刻データ）の場合 ---
    If IsDate(timeValue) Then
        ConvertToMinutesEx = CDbl(CDate(timeValue)) * MINUTES_PER_DAY
        Exit Function
    End If

    ' --- 数値型の場合 ---
    If IsNumeric(timeValue) Then
        If InStr(1, CStr(timeValue), ".") > 0 Then
            ' 小数点を含む場合 → Excelシリアル値として扱う
            ConvertToMinutesEx = CDbl(timeValue) * MINUTES_PER_DAY
        Else
            ' 整数のみの場合 → "HHMM"形式の整数として扱う (例: 130 -> 1時間30分)
            ConvertToMinutesEx = ParseHHMMInteger(CLng(timeValue))
        End If
        Exit Function
    End If

    ' --- 文字列型の場合 ---
    s = Trim$(CStr(timeValue))
    If InStr(s, ":") > 0 Then
        ' コロンを含む場合 → "H:MM"形式の文字列として扱う (例: "1:30")
        ConvertToMinutesEx = ParseHHMMString(s)
    ElseIf IsNumeric(s) Then
        ' 数値のみの文字列の場合 → "HHMM"形式の整数として扱う
        ConvertToMinutesEx = ParseHHMMInteger(CLng(Val(s)))
    End If
End Function

'===============================================================================
' 【機能名】HHMM形式整数の解析
' 【概要】  "HHMM"形式で表現された整数（例: 130, 1030）を分数に変換します。
' 【引数】  hhmmValue: HHMM形式の整数 (Long)
' 【戻り値】Double: 変換後の分数。
'           例: 130 -> 90, 1030 -> 630
'===============================================================================
Private Function ParseHHMMInteger(ByVal hhmmValue As Long) As Double
    ' --- 変数宣言 ---
    Dim hours As Long, minutes As Long  ' 時間と分
    Dim t As String                     ' 整数を文字列に変換後の一時変数

    ' --- 戻り値の初期化と事前チェック ---
    ParseHHMMInteger = 0
    If hhmmValue < 0 Then Exit Function ' 負数は無効

    ' --- 桁数に応じて時間と分を分離 ---
    t = CStr(hhmmValue)
    Select Case Len(t)
        Case 1, 2
            ' 1-2桁の場合: 全て「分」として扱う (例: 5 -> 5分, 30 -> 30分)
            minutes = hhmmValue: hours = 0
        Case 3, 4
            ' 3-4桁の場合: 下2桁を「分」、残りを「時間」として扱う (例: 130 -> 1時間30分)
            hours = CLng(Left$(t, Len(t) - 2))
            minutes = CLng(Right$(t, 2))
        Case Else
            ' 5桁以上は無効な形式とみなし、処理を終了
            Exit Function
    End Select

    ' --- 分の妥当性チェック（0～59の範囲）---
    If minutes >= 0 And minutes < MAX_MINUTES_PER_HOUR Then
        ParseHHMMInteger = hours * MINUTES_PER_HOUR + minutes
    End If
End Function

'===============================================================================
' 【機能名】H:MM形式文字列の解析
' 【概要】  コロン区切りの時間文字列（例: "1:30"）を分数に変換します。
' 【引数】  timeString: H:MM形式の文字列
' 【戻り値】Double: 変換後の分数。
'===============================================================================
Private Function ParseHHMMString(ByVal timeString As String) As Double
    ' --- 変数宣言 ---
    Dim parts() As String               ' コロンで分割した結果を格納する配列
    Dim h As Long, m As Long            ' 時間と分

    ' --- 戻り値の初期化 ---
    ParseHHMMString = 0

    ' --- コロンで文字列を分割 ---
    parts = Split(timeString, ":")
    If UBound(parts) = 1 Then ' 分割結果が2つ（時間と分）であるか確認
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
            h = CLng(parts(0)): m = CLng(parts(1))

            ' --- 分の妥当性チェック（0～59の範囲）---
            If m >= 0 And m < MAX_MINUTES_PER_HOUR Then
                ParseHHMMString = h * MINUTES_PER_HOUR + m
            End If
        End If
    End If
End Function

'===============================================================================
' 【機能名】分数からシリアル値への変換
' 【概要】  分数をExcel内部で使われる時間のシリアル値に変換します。
' 【引数】  totalMinutes: 変換元の分数
' 【戻り値】Double: Excelの時間シリアル値
'===============================================================================
Private Function MinutesToSerial(ByVal totalMinutes As Double) As Double
    MinutesToSerial = totalMinutes / MINUTES_PER_DAY
End Function

'===============================================================================
' 【機能名】分数からH:MM形式文字列への変換
' 【概要】  分数を人間が読みやすい "H:MM" 形式の文字列に変換します。
' 【引数】  totalMinutes: 変換元の分数
' 【戻り値】String: "H:MM" 形式の文字列（例: 90 -> "1:30"）
'===============================================================================
Private Function MinutesToHHMMString(ByVal totalMinutes As Double) As String
    ' --- 変数宣言 ---
    Dim h As Long, m As Long           ' 時間と分

    ' --- 0以下の場合は "0:00" を返す ---
    If totalMinutes <= 0 Then
        MinutesToHHMMString = "0:00": Exit Function
    End If

    ' --- 分数から時間と分を計算 ---
    h = Int(totalMinutes / MINUTES_PER_HOUR)
    m = Round(totalMinutes - h * MINUTES_PER_HOUR, 0) ' 端数処理のため四捨五入

    ' --- 分が60になった場合の繰り上がり処理 ---
    If m = MAX_MINUTES_PER_HOUR Then h = h + 1: m = 0

    ' --- 書式を整えて返却 ---
    MinutesToHHMMString = Format$(h, "0") & ":" & Format$(m, "00")
End Function

'===============================================================================
' 【機能名】シリアル値からH:MM形式文字列への変換
' 【概要】  Excelの時間シリアル値を "H:MM" 形式の文字列に変換します。
' 【引数】  serialValue: Excelの時間シリアル値
' 【戻り値】String: "H:MM" 形式の文字列
'===============================================================================
Private Function SerialToHHMMString(ByVal serialValue As Double) As String
    SerialToHHMMString = MinutesToHHMMString(serialValue * MINUTES_PER_DAY)
End Function

'===============================================================================
' 【機能名】日付一致行の検索
' 【概要】  月次シートの日付列(B列)から、指定された日付と一致する行を検索します。
' 【引数】  wsMonthly: 月次データシートオブジェクト
'           targetDate: 検索する日付
' 【戻り値】Long: 一致した行の番号。見つからない場合は0を返す。
'===============================================================================
Private Function FindMatchingDateRow(ByRef wsMonthly As Worksheet, ByVal targetDate As Date) As Long
    ' --- 変数宣言 ---
    Dim lastRow As Long, r As Long     ' ループ用の行変数
    Dim d As Date                      ' 各行から読み取った日付

    ' --- 戻り値の初期化 ---
    FindMatchingDateRow = 0

    ' --- 日付列の最終行を取得 ---
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then Exit Function ' データが存在しない場合

    ' --- データ開始行から最終行までループ ---
    For r = MONTHLY_DATA_START_ROW To lastRow
        If IsDate(wsMonthly.Cells(r, COL_DATE).Value) Then
            d = CDate(wsMonthly.Cells(r, COL_DATE).Value)
            ' ※重要：時間は無視し、日付部分のみで比較
            If Int(d) = Int(targetDate) Then
                FindMatchingDateRow = r: Exit Function ' 一致した行が見つかったら即座に終了
            End If
        End If
    Next
End Function

'===============================================================================
' 【機能名】Null値安全な数値変換 (NzD)
' 【概要】  Variant型の値を安全にDouble型に変換します。
'           Null、Empty、エラー値、空文字列などの場合は、指定されたデフォルト値を返します。
' 【引数】  value: 変換元のVariant値
'           defaultValue: 変換失敗時に返すデフォルト値（省略時: 0）
' 【戻り値】Double: 変換後の数値
'===============================================================================
Private Function NzD(ByVal value As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    On Error Resume Next ' 数値変換エラーを無視するため

    ' --- 各種の無効値をチェック ---
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or value = "" Then
        NzD = defaultValue
    ElseIf IsNumeric(value) Then
        NzD = CDbl(value)
    Else
        NzD = defaultValue
    End If

    On Error GoTo 0
End Function

'===============================================================================
' 【アプリケーション状態・保護管理プロシージャ群】
' マクロ実行中のExcelの状態（画面更新、計算方法など）やシート保護を
' 安全に管理するための機能を提供します。
'===============================================================================

'===============================================================================
' 【機能名】アプリケーション状態の保存と高速化設定
' 【概要】  現在のExcel設定を構造体に保存し、処理高速化のために設定を変更します。
' 【引数】  prevState: ApplicationState構造体（設定の保存先）
'===============================================================================
Private Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    ' --- 現在の状態を構造体に保存 ---
    With prevState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
    End With

    ' --- 処理中のパフォーマンス向上のため設定を変更 ---
    With Application
        .ScreenUpdating = False              ' 画面描画を停止
        .EnableEvents = False                ' イベント発生を停止
        .Calculation = xlCalculationManual   ' 計算を手動に
    End With
End Sub

'===============================================================================
' 【機能名】アプリケーション状態の復元
' 【概要】  保存しておいたExcel設定に復元します。
' 【引数】  prevState: 保存しておいた状態を持つApplicationState構造体
'===============================================================================
Private Sub RestoreApplicationState(ByRef prevState As ApplicationState)
     ' --- 保存された状態にアプリケーション設定を復元 ---
    With Application
        .Calculation = prevState.Calculation        ' 計算モードを復元
        .EnableEvents = prevState.EnableEvents      ' イベント設定を復元
        .ScreenUpdating = prevState.ScreenUpdating  ' 画面描画を最後に有効化
    End With
End Sub

'===============================================================================
' 【機能名】シート保護の解除（必要時）
' 【概要】  シートが保護されている場合、解除を試みます。まず空パスワードで試行し、
'           失敗した場合はユーザーに入力を求めます。
' 【引数】  ws: 対象ワークシート
'           protInfo: 保護情報を格納する構造体 (出力)
' 【戻り値】Boolean: 解除に成功した場合はTrue、失敗・キャンセルの場合はFalse
'===============================================================================
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    ' --- 現在の保護状態を記録 ---
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    ' --- 保護されていない場合は、何もしないで正常終了 ---
    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    ' --- 保護されている場合の解除処理 ---
    On Error Resume Next ' パスワード間違いのエラーをハンドルするため
    ws.Unprotect ""      ' まずは空パスワードで試行
    If Err.Number = 0 Then
        ' 空パスワードで解除成功
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    ' --- 空パスワード失敗時、ユーザーに入力を促す ---
    Err.Clear
    protInfo.Password = InputBox("シート【" & ws.Name & "】の保護パスワードを入力してください。", "保護解除")

    ' --- ユーザーがキャンセルした場合 ---
    If protInfo.Password = "" Then
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ' --- 入力されたパスワードで解除を試行 ---
    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0) ' エラーが発生しなければ成功
    On Error GoTo 0
End Function

'===============================================================================
' 【機能名】シート保護の復元
' 【概要】  処理開始前にシートが保護されていた場合、元の状態に再保護します。
' 【引数】  ws: 対象ワークシート
'           protInfo: 保存しておいた保護情報
'===============================================================================
Private Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    ' --- 処理前に保護されていた場合のみ再保護を実行 ---
    If protInfo.IsProtected Then
        On Error Resume Next
        ' ※重要：UserInterfaceOnly:=True を指定し、ユーザーの手操作のみを禁止
        ' これにより、後続のマクロからの操作は引き続き許可される
        ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

'===============================================================================
' 【エラーハンドリング・結果表示プロシージャ群】
' エラー処理と、ユーザーへの結果通知を担当します。
'===============================================================================

'===============================================================================
' 【機能名】月次シートのエラー表示クリア
' 【概要】  処理開始時に、月次シートのエラーメッセージ表示セルをクリアします。
'===============================================================================
Private Sub ClearErrorCellOnMonthlySheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    If Not ws Is Nothing Then
        ws.Range(ERR_CELL_ADDR).ClearContents
        ws.Range(ERR_CELL_ADDR).WrapText = True
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' 【機能名】月次シートへのエラー報告
' 【概要】  エラー発生時に、エラーメッセージを月次シートの指定セルに表示します。
'           追記モードを指定すると、既存メッセージに改行して追加します。
' 【引数】  message: 表示するエラーメッセージ文字列
'           append: Trueの場合、既存メッセージに追記する (省略時: False)
'===============================================================================
Private Sub ReportErrorToMonthlySheet(ByVal message As String, Optional ByVal append As Boolean = False)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    If Not ws Is Nothing Then
        With ws.Range(ERR_CELL_ADDR)
            If append And Len(.Value) > 0 Then
                .Value = .Value & vbLf & message
            Else
                .Value = message
            End If
            .WrapText = True ' 長いメッセージでも見やすいように自動折り返し
        End With
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' 【機能名】カスタムエラーの発生
' 【概要】  このモジュール固有のカスタムエラーを発生させます。
' 【引数】  errorCode: 発生させるエラーのコード
'           description: エラーの詳細説明
'===============================================================================
Private Sub RaiseCustomError(ByVal errorCode As Long, ByVal description As String)
    Err.Raise errorCode, "ModDataTransfer", description
End Sub

'===============================================================================
' 【機能名】エラー詳細情報の生成
' 【概要】  発生したエラー番号と説明を元に、ユーザーにとって分かりやすい
'           エラーメッセージの全文を生成します。
' 【引数】  errNumber: エラー番号 (Err.Number)
'           errDescription: エラーの説明 (Err.Description)
' 【戻り値】String: 生成された詳細なエラーメッセージ
'===============================================================================
Private Function GetErrorDetails(ByVal errNumber As Long, ByVal errDescription As String) As String
    Select Case errNumber
        Case ERR_SHEET_NOT_FOUND
            GetErrorDetails = "必要なシートが見つかりません: " & errDescription
        Case ERR_INVALID_DATE
            GetErrorDetails = "日付が無効です: " & errDescription
        Case ERR_NO_DATA
            GetErrorDetails = "転記するデータがありません: " & errDescription
        Case ERR_DATE_NOT_FOUND
            GetErrorDetails = "対象日付が月次シートに見つかりません: " & errDescription
        Case ERR_PROTECTION_FAILED
            GetErrorDetails = "シート保護の解除に失敗しました: " & errDescription
        Case 9 ' Subscript out of range
            GetErrorDetails = FriendlyErrorMessage9(errDescription)
        Case Else
            GetErrorDetails = "予期しないエラーが発生しました (エラー #" & errNumber & "): " & errDescription
    End Select
End Function

'===============================================================================
' 【機能名】エラー#9の詳細説明
' 【概要】  VBAで頻発するエラー#9（インデックスが有効範囲にありません）に対して、
'           原因と対策を具体的に示した、分かりやすいメッセージを生成します。
' 【引数】  errDesc: 元のエラー説明
' 【戻り値】String: 生成された詳細なエラーメッセージ
'===============================================================================
Private Function FriendlyErrorMessage9(ByVal errDesc As String) As String
    FriendlyErrorMessage9 = _
        "エラー #9（インデックスが有効範囲にありません）が発生しました。" & vbCrLf & _
        "これは、指定した項目が見つからない場合に発生する典型的なエラーです。" & vbCrLf & vbCrLf & _
        "【考えられる原因と対策】" & vbCrLf & _
        "1. シート名の確認：ブック内に「" & DATA_SHEET_NAME & "」と「" & MONTHLY_SHEET_NAME & "」という名前のシートが存在するか確認してください。" & vbCrLf & _
        "2. データ形式の確認：「データ登録」シートの作番と作業コードが正しく入力されているか確認してください。" & vbCrLf & _
        vbCrLf & "詳細情報: " & errDesc
End Function

'===============================================================================
' 【機能名】転記結果の表示
' 【概要】  処理完了後、成功・失敗に応じて結果のサマリーをメッセージボックスで表示します。
' 【引数】  result: 処理結果を格納したProcessResult構造体
'===============================================================================
Private Sub ShowTransferResults(ByRef result As ProcessResult)
    ' --- 変数宣言 ---
    Dim message As String

    If result.Success Then
        ' --- 成功時のメッセージを構成 ---
        message = "転記処理が完了しました。" & vbCrLf & vbCrLf & _
                  "処理件数: " & result.ProcessedCount & " 件" & vbCrLf

        ' 重複があった場合のみ情報を追加
        If result.DuplicateCount > 0 Then
            ' ERROR_CELL定数を参照するようにメッセージを修正
            message = message & "重複検知: " & result.DuplicateCount & " 件（詳細は月次データを確認）" & vbCrLf
        End If

        ' 新規列が追加された場合のみ情報を追加
        If result.NewColumnsAdded > 0 Then
            message = message & "新規列追加: " & result.NewColumnsAdded & " 列" & vbCrLf
        End If

        ' その他のメッセージがある場合に追加
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & "メッセージ:" & vbCrLf & result.Messages
        End If

        ' 情報ダイアログを表示
        MsgBox message, vbInformation, "転記完了"

    Else
        ' --- 中止・失敗時のメッセージを構成 ---
        message = "転記処理が中止されました。"
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & vbCrLf & result.Messages
        End If

        ' 警告ダイアログを表示
        MsgBox message, vbExclamation, "処理中止"
    End If
End Sub
