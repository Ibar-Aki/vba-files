Option Explicit
'===============================================================================
' モジュール名: TransferDataModule.bas
' 
' 【概要】「データ登録」シートから「月次データ」シートへの時間データ転記処理
' 【作成】「JJ-07」2025/08
' 【対象環境】Excel 2016+ / Windows
' 【主要機能】
' ・同一キーのデータは自動集計して転記
' ・既存値がある場合は上書き（加算はしない）
' ・重複時は元の値をメッセージ列に記録し、セルを黄色でハイライト
' ・作業コード＋作番の組み合わせで転記先列を特定
' ・存在しない作業コード＋作番の列は確認後に自動追加可能
' 【更新履歴】
' ・v1.0: 初版作成
' ==============================================================================

'=========================
' WinAPI宣言部（64/32ビット両対応）
' クリップボード操作用のWindows API関数
'=========================
#If VBA7 Then
    ' === 64ビット版Office用API宣言 ===
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
    ' === 32ビット版Office用API宣言 ===
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

' === WinAPI関連定数 ===
Private Const GMEM_MOVEABLE As Long = &H2           ' メモリブロック移動可能フラグ
Private Const CF_UNICODETEXT As Long = 13           ' Unicode文字列クリップボード形式

'=========================
' システム設定定数
' ※これらの定数を変更することで、シート構造に合わせてカスタマイズ可能
'=========================

' === シート名定数 ===
Private Const DATA_SHEET_NAME        As String = "データ登録"    ' 転記元シート名
Private Const MONTHLY_SHEET_NAME     As String = "月次データ"    ' 転記先シート名

' === 重要セル位置定数 ===
Private Const DATE_CELL_PRIORITY As String = "D4"  ' 優先日付セル（優先取得）
Private Const DATE_CELL_NORMAL   As String = "D3"  ' 通常日付セル（D4が空の場合）

' === 行・列番号定数 ===
Private Const DATA_START_ROW     As Long = 8       ' データ登録シートのデータ開始行
Private Const MONTHLY_WORKNO_ROW As Long = 8       ' 月次シートの作番行
Private Const MONTHLY_HEADER_ROW As Long = 9       ' 月次シートの作業コード行
Private Const MONTHLY_DATA_START_ROW As Long = 10  ' 月次シートのデータ開始行
Private Const MONTHLY_MIN_COL    As Long = 3       ' 月次シートの最小列（C列）

' === データ登録シートの列定数 ===
Private Const COL_MESSAGE  As Long = 1   ' A列：メッセージ列
Private Const COL_DATE     As Long = 2   ' B列：日付列
Private Const COL_WORKNO   As Long = 3   ' C列：作番列
Private Const COL_CATEGORY As Long = 4   ' D列：作業コード列
Private Const COL_TIME     As Long = 5   ' E列：時間列

' === 時間計算関連定数 ===
Private Const MINUTES_PER_HOUR   As Double = 60#      ' 1時間の分数
Private Const MINUTES_PER_DAY    As Double = 1440#    ' 1日の分数（24時間×60分）
Private Const MAX_MINUTES_PER_HOUR As Long = 60       ' 1時間内の最大分数（検証用）
Private Const MAX_RETRY_CLIPBOARD  As Long = 5        ' クリップボード操作リトライ回数
Private Const DEFAULT_PREVIEW_ROWS As Long = 31       ' デフォルト行数（書式設定用）

' === 文字列操作定数 ===
Private Const KEY_SEPARATOR As String = "|"              ' キー区切り文字（作業コード|作番）
Private Const MESSAGE_SEPARATOR As String = vbLf         ' メッセージ区切り文字
Private Const TIME_FORMAT As String = "[hh]mm"           ' 時間表示書式（24時間以上対応）
Private Const DATE_FORMAT As String = "yyyy/mm/dd(aaa)"  ' 日付表示書式
Private Const PREVIEW_TAB As String = vbTab              ' プレビュー用タブ文字

' === 列追加ポリシー定数 ===
Private Const AddPolicy_Prompt As Long = 0  ' 確認してから列追加
Private Const AddPolicy_Auto   As Long = 1  ' 自動で列追加
Private Const AddPolicy_Reject As Long = 2  ' 列追加を拒否

' === 動作設定定数 ===
Private Const AUTO_ADD_POLICY As Long = AddPolicy_Prompt    ' 列追加時の動作（通常は確認）
Private Const DRY_RUN         As Boolean = False           ' ドライラン実行フラグ
Private Const DUP_HIGHLIGHT_COLOR As Long = vbYellow       ' 重複時のハイライト色

'=========================
' カスタムエラー定数
' 独自エラーコードの定義（vbObjectError + 番号）
'=========================
Private Const ERR_SHEET_NOT_FOUND   As Long = vbObjectError + 1  ' シートが見つからない
Private Const ERR_INVALID_DATE      As Long = vbObjectError + 2  ' 無効な日付
Private Const ERR_NO_DATA           As Long = vbObjectError + 3  ' データなし
Private Const ERR_DATE_NOT_FOUND    As Long = vbObjectError + 4  ' 対象日付が見つからない
Private Const ERR_PROTECTION_FAILED As Long = vbObjectError + 5  ' シート保護解除失敗

'=========================
' データ構造（Type宣言セクション）
' 処理に必要な情報をまとめた構造体
'=========================

' === アプリケーション状態保存用 ===
Private Type ApplicationState
    ScreenUpdating As Boolean    ' 画面更新状態
    EnableEvents   As Boolean    ' イベント有効状態
    Calculation    As Long       ' 計算モード（XlCalculation列挙型を数値保持）
End Type

' === シート保護情報保存用 ===
Private Type SheetProtectionInfo
    IsProtected As Boolean       ' 保護状態
    Password    As String        ' パスワード
End Type

' === 転記処理設定用 ===
Private Type TransferConfig
    targetDate     As Date       ' 対象日付
    targetRow      As Long       ' 対象行番号
    DryRun         As Boolean    ' ドライラン実行フラグ
    AddPolicy      As Long       ' 列追加ポリシー
End Type

' === 処理結果情報用 ===
Private Type ProcessResult
    ProcessedCount  As Long      ' 処理件数
    DuplicateCount  As Long      ' 重複件数
    ErrorCount      As Long      ' エラー件数
    NewColumnsAdded As Long      ' 新規追加列数
    Messages        As String    ' 結果メッセージ
    Success         As Boolean   ' 成功フラグ
End Type

'===============================================================================
' 【メイン処理】
' 公開プロシージャ：ユーザーが実行する転記処理のエントリーポイント
'===============================================================================
Public Sub TransferDataToMonthlySheet()
    ' === 変数宣言 ===
    Dim prevState As ApplicationState         ' Excelアプリケーション状態保存
    Dim config As TransferConfig              ' 転記処理設定
    Dim result As ProcessResult               ' 処理結果
    Dim protectionInfo As SheetProtectionInfo ' シート保護情報
    Dim wsData As Worksheet                   ' データ登録シート
    Dim wsMonthly As Worksheet                ' 月次データシート

    ' === アプリケーション状態の保存と高速化設定 ===
    SaveAndSetApplicationState prevState
    
    ' === エラーハンドリング開始 ===
    On Error GoTo ErrorHandler
    
    ' === 月次シートのエラー表示セルをクリア ===
    ClearErrorCellOnMonthlySheet

    ' === 初期化と検証（シート取得含む） ===
    If Not InitializeTransferConfig(config, protectionInfo, wsData, wsMonthly) Then
        GoTo CleanUp  ' 初期化失敗時は終了処理へ
    End If

    ' === メインのデータ転記処理実行 ===
    ExecuteDataTransfer config, wsData, wsMonthly, result

    ' === 処理結果をユーザーに表示 ===
    ShowTransferResults result
    
CleanUp:
    ' === 後処理：シート保護復元とアプリケーション状態復元 ===
    RestoreSheetProtection wsMonthly, protectionInfo
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    ' === エラー発生時の処理 ===
    Dim emsg As String
    emsg = GetErrorDetails(Err.Number, Err.description)
    ReportErrorToMonthlySheet emsg  ' 月次シートにエラーメッセージを表示
    Resume CleanUp
End Sub

'===============================================================================
' 【初期化と設定】
' 転記処理に必要な各種設定と検証を実行
'===============================================================================
Private Function InitializeTransferConfig( _
    ByRef config As TransferConfig, _
    ByRef protInfo As SheetProtectionInfo, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet) As Boolean

    ' === 初期化：失敗時はFalseを返す ===
    InitializeTransferConfig = False

    ' === シート取得と基本検証 ===
    If Not GetAndValidateWorksheets(wsData, wsMonthly) Then Exit Function

    ' === シート保護の解除（必要に応じてパスワード入力） ===
    If Not UnprotectSheetIfNeeded(wsMonthly, protInfo) Then Exit Function

    ' === 対象日付の決定（優先セル→通常セルの順で検索） ===
    If Not DetermineTargetDate(wsData, config.targetDate) Then Exit Function

    ' === 月次シートから対象日付の行を検索 ===
    config.targetRow = FindMatchingDateRow(wsMonthly, config.targetDate)
    If config.targetRow = 0 Then
        RaiseCustomError ERR_DATE_NOT_FOUND, Format$(config.targetDate, DATE_FORMAT)
        Exit Function
    End If

    ' === その他の動作設定 ===
    config.DryRun = DRY_RUN            ' ドライラン設定
    config.AddPolicy = AUTO_ADD_POLICY  ' 列追加ポリシー設定

    ' === 初期化成功 ===
    InitializeTransferConfig = True
End Function

'===============================================================================
' 【シート取得と基本検証】
' 必要なシートの存在確認と基本構造の検証
'===============================================================================
Private Function GetAndValidateWorksheets(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    On Error GoTo SheetError

    ' === 指定名のシートを取得 ===
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)

    ' === シート構造の基本検証 ===
    If Not ValidateSheetStructure(wsData, wsMonthly) Then
        GetAndValidateWorksheets = False
        Exit Function
    End If

    ' === 正常終了 ===
    GetAndValidateWorksheets = True
    Exit Function

SheetError:
    ' === シート取得エラー時の処理 ===
    RaiseCustomError ERR_SHEET_NOT_FOUND, "シート: " & DATA_SHEET_NAME & ", " & MONTHLY_SHEET_NAME
    GetAndValidateWorksheets = False
End Function

'===============================================================================
' 【シート構造検証】
' 各シートが必要最小限の構造を持っているかチェック
'===============================================================================
Private Function ValidateSheetStructure(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    ' === データ登録シートの構造チェック ===
    ' 作番列にデータ開始行以降にデータがあるか確認
    If wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row < DATA_START_ROW Then
        ValidateSheetStructure = False
        Exit Function
    End If
    
    ' === 月次データシートの構造チェック ===
    ' ヘッダ行に最低限の列（C列まで）があるか確認
    If wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column < MONTHLY_MIN_COL Then
        ValidateSheetStructure = False
        Exit Function
    End If
    
    ' === 構造検証成功 ===
    ValidateSheetStructure = True
End Function

'===============================================================================
' 【対象日付決定】
' データ登録シートから転記対象となる日付を取得
' 優先セル（D4）→通常セル（D3）の順で確認
'===============================================================================
Private Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    DetermineTargetDate = False
    
    ' === 優先セル（D4）から日付取得を試行 ===
    If IsDate(wsData.Range(DATE_CELL_PRIORITY).value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_PRIORITY).value)
        DetermineTargetDate = True
        
    ' === 通常セル（D3）から日付取得を試行 ===
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).value)
        DetermineTargetDate = True
        
    ' === 両セルとも無効な場合はエラー ===
    Else
        RaiseCustomError ERR_INVALID_DATE, "セル " & DATE_CELL_NORMAL & " または " & DATE_CELL_PRIORITY
    End If
End Function

'===============================================================================
' 【データ転記処理の実行】
' メインの転記処理ロジック：データ収集→集計→書き込み
'===============================================================================
Private Sub ExecuteDataTransfer( _
    ByRef config As TransferConfig, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet, _
    ByRef result As ProcessResult)

    ' === 変数宣言 ===
    Dim items As collection              ' 収集データ：各行はArray(WorkNo, Category, Minutes, RowIndex)
    Dim aggregated As Object             ' 集計データ：Scripting.Dictionary (key="作業コード|作番", val=合計分数)
    Dim mapDict As Object                ' 列マッピング：Dictionary (key="作業コード|作番", val=列番号)
    Dim lastCol As Long                  ' 月次シートの最終列番号

    ' === ステップ1：データ登録シートからの時間データ収集 ===
    Set items = CollectTimeDataFromSheet(wsData)
    If items.Count = 0 Then
        RaiseCustomError ERR_NO_DATA, "有効な時間データが見つかりません"
        Exit Sub
    End If

    ' === ステップ2：同一キー（作業コード+作番）のデータを集計 ===
    Set aggregated = AggregateTimeData(items)

    ' === ステップ3：月次シートの列マッピング構築 ===
    Set mapDict = CreateObject("Scripting.Dictionary")
    BuildColumnMapping wsMonthly, lastCol, mapDict

    ' === ステップ4：転記内容のプレビュー表示と確認 ===
    If Not ShowPreviewAndConfirm(config.targetDate, aggregated) Then
        result.Success = False
        Exit Sub
    End If

    ' === ステップ5：ドライラン実行時の処理 ===
    If config.DryRun Then
        result.Messages = "ドライラン完了（実際の書き込みは実行されませんでした）"
        result.Success = True
        Exit Sub
    End If

    ' === ステップ6：データをクリップボードに表形式でコピー ===
    CopyDataToClipboard items, wsData

    ' === ステップ7：メッセージ列のヘッダを確認・設定 ===
    EnsureMessageColumnHeader wsMonthly

    ' === ステップ8：集計データを月次シートに書き込み（既存値は上書き） ===
    WriteAggregatedDataToSheet config, wsMonthly, aggregated, mapDict, lastCol, result

    ' === 処理成功フラグ設定 ===
    result.Success = True
End Sub

'===============================================================================
' 【データ収集】
' データ登録シートから有効な時間データを収集
' 戻り値：Collection（各要素は配列：WorkNo, Category, Minutes, RowIndex）
'===============================================================================
Private Function CollectTimeDataFromSheet(ByRef wsData As Worksheet) As collection
    ' === 変数宣言 ===
    Dim col As New collection           ' 収集結果格納用コレクション
    Dim lastRow As Long, r As Long      ' 行ループ変数
    Dim workNo As String                ' 作番
    Dim category As String              ' 作業コード
    Dim minutes As Double               ' 時間（分数）
    Dim arr(1 To 4) As Variant         ' データ格納配列：1=WorkNo, 2=Category, 3=Minutes, 4=RowIndex

    ' === データ範囲の最終行を取得 ===
    lastRow = wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row
    
    ' === 各行のデータをチェックして有効なもののみ収集 ===
    For r = DATA_START_ROW To lastRow
        ' 各列の値を取得・整形
        workNo = Trim$(CStr(wsData.Cells(r, COL_WORKNO).value))      ' 作番（前後空白除去）
        category = Trim$(CStr(wsData.Cells(r, COL_CATEGORY).value))  ' 作業コード（前後空白除去）
        minutes = ConvertToMinutesEx(wsData.Cells(r, COL_TIME).value) ' 時間を分数に変換
        
        ' === 有効性チェック：作番・作業コード・時間すべてに値があるもののみ ===
        If (workNo <> "") And (category <> "") And (minutes > 0) Then
            ' 配列にデータをセット
            arr(1) = workNo      ' 作番
            arr(2) = category    ' 作業コード
            arr(3) = minutes     ' 分数
            arr(4) = r          ' 元の行番号
            
            ' コレクションに追加
            col.Add arr
        End If
    Next
    
    ' === 収集結果を返却 ===
    Set CollectTimeDataFromSheet = col
End Function

'===============================================================================
' 【データ集計】
' 同一キー（作業コード+作番）のデータを合算
' 引数：items - 収集されたデータのコレクション
' 戻り値：Dictionary（key="作業コード|作番", value=合計分数）
'===============================================================================
Private Function AggregateTimeData(ByRef items As collection) As Object
    ' === 変数宣言 ===
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Long                       ' ループカウンタ
    Dim key As String                   ' 辞書のキー（作業コード|作番）
    Dim v As Variant                    ' コレクションの各要素

    ' === 各データを処理してキー別に集計 ===
    For i = 1 To items.Count
        v = items(i)  ' 配列を取得：[WorkNo, Category, Minutes, RowIndex]
        
        ' キー生成：「作業コード|作番」の形式
        key = CStr(v(2)) & KEY_SEPARATOR & CStr(v(1))  ' Category|WorkNo
        
        ' 既存キーなら加算、新規キーなら新規追加
        If dic.Exists(key) Then
            dic(key) = dic(key) + CDbl(v(3))  ' 分数を加算
        Else
            dic.Add key, CDbl(v(3))           ' 新規追加
        End If
    Next
    
    ' === 集計結果を返却 ===
    Set AggregateTimeData = dic
End Function

'===============================================================================
' 【列マッピング構築】
' 月次シートの既存列から「作業コード+作番」の組み合わせと列番号の対応を構築
'===============================================================================
Private Sub BuildColumnMapping(ByRef wsMonthly As Worksheet, ByRef lastColOut As Long, ByRef mapDict As Object)
    ' === 変数宣言 ===
    Dim lastCol As Long                 ' 月次シートの最終列
    Dim c As Long                       ' 列ループ変数
    Dim categoryName As String          ' 作業コード名
    Dim workNoName As String            ' 作番名
    Dim key As String                   ' マッピング用キー

    ' === 月次シートの最終列を取得 ===
    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    lastColOut = lastCol

    ' === 各列を走査してマッピングを構築 ===
    For c = MONTHLY_MIN_COL To lastCol
        ' ヘッダ行（作業コード）と作番行から値を取得
        categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).value))  ' 作業コード
        workNoName = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).value))    ' 作番
        
        ' === 作業コードが設定されている列のみマッピングに登録 ===
        If categoryName <> "" Then
            key = categoryName & KEY_SEPARATOR & workNoName  ' 「作業コード|作番」キー生成
            
            ' 重複チェックして未登録の場合のみ追加
            If Not mapDict.Exists(key) Then
                mapDict.Add key, c  ' キーと列番号を対応付け
            End If
        End If
    Next
End Sub

'===============================================================================
' 【転記内容プレビュー表示】
' ユーザーに転記内容を事前確認してもらう
' 戻り値：True=続行, False=中止
'===============================================================================
Private Function ShowPreviewAndConfirm(ByVal targetDate As Date, ByRef aggregatedData As Object) As Boolean
    ' === 変数宣言 ===
    Dim msg As String                   ' 表示メッセージ
    Dim key As Variant                  ' 辞書のキー
    Dim n As Long                       ' 表示件数カウンタ
    Dim MAX_LINES As Long              ' 最大表示行数
    MAX_LINES = 50                     ' プレビューで表示する最大行数

    ' === メッセージヘッダ作成 ===
    msg = "以下の内容で転記します。よろしいですか？" & vbCrLf & vbCrLf & _
          "対象日付: " & Format$(targetDate, DATE_FORMAT) & vbCrLf & _
          String(50, "-") & vbCrLf & _
          "作番" & PREVIEW_TAB & " | 作業ｺｰﾄﾞ" & " | 時間" & vbCrLf & _
          String(50, "-") & vbCrLf

    ' === 集計データの内容を表示用に整形 ===
    n = 0
    For Each key In aggregatedData.Keys
        n = n + 1
        
        ' 最大表示行数以内の場合のみ詳細表示
        If n <= MAX_LINES Then
            Dim parts() As String
            parts = Split(CStr(key), KEY_SEPARATOR)  ' 「作業コード|作番」を分割
            
            If UBound(parts) >= 1 Then
                ' 表形式で情報を追加（作番 | 作業コード | 時間）
                msg = msg & parts(1) & PREVIEW_TAB & " | " & parts(0) & PREVIEW_TAB & _
                      " | " & MinutesToHHMMString(aggregatedData(key)) & vbCrLf
            End If
        Else
            ' 最大行数超過時は残り件数のみ表示
            msg = msg & "…ほか " & (aggregatedData.Count - MAX_LINES) & " 件" & vbCrLf
            Exit For
        End If
    Next
    
    ' === ユーザーに確認ダイアログを表示 ===
    ShowPreviewAndConfirm = (MsgBox(msg, vbYesNo + vbQuestion, "転記内容の確認") = vbYes)
End Function

'===============================================================================
' 【集計データの書き込み】
' 集計されたデータを月次シートの適切な列に書き込む
' ※既存値がある場合は上書き（加算はしない）
'===============================================================================
Private Sub WriteAggregatedDataToSheet( _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef aggregatedData As Object, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult)

    ' === 変数宣言 ===
    Dim key As Variant                  ' 辞書のキー（作業コード|作番）
    Dim parts() As String               ' キーの分割結果
    Dim targetCol As Long               ' 書き込み対象列

    ' === 結果カウンタ初期化 ===
    result.ProcessedCount = 0
    result.DuplicateCount = 0
    result.NewColumnsAdded = 0

    ' === 集計データの各項目を処理 ===
    For Each key In aggregatedData.Keys
        parts = Split(CStr(key), KEY_SEPARATOR)  ' 0:作業コード 1:作番
        
        If UBound(parts) >= 1 Then
            ' 書き込み対象列を取得（存在しない場合は新規作成）
            targetCol = GetOrCreateColumn(parts(0), parts(1), config, wsMonthly, mapDict, lastCol, result)
            
            If targetCol > 0 Then
                ' 実際のデータ書き込み実行
                WriteTimeDataToCell wsMonthly, config.targetRow, targetCol, aggregatedData(key), result
                result.ProcessedCount = result.ProcessedCount + 1
            End If
        End If
    Next
End Sub

'===============================================================================
' 【列取得または新規作成】
' 指定された作業コード+作番の列を取得、存在しない場合は設定に応じて新規作成
' 戻り値：列番号（0=作成拒否または失敗）
'===============================================================================
Private Function GetOrCreateColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult) As Long

    ' === 変数宣言 ===
    Dim key As String: key = category & KEY_SEPARATOR & workNo  ' キー生成
    Dim newCol As Long                                          ' 新規作成列番号

    ' === 既存列が存在する場合はその列番号を返す ===
    If mapDict.Exists(key) Then
        GetOrCreateColumn = mapDict(key)
        Exit Function
    End If

    ' === 列追加ポリシーに応じた処理 ===
    Select Case config.AddPolicy
        Case AddPolicy_Reject
            ' 列追加拒否：0を返して処理スキップ
            GetOrCreateColumn = 0
            
        Case AddPolicy_Auto
            ' 自動列追加：確認なしで新規列作成
            newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
            If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
            GetOrCreateColumn = newCol
            
        Case Else ' AddPolicy_Prompt
            ' 確認後列追加：ユーザーに確認してから作成
            If ConfirmColumnCreation(category, workNo) Then
                newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
                If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
                GetOrCreateColumn = newCol
            Else
                GetOrCreateColumn = 0  ' ユーザーが拒否
            End If
    End Select
End Function

'===============================================================================
' 【新規列作成】
' 月次シートに新しい作業コード+作番の列を作成し、適切な書式を設定
' 戻り値：作成した列番号
'===============================================================================
Private Function CreateNewColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long) As Long

    ' === 新規列番号決定 ===
    Dim newCol As Long
    newCol = lastCol + 1

    ' === ヘッダ情報設定 ===
    wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).value = category  ' 作業コード行
    wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).value = workNo    ' 作番行

    ' === 既存列の書式をコピーして適用 ===
    ApplyColumnFormatting wsMonthly, newCol, IIf(lastCol >= MONTHLY_MIN_COL, lastCol, MONTHLY_MIN_COL)
    
    ' === データ部分の時間書式設定 ===
    SetDataColumnFormat wsMonthly, newCol

    ' === マッピング辞書と最終列番号の更新 ===
    mapDict.Add category & KEY_SEPARATOR & workNo, newCol
    lastCol = newCol

    ' === 作成した列番号を返却 ===
    CreateNewColumn = newCol
End Function

'===============================================================================
' 【列書式適用】
' 新規作成列に既存列の書式（幅、配置、色など）をコピー
'===============================================================================
Private Sub ApplyColumnFormatting(ByRef wsMonthly As Worksheet, ByVal newCol As Long, ByVal sourceCol As Long)
    On Error Resume Next  ' 書式エラーは無視して続行
    
    ' === 列幅のコピー ===
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(sourceCol).ColumnWidth
    
    ' === ヘッダ行（作業コード行）の書式コピー ===
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).WrapText
    End With
    
    ' === 作番行の書式コピー ===
    With wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).WrapText
    End With
    
    On Error GoTo 0  ' エラーハンドリングを元に戻す
End Sub

'===============================================================================
' 【データ列書式設定】
' 新規作成列のデータ部分に時間表示書式を適用
'===============================================================================
Private Sub SetDataColumnFormat(ByRef wsMonthly As Worksheet, ByVal col As Long)
    ' === データ範囲の最終行を取得 ===
    Dim lastRow As Long
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then lastRow = MONTHLY_DATA_START_ROW + DEFAULT_PREVIEW_ROWS
    
    ' === データ範囲全体に時間書式を適用 ===
    With wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, col), wsMonthly.Cells(lastRow, col))
        .NumberFormatLocal = TIME_FORMAT  ' [hh]mm 形式（24時間以上対応）
    End With
End Sub

'===============================================================================
' 【列作成確認ダイアログ】
' 新規列作成前にユーザーに確認を求める
' 戻り値：True=作成許可, False=作成拒否
'===============================================================================
Private Function ConfirmColumnCreation(ByVal category As String, ByVal workNo As String) As Boolean
    ConfirmColumnCreation = (MsgBox( _
        "作業ｺｰﾄﾞ『" & category & "』＋作番『" & workNo & "』の列がありません。" & vbCrLf & _
        "月次データシートに新しい列を追加しますか？", _
        vbYesNo + vbQuestion, "列の追加確認") = vbYes)
End Function

'===============================================================================
' 【セルへの時間データ書き込み】
' 指定セルに時間データを書き込み、既存値がある場合は重複処理を実行
' ※重要：既存値は上書き（加算はしない）
'===============================================================================
Private Sub WriteTimeDataToCell( _
    ByRef wsMonthly As Worksheet, _
    ByVal targetRow As Long, ByVal targetCol As Long, _
    ByVal minutes As Double, _
    ByRef result As ProcessResult)

    ' === 変数宣言 ===
    Dim existingValue As Double         ' 既存値（シリアル値）
    Dim newValue As Double              ' 新規値（シリアル値）
    Dim isDup As Boolean                ' 重複フラグ

    ' === 既存値チェック ===
    existingValue = NzD(wsMonthly.Cells(targetRow, targetCol).value, 0#)
    newValue = MinutesToSerial(minutes)  ' 分数をExcelシリアル値に変換
    isDup = (existingValue <> 0#)       ' 既存値があるかチェック

    ' === 重複時の処理 ===
    If isDup Then
        result.DuplicateCount = result.DuplicateCount + 1
        
        ' セルを黄色でハイライト
        HighlightDuplicateCell wsMonthly.Cells(targetRow, targetCol)
        
        ' メッセージ列に元の値を記録（上書きされる値のみ）
        LogDuplicateMessage wsMonthly, targetRow, _
            CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, targetCol).value), _
            CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, targetCol).value), _
            existingValue
    End If

    ' === セルへの値書き込み（上書き固定） ===
    With wsMonthly.Cells(targetRow, targetCol)
        .value = newValue                       ' 新しい値で上書き
        .NumberFormatLocal = TIME_FORMAT        ' 時間表示書式適用
    End With
End Sub

'===============================================================================
' 【重複セルハイライト】
' 重複が発生したセルを指定色でハイライト表示
'===============================================================================
Private Sub HighlightDuplicateCell(ByRef cell As Range)
    With cell.Interior
        .Pattern = xlSolid                      ' 塗りつぶしパターン
        .Color = DUP_HIGHLIGHT_COLOR           ' 指定色（通常は黄色）
    End With
End Sub

'===============================================================================
' 【重複メッセージ記録】
' 重複発生時に元の値をメッセージ列に記録
' ※仕様：元々入っていた時間のみを記録（例: "1:30"）
'===============================================================================
Private Sub LogDuplicateMessage( _
    ByRef wsMonthly As Worksheet, ByVal rowNum As Long, _
    ByVal category As String, ByVal workNo As String, _
    ByVal oldValue As Double)
    
    ' === 重複メッセージ生成 ===
    ' 「既存値検出: [作番|作業コード] 旧=時間表示」の形式
    Dim message As String
    message = "既存値検出: [" & workNo & "|" & category & "] 旧=" & SerialToHHMMString(oldValue)
    
    ' === メッセージ列に追記 ===
    AppendMessageToCell wsMonthly, rowNum, message
End Sub

'===============================================================================
' 【クリップボード操作】
' 収集データをタブ区切りテキストとしてクリップボードにコピー
'===============================================================================
Private Sub CopyDataToClipboard(ByRef items As collection, ByRef wsData As Worksheet)
    ' === 変数宣言 ===
    Dim sb As String                    ' 出力文字列バッファ
    Dim i As Long                       ' ループカウンタ
    Dim v As Variant                    ' コレクションの各要素

    ' === 各データ項目をタブ区切り形式で連結 ===
    For i = 1 To items.Count
        v = items(i)  ' [WorkNo, Category, Minutes, RowIndex]
        
        ' 形式：作番 + タブ + 作業コード + タブ + タブ + 時間表示文字列 + 改行
        ' 重要：時間と作業コードの間に1つ余分なタブを挿入（Excel貼り付け時の体裁調整）
        sb = sb & CStr(v(1)) & vbTab & CStr(v(2)) & vbTab & vbTab & _
                 CStr(wsData.Cells(CLng(v(4)), COL_TIME).text) & vbCrLf
    Next
    
    ' === クリップボードにコピー実行 ===
    If Len(sb) > 0 Then CopyTextToClipboardSafe sb
End Sub

'===============================================================================
' 【安全なクリップボードコピー】
' Forms.DataObjectを優先し、失敗時はWinAPIにフォールバック
'===============================================================================
Private Sub CopyTextToClipboardSafe(ByVal textToCopy As String)
    On Error GoTo APIFallback
    
    ' === 方法1：Forms.DataObjectを使用（参照設定不要） ===
    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject")  ' 無い環境もある
    dataObject.SetText textToCopy
    dataObject.PutInClipboard
    Exit Sub

APIFallback:
    ' === 方法2：WinAPI直接呼び出し（フォールバック） ===
    CopyTextToClipboardWinAPI textToCopy
End Sub

'===============================================================================
' 【WinAPIクリップボードコピー】
' Windows APIを直接使用したクリップボードへのテキストコピー
'===============================================================================
Private Sub CopyTextToClipboardWinAPI(ByVal textToCopy As String)
    ' === 変数宣言（64/32ビット対応） ===
#If VBA7 Then
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
#Else
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
#End If
    Dim bytesNeeded As Long             ' 必要メモリサイズ
    Dim retryCount As Long              ' リトライカウンタ

    ' === 空文字列チェック ===
    If Len(textToCopy) = 0 Then Exit Sub
    
    ' === Unicode文字列用メモリ確保 ===
    bytesNeeded = (Len(textToCopy) + 1) * 2  ' Unicode = 2バイト/文字 + 終端文字
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, bytesNeeded)
    If hGlobalMemory = 0 Then Exit Sub

    ' === メモリロックと文字列コピー ===
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        lstrcpyW lpGlobalMemory, StrPtr(textToCopy)  ' Unicode文字列コピー
        GlobalUnlock hGlobalMemory

        ' === クリップボード操作（リトライ付き） ===
        For retryCount = 1 To MAX_RETRY_CLIPBOARD
            If OpenClipboard(0) <> 0 Then Exit For  ' クリップボード取得成功
            DoEvents  ' 他の処理に制御を渡してからリトライ
        Next retryCount

        If retryCount <= MAX_RETRY_CLIPBOARD Then
            ' クリップボード取得成功時の処理
            EmptyClipboard
            If SetClipboardData(CF_UNICODETEXT, hGlobalMemory) = 0 Then
                GlobalFree hGlobalMemory  ' 失敗時はメモリ解放
            End If
            CloseClipboard
        Else
            ' クリップボード取得失敗時はメモリ解放
            GlobalFree hGlobalMemory
        End If
    Else
        ' メモリロック失敗時はメモリ解放
        GlobalFree hGlobalMemory
    End If
End Sub

'===============================================================================
' 【ユーティリティ関数群】
' 時間変換、文字列処理、検索などの汎用機能
'===============================================================================

'===============================================================================
' 【拡張時間変換】
' 様々な形式の時間データを分数に統一変換
' 対応形式：Date型、シリアル値、HHMM形式（整数・文字列）
'===============================================================================
Private Function ConvertToMinutesEx(ByVal timeValue As Variant) As Double
    Dim s As String
    ConvertToMinutesEx = 0  ' デフォルト値
    
    ' === 空値チェック ===
    If IsEmpty(timeValue) Then Exit Function

    ' === Date型の場合 ===
    If IsDate(timeValue) Then
        ConvertToMinutesEx = CDbl(CDate(timeValue)) * MINUTES_PER_DAY
        Exit Function
    End If

    ' === 数値型の場合 ===
    If IsNumeric(timeValue) Then
        If InStr(1, CStr(timeValue), ".") > 0 Then
            ' 小数点あり → Excelシリアル値として処理
            ConvertToMinutesEx = CDbl(timeValue) * MINUTES_PER_DAY
        Else
            ' 整数 → HHMM形式として処理
            ConvertToMinutesEx = ParseHHMMInteger(CLng(timeValue))
        End If
        Exit Function
    End If

    ' === 文字列型の場合 ===
    s = Trim$(CStr(timeValue))
    If InStr(s, ":") > 0 Then
        ' コロン区切り → H:MM形式として処理
        ConvertToMinutesEx = ParseHHMMString(s)
    ElseIf IsNumeric(s) Then
        ' 数値文字列 → HHMM整数として処理
        ConvertToMinutesEx = ParseHHMMInteger(CLng(Val(s)))
    End If
End Function

'===============================================================================
' 【HHMM整数解析】
' HHMM形式の整数を分数に変換（例：130 → 90分、1030 → 630分）
'===============================================================================
Private Function ParseHHMMInteger(ByVal hhmmValue As Long) As Double
    Dim hours As Long, minutes As Long  ' 時間・分
    Dim t As String                     ' 数値文字列
    
    ParseHHMMInteger = 0  ' デフォルト値
    If hhmmValue < 0 Then Exit Function  ' 負数は無効
    
    ' === 桁数に応じた時間・分の分離 ===
    t = CStr(hhmmValue)
    Select Case Len(t)
        Case 1, 2
            ' 1-2桁：分のみ（例：5 → 0:05, 30 → 0:30）
            minutes = hhmmValue: hours = 0
        Case 3, 4
            ' 3-4桁：HHMM形式（例：130 → 1:30, 1030 → 10:30）
            hours = CLng(Left$(t, Len(t) - 2))
            minutes = CLng(Right$(t, 2))
        Case Else
            ' 5桁以上は無効
            Exit Function
    End Select
    
    ' === 分の有効性チェック（0-59の範囲） ===
    If minutes >= 0 And minutes < MAX_MINUTES_PER_HOUR Then
        ParseHHMMInteger = hours * MINUTES_PER_HOUR + minutes
    End If
End Function

'===============================================================================
' 【H:MM文字列解析】
' 「H:MM」形式の文字列を分数に変換（例：「1:30」 → 90分）
'===============================================================================
Private Function ParseHHMMString(ByVal timeString As String) As Double
    Dim parts() As String               ' コロン分割結果
    Dim h As Long, m As Long           ' 時間・分
    
    ParseHHMMString = 0  ' デフォルト値
    
    ' === コロンで分割 ===
    parts = Split(timeString, ":")
    If UBound(parts) = 1 Then  ' 2つの部分に分割されたかチェック
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
            h = CLng(parts(0)): m = CLng(parts(1))
            
            ' 分の有効性チェック（0-59の範囲）
            If m >= 0 And m < MAX_MINUTES_PER_HOUR Then
                ParseHHMMString = h * MINUTES_PER_HOUR + m
            End If
        End If
    End If
End Function

'===============================================================================
' 【分数→シリアル値変換】
' 分数をExcelの時間シリアル値に変換（Excel内部での時間表現）
'===============================================================================
Private Function MinutesToSerial(ByVal totalMinutes As Double) As Double
    MinutesToSerial = totalMinutes / MINUTES_PER_DAY
End Function

'===============================================================================
' 【分数→H:MM文字列変換】
' 分数を「H:MM」形式の文字列に変換（表示用）
'===============================================================================
Private Function MinutesToHHMMString(ByVal totalMinutes As Double) As String
    Dim h As Long, m As Long           ' 時間・分
    
    ' === 0以下の場合は「0:00」 ===
    If totalMinutes <= 0 Then
        MinutesToHHMMString = "0:00": Exit Function
    End If
    
    ' === 時間・分の計算 ===
    h = Int(totalMinutes / MINUTES_PER_HOUR)         ' 時間部分
    m = Round(totalMinutes - h * MINUTES_PER_HOUR, 0) ' 分部分（四捨五入）
    
    ' === 分が60になった場合の繰り上がり処理 ===
    If m = MAX_MINUTES_PER_HOUR Then h = h + 1: m = 0
    
    ' === 書式整形して返却 ===
    MinutesToHHMMString = Format$(h, "0") & ":" & Format$(m, "00")
End Function

'===============================================================================
' 【シリアル値→H:MM文字列変換】
' Excelシリアル値を「H:MM」形式の文字列に変換
'===============================================================================
Private Function SerialToHHMMString(ByVal serialValue As Double) As String
    SerialToHHMMString = MinutesToHHMMString(serialValue * MINUTES_PER_DAY)
End Function

'===============================================================================
' 【日付一致行検索】
' 月次シートから指定日付と一致する行番号を検索
' 戻り値：行番号（0=見つからない）
'===============================================================================
Private Function FindMatchingDateRow(ByRef wsMonthly As Worksheet, ByVal targetDate As Date) As Long
    Dim lastRow As Long, r As Long     ' 行ループ変数
    Dim d As Date                      ' 各行の日付
    
    FindMatchingDateRow = 0  ' デフォルト値
    
    ' === 日付列の最終行取得 ===
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then Exit Function  ' データなし
    
    ' === 各行の日付をチェック ===
    For r = MONTHLY_DATA_START_ROW To lastRow
        If IsDate(wsMonthly.Cells(r, COL_DATE).value) Then
            d = CDate(wsMonthly.Cells(r, COL_DATE).value)
            ' 日付部分のみ比較（時間は無視）
            If Int(d) = Int(targetDate) Then
                FindMatchingDateRow = r: Exit Function  ' 一致する行が見つかった
            End If
        End If
    Next
End Function

'===============================================================================
' 【メッセージ列ヘッダ確保】
' 月次シートのメッセージ列（A列）にヘッダが設定されていることを確認・設定
'===============================================================================
Private Sub EnsureMessageColumnHeader(ByRef wsMonthly As Worksheet)
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, COL_MESSAGE)
        ' ヘッダが空の場合のみ設定
        If Trim$(CStr(.value)) = "" Then
            .value = "メッセージ"
            .Font.Bold = True  ' 太字で強調
        End If
    End With
End Sub

'===============================================================================
' 【メッセージセル追記】
' 指定行のメッセージ列にメッセージを追記（既存内容がある場合は改行で区切り）
'===============================================================================
Private Sub AppendMessageToCell(ByRef wsMonthly As Worksheet, ByVal rowNum As Long, ByVal message As String)
    With wsMonthly.Cells(rowNum, COL_MESSAGE)
        If Len(.value) = 0 Then
            ' 初回メッセージ
            .value = message
        Else
            ' 既存メッセージに追記（改行区切り）
            .value = CStr(.value) & MESSAGE_SEPARATOR & message
        End If
    End With
End Sub

'===============================================================================
' 【Null値安全数値変換】
' Variant値を安全にDouble型に変換（エラー・空値時はデフォルト値）
'===============================================================================
Private Function NzD(ByVal value As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    On Error Resume Next
    
    ' === 各種無効値のチェック ===
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Or value = "" Then
        NzD = defaultValue
    ElseIf IsNumeric(value) Then
        NzD = CDbl(value)  ' 数値変換
    Else
        NzD = defaultValue  ' 変換不可時はデフォルト
    End If
    
    On Error GoTo 0
End Function

'===============================================================================
' 【月次シートエラー表示クリア】
' 処理開始時に月次シートのエラー表示セル（I1）をクリア
'===============================================================================
Private Sub ClearErrorCellOnMonthlySheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    If Not ws Is Nothing Then
        ws.Range("I1").ClearContents  ' 内容クリア
        ws.Range("I1").WrapText = True ' 自動折り返し有効
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' 【月次シートエラー報告】
' エラー発生時にエラーメッセージを月次シートの指定セル（I1）に表示
'===============================================================================
Private Sub ReportErrorToMonthlySheet(ByVal message As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    If Not ws Is Nothing Then
        With ws.Range("I1")
            .value = message           ' エラーメッセージ設定
            .WrapText = True          ' 自動折り返しで見やすく表示
        End With
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' 【アプリケーション状態管理】
' Excel処理高速化のための状態保存・設定・復元
'===============================================================================

'===============================================================================
' 【アプリケーション状態保存・高速化設定】
' 現在の設定を保存して処理高速化のための設定に変更
'===============================================================================
Private Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    ' === 現在の状態を保存 ===
    With prevState
        .ScreenUpdating = Application.ScreenUpdating  ' 画面更新状態
        .EnableEvents = Application.EnableEvents      ' イベント有効状態
        .Calculation = Application.Calculation        ' 計算モード
    End With
    
    ' === 高速化のための設定変更 ===
    With Application
        .ScreenUpdating = False              ' 画面更新停止（描画処理を省略）
        .EnableEvents = False               ' イベント処理停止（変更イベント等を無効化）
        .Calculation = xlCalculationManual  ' 自動計算停止（数式再計算を抑制）
    End With
End Sub

'===============================================================================
' 【アプリケーション状態復元】
' 保存していた元の設定に復元
'===============================================================================
Private Sub RestoreApplicationState(ByRef prevState As ApplicationState)
    With Application
        .Calculation = prevState.Calculation        ' 計算モード復元
        .EnableEvents = prevState.EnableEvents      ' イベント処理復元
        .ScreenUpdating = prevState.ScreenUpdating  ' 画面更新復元（最後に実行）
    End With
End Sub

'===============================================================================
' 【シート保護管理】
' 月次シートの保護状態を適切に管理（解除→処理→復元）
'===============================================================================

'===============================================================================
' 【シート保護解除（必要時）】
' シートが保護されている場合の安全な解除処理
' パスワードが必要な場合はユーザーに入力を求める
'===============================================================================
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    ' === 現在の保護状態を記録 ===
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    ' === 保護されていない場合は何もしない ===
    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    ' === パスワードなし保護解除の試行 ===
    On Error Resume Next
    ws.Unprotect ""  ' 空パスワードで解除試行
    If Err.Number = 0 Then
        ' 解除成功
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    ' === パスワード入力による保護解除 ===
    Err.Clear
    protInfo.Password = InputBox("シート『" & ws.Name & "』のパスワードを入力してください。", "保護解除")
    
    ' ユーザーがキャンセルした場合
    If protInfo.Password = "" Then
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ' パスワードによる解除試行
    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0)  ' エラーがなければ成功
    On Error GoTo 0
End Function

'===============================================================================
' 【シート保護復元】
' 処理完了後に元の保護状態を復元
'===============================================================================
Private Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    ' === 元々保護されていた場合のみ復元 ===
    If protInfo.IsProtected Then
        On Error Resume Next
        
        If protInfo.Password = "" Then
            ' パスワードなし保護
            ws.Protect UserInterfaceOnly:=True
        Else
            ' パスワード付き保護
            ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        End If
        
        On Error GoTo 0
    End If
End Sub

'===============================================================================
' 【エラーハンドリング】
' 統一されたエラー処理とユーザーフレンドリーなメッセージ生成
'===============================================================================

'===============================================================================
' 【カスタムエラー発生】
' 独自エラーコードとメッセージでエラーを発生させる
'===============================================================================
Private Sub RaiseCustomError(ByVal errorCode As Long, ByVal description As String)
    Err.Raise errorCode, "TransferDataModule", description
End Sub

'===============================================================================
' 【エラー詳細情報生成】
' エラー番号に応じたユーザーフレンドリーなエラーメッセージを生成
'===============================================================================
Private Function GetErrorDetails(ByVal errNumber As Long, ByVal errDescription As String) As String
    Select Case errNumber
        Case ERR_SHEET_NOT_FOUND
            ' シート不存在エラー
            GetErrorDetails = "必要なシートが見つかりません: " & errDescription
            
        Case ERR_INVALID_DATE
            ' 日付無効エラー
            GetErrorDetails = "日付が無効です: " & errDescription
            
        Case ERR_NO_DATA
            ' データなしエラー
            GetErrorDetails = "転記するデータがありません: " & errDescription
            
        Case ERR_DATE_NOT_FOUND
            ' 対象日付不存在エラー
            GetErrorDetails = "対象日付が月次シートに見つかりません: " & errDescription
            
        Case ERR_PROTECTION_FAILED
            ' 保護解除失敗エラー
            GetErrorDetails = "シート保護の解除に失敗しました: " & errDescription
            
        Case 9 ' Subscript out of range
            ' 配列・コレクション範囲外エラー（特に詳しく説明）
            GetErrorDetails = FriendlyErrorMessage9(errDescription)
            
        Case Else
            ' その他の予期しないエラー
            GetErrorDetails = "予期しないエラーが発生しました (エラー #" & errNumber & "): " & errDescription
    End Select
End Function

'===============================================================================
' 【エラー#9詳細説明】
' 最も頻発するエラー#9に対する詳細で分かりやすい説明
'===============================================================================
Private Function FriendlyErrorMessage9(ByVal errDesc As String) As String
    FriendlyErrorMessage9 = _
        "エラー #9（インデックスが有効範囲にありません）" & vbCrLf & _
        "考えられる原因と対処:" & vbCrLf & _
        "・シート名の確認：『" & DATA_SHEET_NAME & "』『" & MONTHLY_SHEET_NAME & "』が存在するか" & vbCrLf & _
        "・データ形式の確認：作番と作業ｺｰﾄﾞが正しく入力されているか" & vbCrLf & _
        "・列構造の確認：必要な列が存在し、正しい位置にあるか" & vbCrLf & _
        vbCrLf & "詳細: " & errDesc
End Function

'===============================================================================
' 【結果表示】
' 処理完了時のユーザーへの結果報告
'===============================================================================

'===============================================================================
' 【転記結果表示】
' 処理結果の詳細をユーザーに分かりやすく表示
'===============================================================================
Private Sub ShowTransferResults(ByRef result As ProcessResult)
    Dim message As String
    
    If result.Success Then
        ' === 成功時のメッセージ構成 ===
        message = "転記処理が完了しました。" & vbCrLf & vbCrLf & _
                  "処理件数: " & result.ProcessedCount & " 件" & vbCrLf
        
        ' 重複検知情報（該当する場合のみ）
        If result.DuplicateCount > 0 Then
            message = message & "重複検知: " & result.DuplicateCount & " 件（黄色ハイライト表示）" & vbCrLf
        End If
        
        ' 新規列追加情報（該当する場合のみ）
        If result.NewColumnsAdded > 0 Then
            message = message & "新規列追加: " & result.NewColumnsAdded & " 列" & vbCrLf
        End If
        
        ' 追加メッセージ（ある場合のみ）
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & "メッセージ:" & vbCrLf & result.Messages
        End If
        
        ' 成功ダイアログ表示
        MsgBox message, vbInformation, "転記完了"
        
    Else
        ' === 中止・失敗時のメッセージ ===
        message = "転記処理が中止されました。"
        
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & vbCrLf & result.Messages
        End If
        
        ' 警告ダイアログ表示
        MsgBox message, vbExclamation, "処理中止"
    End If
End Sub

'===============================================================================
' 【モジュール終了】
' 
' 【使用方法】
' 1. このモジュールをVBAプロジェクトにインポート
' 2. TransferDataToMonthlySheet() を実行
' 3. 必要に応じて定数セクションの設定を環境に合わせて調整
' 
' 【カスタマイズポイント】
' ・シート名：DATA_SHEET_NAME, MONTHLY_SHEET_NAME
' ・セル位置：DATE_CELL_PRIORITY, DATE_CELL_NORMAL
' ・行列番号：各種 _ROW, _COL 定数
' ・動作設定：AUTO_ADD_POLICY, DRY_RUN
' 
' 【注意事項】
' ・Excel 2016以降、Windows 11での動作を想定
' ・シート保護がある場合は解除パスワードが必要
' ・大量データ処理時は画面更新停止により高速化
'===============================================================================