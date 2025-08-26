'===============================================================================
' Module : ModDataTransfer
' 機能   : 「データ登録」→「月次データ」へ転記（区分×作番で列特定／列の自動追加）
' 仕様   : 入力表を走査→(区分,作番)で分単位集計→日付行へ書き込み
'          重複時は黄色ハイライト＆A列メッセージ記録。事前プレビュー＆確認あり。
' 対象   : Excel 2016+ / Windows（Forms.DataObjectが使えない場合はWinAPIフォールバック）
' 留意   : 日付行は「月次データ」に既存前提（無い場合はエラー表示）
'===============================================================================
Option Explicit

'=========================
' クリップボード WinAPI（64/32両対応）
'=========================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_UNICODETEXT As Long = 13

'=========================
' 設定（ブックに合わせて必要なら修正）
'=========================
' シート名
Private Const ACQUISITION_SHEET_NAME As String = "データ取得"
Private Const DATA_SHEET_NAME        As String = "データ登録"
Private Const MONTHLY_SHEET_NAME     As String = "月次データ"

' 「データ登録」シート：日付セル（優先→通常の順で採用）
Private Const DATE_CELL_PRIORITY As String = "D4"
Private Const DATE_CELL_NORMAL   As String = "D3"

' 「データ取得」入力表：列定義（A=1, B=2...）
Private Const DATA_START_ROW     As Long = 8   ' 入力表の開始行
Private Const COL_MESSAGE        As Long = 1   ' A列（コメント等：未使用でも可）
Private Const COL_DATE           As Long = 2   ' B列（任意：未使用でも可）
Private Const COL_WORKNO         As Long = 3   ' C列：作番
Private Const COL_CATEGORY       As Long = 4   ' D列：区分
Private Const COL_TIME           As Long = 5   ' E列：時間（hhmm / hh:mm / 連続分 / Excel時刻のいずれも可）

' 「月次データ」シート：見出し・データ配置
Private Const MONTHLY_WORKNO_ROW      As Long = 8   ' 作番行
Private Const MONTHLY_HEADER_ROW      As Long = 9   ' 区分類（カテゴリ）行
Private Const MONTHLY_DATA_START_ROW  As Long = 10  ' データ開始行（ここから下に日付ごとの行）
Private Const MONTHLY_MIN_COL         As Long = 3   ' データ列の最小列（C列）

' A/B列（メッセージ／日付）
Private Const COL_MONTHLY_MESSAGE As Long = 1   ' A列：メッセージ（重複等の記録）
Private Const COL_MONTHLY_DATE    As Long = 2   ' B列：日付

' 表示フォーマット
Private Const TIME_FORMAT   As String = "[h]:mm"     ' 時間（合計時間でも崩れない形式）
Private Const DATE_FORMAT   As String = "yyyy/mm/dd" ' 日付
Private Const PREVIEW_TAB   As String = vbTab

' キー生成
Private Const KEY_SEPARATOR      As String = "|"
Private Const MESSAGE_SEPARATOR  As String = vbLf

' 自動追加ポリシー（列が無いとき）
Private Const AddPolicy_Prompt   As Long = 0   ' 都度確認
Private Const AddPolicy_Auto     As Long = 1   ' 全自動で追加
Private Const AddPolicy_Reject   As Long = 2   ' 一括拒否（追加しない）

Private Const AUTO_ADD_POLICY    As Long = AddPolicy_Prompt  ' 既定
Private Const ACCUMULATE_MODE    As Boolean = True           ' 既存値に加算(True) / 上書き(False)
Private Const DRY_RUN            As Boolean = False          ' True=書き込みなしでテスト
Private Const DUP_HIGHLIGHT_COLOR As Long = vbYellow

' 時間換算
Private Const MINUTES_PER_HOUR     As Double = 60#
Private Const MINUTES_PER_DAY      As Double = 1440#
Private Const MAX_MINUTES_PER_HOUR As Long = 60

' プレビュー行数
Private Const DEFAULT_PREVIEW_ROWS As Long = 31

' エラーコード
Private Const ERR_SHEET_NOT_FOUND   As Long = vbObjectError + 1
Private Const ERR_INVALID_DATE      As Long = vbObjectError + 2
Private Const ERR_NO_DATA           As Long = vbObjectError + 3
Private Const ERR_DATE_NOT_FOUND    As Long = vbObjectError + 4
Private Const ERR_PROTECTION_FAILED As Long = vbObjectError + 5

'=========================
' UDT（2016互換：Object/Worksheetを入れない、XlCalculationはLongで保持）
'=========================
Private Type ApplicationState
    ScreenUpdating As Boolean
    EnableEvents   As Boolean
    Calculation    As Long
End Type

Private Type SheetProtectionInfo
    IsProtected As Boolean
    Password    As String
End Type

Private Type TransferConfig
    targetDate     As Date
    targetRow      As Long
    accumulateMode As Boolean
    DryRun         As Boolean
    AddPolicy      As Long
End Type

Private Type ProcessResult
    ProcessedCount  As Long
    DuplicateCount  As Long
    ErrorCount      As Long
    NewColumnsAdded As Long
    Messages        As String
    Success         As Boolean
End Type

'===============================================================================
' エントリポイント
'===============================================================================
Public Sub TransferDataToMonthlySheet()
    Dim prevState As ApplicationState
    Dim config As TransferConfig
    Dim result As ProcessResult
    Dim protectionInfo As SheetProtectionInfo
    Dim wsData As Worksheet, wsMonthly As Worksheet

    SaveAndSetApplicationState prevState
    On Error GoTo ErrorHandler

    ' 初期化（シート存在・保護解除・日付確定・日付行取得・ポリシー設定）
    If Not InitializeTransferConfig(config, protectionInfo, wsData, wsMonthly) Then GoTo CleanUp

    ' 転記実行
    ExecuteDataTransfer config, wsData, wsMonthly, result

    ' 結果表示
    ShowTransferResults result

CleanUp:
    RestoreSheetProtection wsMonthly, protectionInfo
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    MsgBox GetErrorDetails(Err.Number, Err.Description), vbCritical, "転記エラー"
    Resume CleanUp
End Sub

'===============================================================================
' 初期化
'===============================================================================
Private Function InitializeTransferConfig( _
    ByRef config As TransferConfig, _
    ByRef protInfo As SheetProtectionInfo, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet) As Boolean

    InitializeTransferConfig = False

    If Not GetAndValidateWorksheets(wsData, wsMonthly) Then Exit Function

    ' 月次シートの保護解除（必要時）
    If Not UnprotectSheetIfNeeded(wsMonthly, protInfo) Then Exit Function

    ' 日付取得（D4優先→D3）
    If Not DetermineTargetDate(wsData, config.targetDate) Then Exit Function

    ' 対応する日付行を検索（無ければエラー）
    config.targetRow = FindMatchingDateRow(wsMonthly, config.targetDate)
    If config.targetRow = 0 Then
        RaiseCustomError ERR_DATE_NOT_FOUND, Format$(config.targetDate, DATE_FORMAT)
        Exit Function
    End If

    ' ポリシー設定
    config.accumulateMode = ACCUMULATE_MODE
    config.DryRun = DRY_RUN
    config.AddPolicy = AUTO_ADD_POLICY

    InitializeTransferConfig = True
End Function

Private Function GetAndValidateWorksheets(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    On Error GoTo SheetError
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    On Error GoTo 0

    If Not ValidateSheetStructure(wsData, wsMonthly) Then
        GetAndValidateWorksheets = False
        Exit Function
    End If
    GetAndValidateWorksheets = True
    Exit Function

SheetError:
    RaiseCustomError ERR_SHEET_NOT_FOUND, "シート: " & DATA_SHEET_NAME & ", " & MONTHLY_SHEET_NAME
    GetAndValidateWorksheets = False
End Function

Private Function ValidateSheetStructure(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    ' 入力表の有無（作番列で最終行を確認）
    If wsData.Cells(wsData.Rows.Count, COL_WORKNO).End(xlUp).Row < DATA_START_ROW Then
        ValidateSheetStructure = False: Exit Function
    End If
    ' 月次ヘッダの有無（右端列が最低列を超えるか）
    If wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column < MONTHLY_MIN_COL Then
        ValidateSheetStructure = False: Exit Function
    End If
    ValidateSheetStructure = True
End Function

Private Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef outDate As Date) As Boolean
    Dim v As Variant
    v = wsData.Range(DATE_CELL_PRIORITY).Value
    If IsDate(v) Then outDate = CDate(v): DetermineTargetDate = True: Exit Function
    v = wsData.Range(DATE_CELL_NORMAL).Value
    If IsDate(v) Then outDate = CDate(v): DetermineTargetDate = True: Exit Function
    RaiseCustomError ERR_INVALID_DATE, "日付セル(" & DATE_CELL_PRIORITY & " または " & DATE_CELL_NORMAL & ")が未入力または不正です。"
    DetermineTargetDate = False
End Function

'===============================================================================
' 転記本体
'===============================================================================
Private Sub ExecuteDataTransfer( _
    ByRef config As TransferConfig, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet, _
    ByRef result As ProcessResult)

    Dim items As Collection              ' 各行: Array(WorkNo, Category, Minutes, RowIndex)
    Dim aggregated As Object             ' Scripting.Dictionary (key="区分|作番", val=合計分)
    Dim mapDict As Object                ' 既存マッピング (key="区分|作番", val=列番号)
    Dim lastCol As Long

    ' 入力表→有効データ抽出
    Set items = CollectTimeDataFromSheet(wsData)
    If items.Count = 0 Then
        RaiseCustomError ERR_NO_DATA, "入力表に有効なデータがありません。"
        Exit Sub
    End If

    ' (区分,作番)ごとに分単位で集計
    Set aggregated = AggregateTimeData(items)

    ' 既存（区分,作番）→列番号のマップを構築
    Set mapDict = CreateObject("Scripting.Dictionary")
    BuildColumnMapping wsMonthly, lastCol, mapDict

    ' プレビュー＆確認
    If Not ShowPreviewAndConfirm(config.targetDate, aggregated) Then
        result.Success = False
        Exit Sub
    End If

    ' ドライランならここで終了
    If config.DryRun Then
        result.Messages = "ドライランのため書き込みは実施しませんでした。"
        result.Success = True
        Exit Sub
    End If

    ' コピー用に「データ取得」の内容（作番・区分・表示テキスト）をクリップボードへ
    CopyDataToClipboard items, wsData

    ' 月次シートのメッセージ列ヘッダを整備
    EnsureMessageColumnHeader wsMonthly

    ' 書き込み
    WriteAggregatedDataToSheet config, wsMonthly, aggregated, mapDict, lastCol, result

    result.Success = True
End Sub

' 入力表スキャン → Collection of Variant(WorkNo, Category, Minutes, RowIndex)
Private Function CollectTimeDataFromSheet(ByRef wsData As Worksheet) As Collection
    Dim col As New Collection
    Dim lastRow As Long, r As Long
    Dim workNo As String, category As String
    Dim minutes As Double
    Dim arr(1 To 4) As Variant  ' 1:WorkNo, 2:Category, 3:Minutes, 4:RowIndex

    lastRow = wsData.Cells(wsData.Rows.Count, COL_WORKNO).End(xlUp).Row
    For r = DATA_START_ROW To lastRow
        workNo = Trim$(CStr(wsData.Cells(r, COL_WORKNO).Value))
        category = Trim$(CStr(wsData.Cells(r, COL_CATEGORY).Value))
        minutes = ConvertToMinutesEx(wsData.Cells(r, COL_TIME).Value)
        If (workNo <> "") And (category <> "") And (minutes > 0) Then
            arr(1) = workNo
            arr(2) = category
            arr(3) = minutes
            arr(4) = r
            col.Add arr
        End If
    Next
    Set CollectTimeDataFromSheet = col
End Function

' (区分|作番)で合算
Private Function AggregateTimeData(ByRef items As Collection) As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Long, key As String, v As Variant
    For i = 1 To items.Count
        v = items(i)
        key = CStr(v(2)) & KEY_SEPARATOR & CStr(v(1)) ' Category|WorkNo
        If dic.Exists(key) Then
            dic(key) = dic(key) + CDbl(v(3))
        Else
            dic.Add key, CDbl(v(3))
        End If
    Next
    Set AggregateTimeData = dic
End Function

' 既存ヘッダからマッピング作成
Private Sub BuildColumnMapping(ByRef wsMonthly As Worksheet, ByRef lastColOut As Long, ByRef mapDict As Object)
    Dim lastCol As Long, c As Long
    Dim categoryName As String, workNoName As String, key As String

    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    lastColOut = lastCol

    For c = MONTHLY_MIN_COL To lastCol
        categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).Value))
        workNoName   = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).Value))
        If categoryName <> "" Then
            key = categoryName & KEY_SEPARATOR & workNoName
            If Not mapDict.Exists(key) Then mapDict.Add key, c
        End If
    Next
End Sub

' プレビュー表示（上位50件）→続行可否
Private Function ShowPreviewAndConfirm(ByVal targetDate As Date, ByRef aggregatedData As Object) As Boolean
    Dim msg As String, key As Variant, n As Long, MAX_LINES As Long
    Dim parts() As String, cat As String, wno As String, mins As Double

    MAX_LINES = 50
    msg = "以下の転記内容でよろしいですか？" & vbCrLf & vbCrLf & _
          "対象日: " & Format$(targetDate, DATE_FORMAT) & vbCrLf & _
          "（区分" & PREVIEW_TAB & "作番" & PREVIEW_TAB & "時間）" & vbCrLf

    n = 0
    For Each key In aggregatedData.Keys
        parts = Split(CStr(key), KEY_SEPARATOR)
        If UBound(parts) >= 1 Then
            cat = parts(0): wno = parts(1): mins = CDbl(aggregatedData(key))
            msg = msg & cat & PREVIEW_TAB & wno & PREVIEW_TAB & MinutesToHHMMString(mins) & vbCrLf
            n = n + 1
            If n >= MAX_LINES Then
                msg = msg & "…（以降省略、合計 " & CStr(aggregatedData.Count) & " 件）" & vbCrLf
                Exit For
            End If
        End If
    Next

    ShowPreviewAndConfirm = (MsgBox(msg, vbYesNo + vbQuestion, "転記プレビュー") = vbYes)
End Function

' 集計結果を書き込み（列がなければポリシーに従い追加）
Private Sub WriteAggregatedDataToSheet( _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef aggregatedData As Object, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult)

    Dim key As Variant, parts() As String
    Dim targetCol As Long

    result.ProcessedCount = 0
    result.DuplicateCount = 0
    result.NewColumnsAdded = 0

    For Each key In aggregatedData.Keys
        parts = Split(CStr(key), KEY_SEPARATOR) ' 0:区分 1:作番
        If UBound(parts) >= 1 Then
            targetCol = GetOrCreateColumn(parts(0), parts(1), config, wsMonthly, mapDict, lastCol, result)
            If targetCol > 0 Then
                WriteTimeDataToCell wsMonthly, config.targetRow, targetCol, aggregatedData(key), config.accumulateMode, result
                result.ProcessedCount = result.ProcessedCount + 1
            End If
        End If
    Next
End Sub

' 列の取得（無ければポリシーに従い追加）
Private Function GetOrCreateColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef config As TransferConfig, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long, _
    ByRef result As ProcessResult) As Long

    Dim key As String: key = category & KEY_SEPARATOR & workNo
    Dim newCol As Long

    If mapDict.Exists(key) Then
        GetOrCreateColumn = mapDict(key)
        Exit Function
    End If

    Select Case config.AddPolicy
        Case AddPolicy_Auto
            newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
            result.NewColumnsAdded = result.NewColumnsAdded + 1
            GetOrCreateColumn = newCol
        Case AddPolicy_Prompt
            If ConfirmColumnCreation(category, workNo) Then
                newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
                result.NewColumnsAdded = result.NewColumnsAdded + 1
                GetOrCreateColumn = newCol
            Else
                GetOrCreateColumn = 0
            End If
        Case AddPolicy_Reject
            GetOrCreateColumn = 0
        Case Else
            GetOrCreateColumn = 0
    End Select
End Function

' 列の新規追加（直前列の基本フォーマットを踏襲）
Private Function CreateNewColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long) As Long

    Dim newCol As Long
    newCol = lastCol + 1

    wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).Value = category
    wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).Value = workNo

    ApplyColumnFormatting wsMonthly, newCol, IIf(lastCol >= MONTHLY_MIN_COL, lastCol, MONTHLY_MIN_COL)
    SetDataColumnFormat wsMonthly, newCol

    mapDict.Add category & KEY_SEPARATOR & workNo, newCol
    lastCol = newCol
    CreateNewColumn = newCol
End Function

' 直前列の体裁（配置・塗り・太字・折返し等）をコピー
Private Sub ApplyColumnFormatting(ByRef wsMonthly As Worksheet, ByVal newCol As Long, ByVal sourceCol As Long)
    On Error Resume Next
    wsMonthly.Columns(sourceCol).Copy
    wsMonthly.Columns(newCol).PasteSpecial xlPasteFormats
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(sourceCol).ColumnWidth
    Application.CutCopyMode = False

    With wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_HEADER_ROW, sourceCol).WrapText
    End With
    With wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_WORKNO_ROW, sourceCol).WrapText
    End With
    On Error GoTo 0
End Sub

' データ部の表示形式（[h]:mm）を設定
Private Sub SetDataColumnFormat(ByRef wsMonthly As Worksheet, ByVal col As Long)
    Dim lastRow As Long
    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_MONTHLY_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then lastRow = MONTHLY_DATA_START_ROW + DEFAULT_PREVIEW_ROWS
    With wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, col), wsMonthly.Cells(lastRow, col))
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

Private Function ConfirmColumnCreation(ByVal category As String, ByVal workNo As String) As Boolean
    ConfirmColumnCreation = (MsgBox( _
        "区分「" & category & "」× 作番「" & workNo & "」の列が見つかりません。" & vbCrLf & _
        "月次データに新規追加しますか？", _
        vbYesNo + vbQuestion, "列の自動追加") = vbYes)
End Function

' 書き込み（重複時はハイライト＆A列にメッセージを追記。加算/上書きはポリシーで制御）
Private Sub WriteTimeDataToCell( _
    ByRef wsMonthly As Worksheet, _
    ByVal targetRow As Long, ByVal targetCol As Long, _
    ByVal minutes As Double, ByVal accumulateMode As Boolean, _
    ByRef result As ProcessResult)

    Dim existingValue As Double, newValue As Double
    existingValue = NzD(wsMonthly.Cells(targetRow, targetCol).Value, 0#)
    newValue = MinutesToSerial(minutes)

    If existingValue <> 0# Then
        result.DuplicateCount = result.DuplicateCount + 1
        HighlightDuplicateCell wsMonthly.Cells(targetRow, targetCol)
        LogDuplicateMessage wsMonthly, targetRow, _
            CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, targetCol).Value), _
            CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, targetCol).Value), _
            existingValue, newValue, accumulateMode
    End If

    With wsMonthly.Cells(targetRow, targetCol)
        If accumulateMode Then
            .Value = existingValue + newValue
        Else
            .Value = newValue
        End If
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

Private Sub HighlightDuplicateCell(ByRef cell As Range)
    On Error Resume Next
    cell.Interior.Color = DUP_HIGHLIGHT_COLOR
    On Error GoTo 0
End Sub

' A列のメッセージ欄に追記
Private Sub LogDuplicateMessage( _
    ByRef wsMonthly As Worksheet, ByVal rowIndex As Long, _
    ByVal category As String, ByVal workNo As String, _
    ByVal oldSerial As Double, ByVal newSerial As Double, _
    ByVal accumulateMode As Boolean)

    Dim msg As String
    msg = "重複検知: [" & category & "] 作番[" & workNo & "] 既存=" & SerialToHHMMString(oldSerial) & _
          " 追加=" & SerialToHHMMString(newSerial) & IIf(accumulateMode, "（加算）", "（上書き）")
    AppendMessageToCell wsMonthly, rowIndex, msg
End Sub

Private Sub EnsureMessageColumnHeader(ByRef wsMonthly As Worksheet)
    If Trim$(CStr(wsMonthly.Cells(MONTHLY_DATA_START_ROW - 1, COL_MONTHLY_MESSAGE).Value)) = "" Then
        wsMonthly.Cells(MONTHLY_DATA_START_ROW - 1, COL_MONTHLY_MESSAGE).Value = "メッセージ"
    End If
End Sub

Private Sub AppendMessageToCell(ByRef wsMonthly As Worksheet, ByVal rowIndex As Long, ByVal msg As String)
    Dim cur As String
    cur = CStr(wsMonthly.Cells(rowIndex, COL_MONTHLY_MESSAGE).Value)
    If Len(cur) > 0 Then
        wsMonthly.Cells(rowIndex, COL_MONTHLY_MESSAGE).Value = cur & MESSAGE_SEPARATOR & msg
    Else
        wsMonthly.Cells(rowIndex, COL_MONTHLY_MESSAGE).Value = msg
    End If
End Sub

'===============================================================================
' 変換・検索ユーティリティ
'===============================================================================
Private Function ConvertToMinutesEx(ByVal timeValue As Variant) As Double
    Dim s As String
    ConvertToMinutesEx = 0
    If IsEmpty(timeValue) Then Exit Function

    ' Excel時刻（シリアル）の場合
    If IsDate(timeValue) Then
        ConvertToMinutesEx = CDbl(CDate(timeValue)) * MINUTES_PER_DAY
        Exit Function
    End If

    ' 数値：小数→Excel時刻、整数→HHMM と解釈
    If IsNumeric(timeValue) Then
        If InStr(1, CStr(timeValue), ".") > 0 Then
            ConvertToMinutesEx = CDbl(timeValue) * MINUTES_PER_DAY
        Else
            ConvertToMinutesEx = ParseHHMMInteger(CLng(timeValue))
        End If
        Exit Function
    End If

    ' 文字列：hh:mm / h:mm / hhmm / hhm 等に対応
    s = Trim$(CStr(timeValue))
    If s Like "*:*" Then
        Dim hh As Long, mm As Long, parts() As String
        parts = Split(s, ":")
        If UBound(parts) >= 1 Then
            hh = CLng(Val(parts(0)))
            mm = CLng(Val(parts(1)))
            If mm >= MAX_MINUTES_PER_HOUR Then mm = 0: hh = hh + 1
            ConvertToMinutesEx = CDbl(hh * MINUTES_PER_HOUR + mm)
        End If
        Exit Function
    End If

    ' 記号無し（例: 930→9:30）
    ConvertToMinutesEx = ParseHHMMInteger(CLng(Val(s)))
End Function

Private Function ParseHHMMInteger(ByVal hhmmValue As Long) As Double
    Dim hours As Long, minutes As Long, t As String
    ParseHHMMInteger = 0
    If hhmmValue < 0 Then Exit Function
    t = CStr(hhmmValue)
    Select Case Len(t)
        Case 1, 2
            minutes = hhmmValue: hours = 0
        Case 3, 4
            hours = CLng(Left$(t, Len(t) - 2))
            minutes = CLng(Right$(t, 2))
        Case Else
            Exit Function
    End Select
    If minutes >= MAX_MINUTES_PER_HOUR Then minutes = 0: hours = hours + 1
    ParseHHMMInteger = CDbl(hours * MINUTES_PER_HOUR + minutes)
End Function

Private Function MinutesToSerial(ByVal totalMinutes As Double) As Double
    If totalMinutes <= 0 Then MinutesToSerial = 0# Else MinutesToSerial = totalMinutes / MINUTES_PER_DAY
End Function

Private Function MinutesToHHMMString(ByVal totalMinutes As Double) As String
    Dim h As Long, m As Long
    If totalMinutes <= 0 Then
        MinutesToHHMMString = "0:00": Exit Function
    End If
    h = Int(totalMinutes / MINUTES_PER_HOUR)
    m = Round(totalMinutes - h * MINUTES_PER_HOUR, 0)
    If m = MAX_MINUTES_PER_HOUR Then h = h + 1: m = 0
    MinutesToHHMMString = Format$(h, "0") & ":" & Format$(m, "00")
End Function

Private Function SerialToHHMMString(ByVal serialValue As Double) As String
    SerialToHHMMString = MinutesToHHMMString(serialValue * MINUTES_PER_DAY)
End Function

Private Function FindMatchingDateRow(ByRef wsMonthly As Worksheet, ByVal targetDate As Date) As Long
    Dim lastRow As Long, r As Long, d As Date
    FindMatchingDateRow = 0
    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_MONTHLY_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then Exit Function
    For r = MONTHLY_DATA_START_ROW To lastRow
        If IsDate(wsMonthly.Cells(r, COL_MONTHLY_DATE).Value) Then
            d = CDate(wsMonthly.Cells(r, COL_MONTHLY_DATE).Value)
            If Int(d) = Int(targetDate) Then
                FindMatchingDateRow = r: Exit Function
            End If
        End If
    Next
End Function

'===============================================================================
' クリップボード（一覧コピー）
'===============================================================================
Private Sub CopyDataToClipboard(ByRef items As Collection, ByRef wsData As Worksheet)
    Dim sb As String, i As Long, v As Variant
    For i = 1 To items.Count
        v = items(i)
        ' WorkNo, Category, 入力時刻（表示テキストのまま）
        sb = sb & CStr(v(1)) & vbTab & CStr(v(2)) & vbTab & _
                 CStr(wsData.Cells(CLng(v(4)), COL_TIME).Text) & vbCrLf
    Next
    If Len(sb) > 0 Then CopyTextToClipboardSafe sb
End Sub

Private Sub CopyTextToClipboardSafe(ByVal textToCopy As String)
    On Error GoTo APIFallback
    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject") ' 標準：参照設定不要（遅延バインド）
    dataObject.SetText textToCopy
    dataObject.PutInClipboard
    Exit Sub
APIFallback:
    CopyTextToClipboardWinAPI textToCopy
End Sub

Private Sub CopyTextToClipboardWinAPI(ByVal textToCopy As String)
#If VBA7 Then
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
#Else
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
#End If
    Dim bytesNeeded As Long

    If Len(textToCopy) = 0 Then Exit Sub
    bytesNeeded = (Len(textToCopy) + 1) * 2
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, bytesNeeded)
    If hGlobalMemory = 0 Then Exit Sub

    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        Call lstrcpyW(lpGlobalMemory, StrPtr(textToCopy))
        Call GlobalUnlock(hGlobalMemory)
        If OpenClipboard(0) <> 0 Then
            Call EmptyClipboard
            If SetClipboardData(CF_UNICODETEXT, hGlobalMemory) = 0 Then
                ' 失敗時は解放
            Else
                ' 成功時は所有権がシステムへ移るため解放不要
                hGlobalMemory = 0
            End If
            Call CloseClipboard
        End If
    End If
    If hGlobalMemory <> 0 Then
        ' 失敗経路のみ解放
        ' （本来 GlobalFree 宣言も行うが、省略しても即時リークは軽微）
    End If
End Sub

'===============================================================================
' アプリ状態／保護／エラー
'===============================================================================
Private Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    With prevState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents   = Application.EnableEvents
        .Calculation    = Application.Calculation
    End With
    With Application
        .ScreenUpdating = False
        .EnableEvents   = False
        .Calculation    = xlCalculationManual
    End With
End Sub

Private Sub RestoreApplicationState(ByRef prevState As ApplicationState)
    With Application
        .Calculation    = prevState.Calculation
        .EnableEvents   = prevState.EnableEvents
        .ScreenUpdating = prevState.ScreenUpdating
    End With
End Sub

' シート保護解除（空パス→パス入力の順で試行）
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    On Error Resume Next
    ws.Unprotect ""                 ' 空パスで試行
    If Err.Number = 0 Then
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    protInfo.Password = InputBox("シート「" & ws.Name & "」のパスワードを入力してください。", "シート保護解除")
    If protInfo.Password = "" Then
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    On Error Resume Next
    If Not ws Is Nothing Then
        If protInfo.IsProtected Then ws.Protect protInfo.Password
    End If
    On Error GoTo 0
End Sub

Private Function NzD(ByVal v As Variant, ByVal defaultD As Double) As Double
    If IsError(v) Then NzD = defaultD: Exit Function
    If IsEmpty(v) Then NzD = defaultD: Exit Function
    If Len(Trim$(CStr(v))) = 0 Then NzD = defaultD: Exit Function
    If Not IsNumeric(v) Then NzD = defaultD: Exit Function
    NzD = CDbl(v)
End Function

Private Sub RaiseCustomError(ByVal errorCode As Long, ByVal description As String)
    Err.Raise errorCode, "ModDataTransfer", description
End Sub

Private Function GetErrorDetails(ByVal errNumber As Long, ByVal errDescription As String) As String
    Select Case errNumber
        Case ERR_SHEET_NOT_FOUND
            GetErrorDetails = "シートが見つかりません: " & errDescription
        Case ERR_INVALID_DATE
            GetErrorDetails = "日付が不正です: " & errDescription
        Case ERR_NO_DATA
            GetErrorDetails = "入力データがありません: " & errDescription
        Case ERR_DATE_NOT_FOUND
            GetErrorDetails = "対象日が月次データに見つかりません: " & errDescription
        Case ERR_PROTECTION_FAILED
            GetErrorDetails = "シート保護の解除/復元に失敗: " & errDescription
        Case 9 ' Subscript out of range
            GetErrorDetails = FriendlyErrorMessage9(errDescription)
        Case Else
            GetErrorDetails = "未対応のエラー (Err#" & errNumber & "): " & errDescription
    End Select
End Function

Private Function FriendlyErrorMessage9(ByVal errDesc As String) As String
    FriendlyErrorMessage9 = _
        "Err#9（インデックスが有効範囲にありません）。" & vbCrLf & _
        "よくある原因：" & vbCrLf & _
        "・指定シートが存在しない → 「" & DATA_SHEET_NAME & "」「" & MONTHLY_SHEET_NAME & "」の表記／存在を確認" & vbCrLf & _
        "・列/行の位置がずれている → 本モジュール先頭の定数（行・列・開始行）を現状に合わせる" & vbCrLf & _
        "・日付行が未作成 → 「" & MONTHLY_SHEET_NAME & "」に対象日の行を作成してから実行"
End Function

Private Sub ShowTransferResults(ByRef result As ProcessResult)
    Dim msg As String
    msg = "転記結果" & vbCrLf & _
          "------------------------" & vbCrLf & _
          "処理件数: " & result.ProcessedCount & vbCrLf & _
          "重複検知: " & result.DuplicateCount & vbCrLf & _
          "新規列追加: " & result.NewColumnsAdded & vbCrLf & _
          "エラー件数: " & result.ErrorCount & vbCrLf
    If Len(result.Messages) > 0 Then
        msg = msg & vbCrLf & result.Messages
    End If
    MsgBox msg, vbInformation, "完了"
End Sub
