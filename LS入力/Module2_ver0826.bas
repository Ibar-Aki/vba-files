Option Explicit
'===============================================================================
' モジュール名: TransferDataModule.bas
' 機能: 「データ登録」→「月次データ」へ転記（区分＋作番で列特定）
' 対象: Excel 2016+ / Windows 11 / 日本語環境
'===============================================================================

'=========================
' WinAPI（64/32両対応）
'=========================
#If VBA7 Then
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

Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_UNICODETEXT As Long = 13

'=========================
' システム設定定数
'=========================
' シート名
Private Const DATA_SHEET_NAME        As String = "データ登録"
Private Const MONTHLY_SHEET_NAME     As String = "月次データ"

' セル位置
Private Const DATE_CELL_PRIORITY As String = "D4"
Private Const DATE_CELL_NORMAL   As String = "D3"

' 行・列番号
Private Const DATA_START_ROW     As Long = 8
Private Const MONTHLY_WORKNO_ROW As Long = 8
Private Const MONTHLY_HEADER_ROW As Long = 9
Private Const MONTHLY_DATA_START_ROW As Long = 10
Private Const MONTHLY_MIN_COL    As Long = 3   ' C列～

' 列（数値インデックス）
Private Const COL_MESSAGE  As Long = 1   ' A
Private Const COL_DATE     As Long = 2   ' B
Private Const COL_WORKNO   As Long = 3   ' C
Private Const COL_CATEGORY As Long = 4   ' D
Private Const COL_TIME     As Long = 5   ' E

' 計算定数
Private Const MINUTES_PER_HOUR   As Double = 60#
Private Const MINUTES_PER_DAY    As Double = 1440#
Private Const MAX_MINUTES_PER_HOUR As Long = 60
Private Const MAX_RETRY_CLIPBOARD  As Long = 5
Private Const DEFAULT_PREVIEW_ROWS As Long = 31

' 文字列定数
Private Const KEY_SEPARATOR As String = "|"
Private Const MESSAGE_SEPARATOR As String = vbLf
Private Const TIME_FORMAT As String = "[hh]mm"
Private Const DATE_FORMAT As String = "yyyy/mm/dd"
Private Const PREVIEW_TAB As String = vbTab

' 列追加ポリシー
Private Const AddPolicy_Prompt As Long = 0
Private Const AddPolicy_Auto   As Long = 1
Private Const AddPolicy_Reject As Long = 2

' 動作設定
Private Const AUTO_ADD_POLICY As Long = AddPolicy_Prompt
Private Const ACCUMULATE_MODE As Boolean = True
Private Const DRY_RUN         As Boolean = False
Private Const DUP_HIGHLIGHT_COLOR As Long = vbYellow

'=========================
' カスタムエラー定数
'=========================
Private Const ERR_SHEET_NOT_FOUND   As Long = vbObjectError + 1
Private Const ERR_INVALID_DATE      As Long = vbObjectError + 2
Private Const ERR_NO_DATA           As Long = vbObjectError + 3
Private Const ERR_DATE_NOT_FOUND    As Long = vbObjectError + 4
Private Const ERR_PROTECTION_FAILED As Long = vbObjectError + 5

'=========================
' データ構造（宣言セクション）
'=========================
Private Type ApplicationState
    ScreenUpdating As Boolean
    EnableEvents   As Boolean
    Calculation    As Long   ' XlCalculation を数値で保持
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
' メイン処理
'===============================================================================
Public Sub TransferDataToMonthlySheet()
    Dim prevState As ApplicationState
    Dim config As TransferConfig
    Dim result As ProcessResult
    Dim protectionInfo As SheetProtectionInfo
    Dim wsData As Worksheet, wsMonthly As Worksheet

    SaveAndSetApplicationState prevState
    On Error GoTo ErrorHandler

    ' 初期化と検証（シート返却）
    If Not InitializeTransferConfig(config, protectionInfo, wsData, wsMonthly) Then GoTo CleanUp

    ' データ処理実行
    ExecuteDataTransfer config, wsData, wsMonthly, result

    ' 結果表示
    ShowTransferResults result

CleanUp:
    RestoreSheetProtection wsMonthly, protectionInfo
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    MsgBox GetErrorDetails(Err.Number, Err.description), vbCritical, "転記処理エラー"
    Resume CleanUp
End Sub

'===============================================================================
' 初期化と設定
'===============================================================================
Private Function InitializeTransferConfig( _
    ByRef config As TransferConfig, _
    ByRef protInfo As SheetProtectionInfo, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet) As Boolean

    InitializeTransferConfig = False

    If Not GetAndValidateWorksheets(wsData, wsMonthly) Then Exit Function

    ' シート保護の解除
    If Not UnprotectSheetIfNeeded(wsMonthly, protInfo) Then Exit Function

    ' 日付の決定
    If Not DetermineTargetDate(wsData, config.targetDate) Then Exit Function

    ' 対象行の取得
    config.targetRow = FindMatchingDateRow(wsMonthly, config.targetDate)
    If config.targetRow = 0 Then
        RaiseCustomError ERR_DATE_NOT_FOUND, Format$(config.targetDate, DATE_FORMAT)
        Exit Function
    End If

    ' その他設定
    config.accumulateMode = ACCUMULATE_MODE
    config.DryRun = DRY_RUN
    config.AddPolicy = AUTO_ADD_POLICY

    InitializeTransferConfig = True
End Function

Private Function GetAndValidateWorksheets(ByRef wsData As Worksheet, ByRef wsMonthly As Worksheet) As Boolean
    On Error GoTo SheetError

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)

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
    ' データシートの構造チェック
    If wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row < DATA_START_ROW Then
        ValidateSheetStructure = False: Exit Function
    End If
    ' 月次シートの構造チェック（ヘッダ行に最低C列まである）
    If wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column < MONTHLY_MIN_COL Then
        ValidateSheetStructure = False: Exit Function
    End If
    ValidateSheetStructure = True
End Function

Private Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    DetermineTargetDate = False
    If IsDate(wsData.Range(DATE_CELL_PRIORITY).value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_PRIORITY).value)
        DetermineTargetDate = True
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).value)
        DetermineTargetDate = True
    Else
        RaiseCustomError ERR_INVALID_DATE, "セル " & DATE_CELL_NORMAL & " または " & DATE_CELL_PRIORITY
    End If
End Function

'===============================================================================
' データ処理の実行
'===============================================================================
Private Sub ExecuteDataTransfer( _
    ByRef config As TransferConfig, _
    ByRef wsData As Worksheet, _
    ByRef wsMonthly As Worksheet, _
    ByRef result As ProcessResult)

    Dim items As collection              ' 各行: Array(WorkNo, Category, Minutes, RowIndex)
    Dim aggregated As Object             ' Scripting.Dictionary (key="区分|作番", val=合計分)
    Dim mapDict As Object                ' 列マッピング辞書 (key="区分|作番", val=列番号)
    Dim lastCol As Long

    ' データ収集
    Set items = CollectTimeDataFromSheet(wsData)
    If items.Count = 0 Then
        RaiseCustomError ERR_NO_DATA, "有効な時間データが見つかりません"
        Exit Sub
    End If

    ' データ集計
    Set aggregated = AggregateTimeData(items)

    ' 列マッピングの構築
    Set mapDict = CreateObject("Scripting.Dictionary")
    BuildColumnMapping wsMonthly, lastCol, mapDict

    ' プレビュー表示と確認
    If Not ShowPreviewAndConfirm(config.targetDate, aggregated) Then
        result.Success = False
        Exit Sub
    End If

    ' ドライラン
    If config.DryRun Then
        result.Messages = "ドライラン完了（実際の書き込みは実行されませんでした）"
        result.Success = True
        Exit Sub
    End If

    ' クリップボードへ（表形式）
    CopyDataToClipboard items, wsData

    ' メッセージ列ヘッダ
    EnsureMessageColumnHeader wsMonthly

    ' 書き込み
    WriteAggregatedDataToSheet config, wsMonthly, aggregated, mapDict, lastCol, result

    result.Success = True
End Sub

'===============================================================================
' データ収集（Collection of Variant()）
'===============================================================================
Private Function CollectTimeDataFromSheet(ByRef wsData As Worksheet) As collection
    Dim col As New collection
    Dim lastRow As Long, r As Long
    Dim workNo As String, category As String
    Dim minutes As Double
    Dim arr(1 To 4) As Variant  ' 1:WorkNo, 2:Category, 3:Minutes, 4:RowIndex

    lastRow = wsData.Cells(wsData.rows.Count, COL_WORKNO).End(xlUp).Row
    For r = DATA_START_ROW To lastRow
        workNo = Trim$(CStr(wsData.Cells(r, COL_WORKNO).value))
        category = Trim$(CStr(wsData.Cells(r, COL_CATEGORY).value))
        minutes = ConvertToMinutesEx(wsData.Cells(r, COL_TIME).value)
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

'===============================================================================
' 集計（key="区分|作番"）
'===============================================================================
Private Function AggregateTimeData(ByRef items As collection) As Object
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

'===============================================================================
' 列マッピング構築
'===============================================================================
Private Sub BuildColumnMapping(ByRef wsMonthly As Worksheet, ByRef lastColOut As Long, ByRef mapDict As Object)
    Dim lastCol As Long, c As Long
    Dim categoryName As String, workNoName As String, key As String

    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    lastColOut = lastCol

    For c = MONTHLY_MIN_COL To lastCol
        categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).value))
        workNoName = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).value))
        If categoryName <> "" Then
            key = categoryName & KEY_SEPARATOR & workNoName
            If Not mapDict.Exists(key) Then mapDict.Add key, c
        End If
    Next
End Sub

'===============================================================================
' プレビュー
'===============================================================================
Private Function ShowPreviewAndConfirm(ByVal targetDate As Date, ByRef aggregatedData As Object) As Boolean
    Dim msg As String, key As Variant, n As Long, MAX_LINES As Long
    MAX_LINES = 50

    msg = "以下の内容で転記します。よろしいですか？" & vbCrLf & vbCrLf & _
          "対象日付: " & Format$(targetDate, DATE_FORMAT) & vbCrLf & _
          String(50, "-") & vbCrLf & _
          "作番" & PREVIEW_TAB & " | 区分" & PREVIEW_TAB & " | 時間" & vbCrLf & _
          String(50, "-") & vbCrLf

    n = 0
    For Each key In aggregatedData.Keys
        n = n + 1
        If n <= MAX_LINES Then
            Dim parts() As String
            parts = Split(CStr(key), KEY_SEPARATOR)
            If UBound(parts) >= 1 Then
                msg = msg & parts(1) & PREVIEW_TAB & " | " & parts(0) & PREVIEW_TAB & _
                      " | " & MinutesToHHMMString(aggregatedData(key)) & vbCrLf
            End If
        Else
            msg = msg & "…ほか " & (aggregatedData.Count - MAX_LINES) & " 件" & vbCrLf
            Exit For
        End If
    Next
    ShowPreviewAndConfirm = (MsgBox(msg, vbYesNo + vbQuestion, "転記内容の確認") = vbYes)
End Function

'===============================================================================
' データ書き込み
'===============================================================================
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
        Case AddPolicy_Reject
            GetOrCreateColumn = 0
        Case AddPolicy_Auto
            newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
            If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
            GetOrCreateColumn = newCol
        Case Else ' Prompt
            If ConfirmColumnCreation(category, workNo) Then
                newCol = CreateNewColumn(category, workNo, wsMonthly, mapDict, lastCol)
                If newCol > 0 Then result.NewColumnsAdded = result.NewColumnsAdded + 1
                GetOrCreateColumn = newCol
            Else
                GetOrCreateColumn = 0
            End If
    End Select
End Function

Private Function CreateNewColumn( _
    ByVal category As String, ByVal workNo As String, _
    ByRef wsMonthly As Worksheet, _
    ByRef mapDict As Object, _
    ByRef lastCol As Long) As Long

    Dim newCol As Long
    newCol = lastCol + 1

    wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).value = category
    wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).value = workNo

    ApplyColumnFormatting wsMonthly, newCol, IIf(lastCol >= MONTHLY_MIN_COL, lastCol, MONTHLY_MIN_COL)
    SetDataColumnFormat wsMonthly, newCol

    mapDict.Add category & KEY_SEPARATOR & workNo, newCol
    lastCol = newCol

    CreateNewColumn = newCol
End Function

Private Sub ApplyColumnFormatting(ByRef wsMonthly As Worksheet, ByVal newCol As Long, ByVal sourceCol As Long)
    On Error Resume Next
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(sourceCol).ColumnWidth
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

Private Sub SetDataColumnFormat(ByRef wsMonthly As Worksheet, ByVal col As Long)
    Dim lastRow As Long
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then lastRow = MONTHLY_DATA_START_ROW + DEFAULT_PREVIEW_ROWS
    With wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, col), wsMonthly.Cells(lastRow, col))
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

Private Function ConfirmColumnCreation(ByVal category As String, ByVal workNo As String) As Boolean
    ConfirmColumnCreation = (MsgBox( _
        "区分『" & category & "』＋作番『" & workNo & "』の列がありません。" & vbCrLf & _
        "月次データシートに新しい列を追加しますか？", _
        vbYesNo + vbQuestion, "列の追加確認") = vbYes)
End Function

Private Sub WriteTimeDataToCell( _
    ByRef wsMonthly As Worksheet, _
    ByVal targetRow As Long, ByVal targetCol As Long, _
    ByVal minutes As Double, ByVal accumulateMode As Boolean, _
    ByRef result As ProcessResult)

    Dim existingValue As Double, newValue As Double
    existingValue = NzD(wsMonthly.Cells(targetRow, targetCol).value, 0#)
    newValue = MinutesToSerial(minutes)

    If existingValue <> 0# Then
        result.DuplicateCount = result.DuplicateCount + 1
        HighlightDuplicateCell wsMonthly.Cells(targetRow, targetCol)
        LogDuplicateMessage wsMonthly, targetRow, _
            CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, targetCol).value), _
            CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, targetCol).value), _
            existingValue, newValue, accumulateMode
    End If

    With wsMonthly.Cells(targetRow, targetCol)
        If accumulateMode Then
            .value = existingValue + newValue
        Else
            .value = newValue
        End If
        .NumberFormatLocal = TIME_FORMAT
    End With
End Sub

Private Sub HighlightDuplicateCell(ByRef cell As Range)
    With cell.Interior
        .Pattern = xlSolid
        .Color = DUP_HIGHLIGHT_COLOR
    End With
End Sub

Private Sub LogDuplicateMessage( _
    ByRef wsMonthly As Worksheet, ByVal rowNum As Long, _
    ByVal category As String, ByVal workNo As String, _
    ByVal oldValue As Double, ByVal newValue As Double, _
    ByVal accumulateMode As Boolean)

    Dim message As String
    message = "既存値検出: [" & workNo & "|" & category & "] 旧=" & SerialToHHMMString(oldValue) & _
              " 新=" & SerialToHHMMString(newValue) & IIf(accumulateMode, " (加算)", " (上書)")
    AppendMessageToCell wsMonthly, rowNum, message
End Sub

'===============================================================================
' クリップボード
'===============================================================================
Private Sub CopyDataToClipboard(ByRef items As collection, ByRef wsData As Worksheet)
    Dim sb As String, i As Long, v As Variant
    For i = 1 To items.Count
        v = items(i)
        ' WorkNo, Category, 表示文字列としての時間
        sb = sb & CStr(v(1)) & vbTab & CStr(v(2)) & vbTab & _
                 CStr(wsData.Cells(CLng(v(4)), COL_TIME).text) & vbCrLf
    Next
    If Len(sb) > 0 Then CopyTextToClipboardSafe sb
End Sub

Private Sub CopyTextToClipboardSafe(ByVal textToCopy As String)
    On Error GoTo APIFallback
    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject") ' 参照設定不要／無い環境もある
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
    Dim bytesNeeded As Long, retryCount As Long

    If Len(textToCopy) = 0 Then Exit Sub
    bytesNeeded = (Len(textToCopy) + 1) * 2
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, bytesNeeded)
    If hGlobalMemory = 0 Then Exit Sub

    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        lstrcpyW lpGlobalMemory, StrPtr(textToCopy)
        GlobalUnlock hGlobalMemory

        For retryCount = 1 To MAX_RETRY_CLIPBOARD
            If OpenClipboard(0) <> 0 Then Exit For
            DoEvents
        Next retryCount

        If retryCount <= MAX_RETRY_CLIPBOARD Then
            EmptyClipboard
            If SetClipboardData(CF_UNICODETEXT, hGlobalMemory) = 0 Then
                GlobalFree hGlobalMemory
            End If
            CloseClipboard
        Else
            GlobalFree hGlobalMemory
        End If
    Else
        GlobalFree hGlobalMemory
    End If
End Sub

'===============================================================================
' ユーティリティ
'===============================================================================
Private Function ConvertToMinutesEx(ByVal timeValue As Variant) As Double
    Dim s As String
    ConvertToMinutesEx = 0
    If IsEmpty(timeValue) Then Exit Function

    If IsDate(timeValue) Then
        ConvertToMinutesEx = CDbl(CDate(timeValue)) * MINUTES_PER_DAY
        Exit Function
    End If

    If IsNumeric(timeValue) Then
        If InStr(1, CStr(timeValue), ".") > 0 Then
            ConvertToMinutesEx = CDbl(timeValue) * MINUTES_PER_DAY   ' シリアル
        Else
            ConvertToMinutesEx = ParseHHMMInteger(CLng(timeValue))   ' HHMM
        End If
        Exit Function
    End If

    s = Trim$(CStr(timeValue))
    If InStr(s, ":") > 0 Then
        ConvertToMinutesEx = ParseHHMMString(s)
    ElseIf IsNumeric(s) Then
        ConvertToMinutesEx = ParseHHMMInteger(CLng(Val(s)))
    End If
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
    If minutes >= 0 And minutes < MAX_MINUTES_PER_HOUR Then
        ParseHHMMInteger = hours * MINUTES_PER_HOUR + minutes
    End If
End Function

Private Function ParseHHMMString(ByVal timeString As String) As Double
    Dim parts() As String, h As Long, m As Long
    ParseHHMMString = 0
    parts = Split(timeString, ":")
    If UBound(parts) = 1 Then
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
            h = CLng(parts(0)): m = CLng(parts(1))
            If m >= 0 And m < MAX_MINUTES_PER_HOUR Then
                ParseHHMMString = h * MINUTES_PER_HOUR + m
            End If
        End If
    End If
End Function

Private Function MinutesToSerial(ByVal totalMinutes As Double) As Double
    MinutesToSerial = totalMinutes / MINUTES_PER_DAY
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
    lastRow = wsMonthly.Cells(wsMonthly.rows.Count, COL_DATE).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then Exit Function
    For r = MONTHLY_DATA_START_ROW To lastRow
        If IsDate(wsMonthly.Cells(r, COL_DATE).value) Then
            d = CDate(wsMonthly.Cells(r, COL_DATE).value)
            If Int(d) = Int(targetDate) Then
                FindMatchingDateRow = r: Exit Function
            End If
        End If
    Next
End Function

Private Sub EnsureMessageColumnHeader(ByRef wsMonthly As Worksheet)
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, COL_MESSAGE)
        If Trim$(CStr(.value)) = "" Then
            .value = "メッセージ"
            .Font.Bold = True
        End If
    End With
End Sub

Private Sub AppendMessageToCell(ByRef wsMonthly As Worksheet, ByVal rowNum As Long, ByVal message As String)
    With wsMonthly.Cells(rowNum, COL_MESSAGE)
        If Len(.value) = 0 Then
            .value = message
        Else
            .value = CStr(.value) & MESSAGE_SEPARATOR & message
        End If
    End With
End Sub

Private Function NzD(ByVal value As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    On Error Resume Next
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
' アプリケーション状態管理
'===============================================================================
Private Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    With prevState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
    End With
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub RestoreApplicationState(ByRef prevState As ApplicationState)
    With Application
        .Calculation = prevState.Calculation
        .EnableEvents = prevState.EnableEvents
        .ScreenUpdating = prevState.ScreenUpdating
    End With
End Sub

'===============================================================================
' シート保護管理
'===============================================================================
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    On Error Resume Next
    ws.Unprotect ""
    If Err.Number = 0 Then
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    protInfo.Password = InputBox("シート『" & ws.Name & "』のパスワードを入力してください。", "保護解除")
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
    If protInfo.IsProtected Then
        On Error Resume Next
        If protInfo.Password = "" Then
            ws.Protect UserInterfaceOnly:=True
        Else
            ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    End If
End Sub

'===============================================================================
' エラーハンドリング
'===============================================================================
Private Sub RaiseCustomError(ByVal errorCode As Long, ByVal description As String)
    Err.Raise errorCode, "TransferDataModule", description
End Sub

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

Private Function FriendlyErrorMessage9(ByVal errDesc As String) As String
    FriendlyErrorMessage9 = _
        "エラー #9（インデックスが有効範囲にありません）" & vbCrLf & _
        "考えられる原因と対処:" & vbCrLf & _
        "・シート名の確認：『" & DATA_SHEET_NAME & "』『" & MONTHLY_SHEET_NAME & "』が存在するか" & vbCrLf & _
        "・データ形式の確認：区分と作番が正しく入力されているか" & vbCrLf & _
        "・列構造の確認：必要な列が存在し、正しい位置にあるか" & vbCrLf & _
        vbCrLf & "詳細: " & errDesc
End Function

'===============================================================================
' 結果表示
'===============================================================================
Private Sub ShowTransferResults(ByRef result As ProcessResult)
    Dim message As String
    If result.Success Then
        message = "転記処理が完了しました。" & vbCrLf & vbCrLf & _
                  "処理件数: " & result.ProcessedCount & " 件" & vbCrLf
        If result.DuplicateCount > 0 Then
            message = message & "重複検知: " & result.DuplicateCount & " 件（黄色ハイライト表示）" & vbCrLf
        End If
        If result.NewColumnsAdded > 0 Then
            message = message & "新規列追加: " & result.NewColumnsAdded & " 列" & vbCrLf
        End If
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & "メッセージ:" & vbCrLf & result.Messages
        End If
        MsgBox message, vbInformation, "転記完了"
    Else
        message = "転記処理が中止されました。"
        If Len(result.Messages) > 0 Then
            message = message & vbCrLf & vbCrLf & result.Messages
        End If
        MsgBox message, vbExclamation, "処理中止"
    End If
End Sub