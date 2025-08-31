Option Explicit
'
'===============================================================================
' モジュール名: ModMonthlyMaintenance
'
' 概要:
'   月次シートの転記データを一括クリア（値と塗りつぶしを解除）し、
'   対象月のカレンダー（B列の日付）を更新します。
'   カレンダー更新は Yes/No の確認ダイアログで実行可否を確認します。
'
' 対象環境: Excel 2016+ / Windows
'===============================================================================

' --- シート名・位置など（本モジュール内で使用する定数） ---
Private Const DATA_SHEET_NAME    As String = "データ登録"   ' 対象日を取得するシート
Private Const MONTHLY_SHEET_NAME As String = "月次データ"   ' クリア・更新の対象シート

Private Const COL_DATE               As Long = 2   ' B列: 日付列
Private Const MONTHLY_MIN_COL        As Long = 3   ' C列以降が作業列
Private Const MONTHLY_HEADER_ROW     As Long = 11  ' 見出し（作業コードなど）の行
Private Const MONTHLY_DATA_START_ROW As Long = 12  ' データ開始行
 ' エラー表示セルは共通定数 ERR_CELL_ADDR を使用（ModAppConfig.bas）

'===============================================================================
' 機能名: 月次データの全クリア＋カレンダー更新
' 概要  : 月次シートの転記データ（値/時間）を全消去し、塗りつぶしも解除します。
'         その後、「データ登録」シートの対象日（D4優先→D3）と同じ月で、
'         B列にカレンダー（日付）を再作成します。実行前に確認ダイアログを表示します。
'===============================================================================
Public Sub ClearMonthlyDataAndRefreshCalendar()
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation

    Dim wsMonthly As Worksheet
    Dim wsData As Worksheet
    Dim wasProtected As Boolean
    Dim pwd As String
    Dim targetDate As Date
    Dim ret As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' --- アプリ状態の退避と負荷軽減 ---
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' --- シート取得 ---
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' --- 実行前にエラー表示セルをクリア（存在すれば） ---
    On Error Resume Next
    wsMonthly.Range(ERR_CELL_ADDR).ClearContents
    wsMonthly.Range(ERR_CELL_ADDR).WrapText = True
    On Error GoTo ErrorHandler

    ' --- 保護の一時解除 ---
    wasProtected = wsMonthly.ProtectContents
    If wasProtected Then
        On Error Resume Next
        wsMonthly.Unprotect ""
        If Err.Number <> 0 Then
            Err.Clear
            pwd = InputBox("シートが保護されています。パスワードを入力してください:", _
                           "シート保護解除")
            If Len(pwd) = 0 Then GoTo CleanUp
            wsMonthly.Unprotect Password:=pwd
        End If
        On Error GoTo ErrorHandler
    End If

    ' --- 転記データ領域の一括クリア（値/塗りつぶし） ---
    ' （対象月確定後にクリアを実施）

    ' --- 対象日（D4優先→D3）取得 ---
    If Not DetermineTargetDateLocal(wsData, targetDate) Then
    ReportErrorToMonthlySheetLocal wsMonthly, _
        "対象日付が取得できません（D4 または D3 を設定してください）", True
        GoTo CleanUp
    End If

    ' --- 転記データ領域の一括クリア（値/塗りつぶし）---
    '     対象月の末日行までを対象とし、それ以降（合計・メモ行）は削除しない
    Dim daysInMonth As Long, lastDayRow As Long
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))
    lastDayRow = MONTHLY_DATA_START_ROW + daysInMonth - 1
    ClearAllMonthlyTransferArea wsMonthly, lastDayRow

    ' --- カレンダー更新の確認 ---
    ret = MsgBox( _
        "対象月のカレンダー（日付列）を更新します。" & vbCrLf & _
        "対象月: " & Format$(targetDate, "m/dd(aaa)") & vbCrLf & vbCrLf & _
        "よろしいですか？", _
        vbYesNo + vbQuestion, "カレンダー更新の確認")
    If ret <> vbYes Then GoTo CleanUp

    ' --- カレンダー更新 ---
    RefreshMonthlyCalendar wsMonthly, targetDate

CleanUp:
    ' --- 保護の復元 ---
    If wasProtected Then
        On Error Resume Next
        If Len(pwd) > 0 Then
            wsMonthly.Protect Password:=pwd, UserInterfaceOnly:=True
        Else
            wsMonthly.Protect UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    End If

    ' --- アプリ状態の復元 ---
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrorHandler:
    ' --- 簡易エラー報告（J3に追記） ---
    On Error Resume Next
    ReportErrorToMonthlySheetLocal wsMonthly, _
        "月次クリア/カレンダー更新エラー: " & Err.Description, True
    On Error GoTo 0
    Resume CleanUp
End Sub

'-------------------------------------------------------------------------------
' 機能名: 転記データ領域の全クリア（値と塗りつぶし）
' 引数  : wsMonthly（月次シート）
'-------------------------------------------------------------------------------
Private Sub ClearAllMonthlyTransferArea(ByRef wsMonthly As Worksheet, ByVal lastDayRow As Long)
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range

    ' クリア対象の最終行は、対象月の末日行まで
    lastRow = lastDayRow

    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    If lastCol < MONTHLY_MIN_COL Then lastCol = MONTHLY_MIN_COL

    ' 値クリア＋塗りつぶし解除
    On Error Resume Next
    Set rng = wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, MONTHLY_MIN_COL), _
                              wsMonthly.Cells(lastRow, lastCol))
    If Not rng Is Nothing Then
        rng.ClearContents
        rng.Interior.Pattern = xlNone
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' 機能名: カレンダー（日付列B）の更新
' 概要  : 対象月の1日〜末日をB列に連続で設定し、余剰行があればクリア
' 引数  : wsMonthly（月次シート）、targetDate（対象日）
'-------------------------------------------------------------------------------
Private Sub RefreshMonthlyCalendar(ByRef wsMonthly As Worksheet, ByVal targetDate As Date)
    Dim firstDate As Date
    Dim daysInMonth As Long
    Dim r As Long, rowStart As Long
    Dim lastRow As Long

    rowStart = MONTHLY_DATA_START_ROW
    firstDate = DateSerial(Year(targetDate), Month(targetDate), 1)
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))

    ' 必要行数分、日付を設定
    For r = 0 To daysInMonth - 1
        With wsMonthly.Cells(rowStart + r, COL_DATE)
            .Value = firstDate + r
            .NumberFormatLocal = "mm/dd(aaa)"
            .Interior.Pattern = xlNone
        End With
    Next r

    ' 末日以降（合計・メモ行など）は保持するため、ここでは何もしない
End Sub

'-------------------------------------------------------------------------------
' 機能名: 対象日（D4優先→D3）の取得
' 引数  : wsData（データ登録シート）
' 戻り値: 取得できた場合 True
'-------------------------------------------------------------------------------
Private Function DetermineTargetDateLocal(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    DetermineTargetDateLocal = False
    If IsDate(wsData.Range("D4").Value) Then
        targetDate = CDate(wsData.Range("D4").Value)
        DetermineTargetDateLocal = True
    ElseIf IsDate(wsData.Range("D3").Value) Then
        targetDate = CDate(wsData.Range("D3").Value)
        DetermineTargetDateLocal = True
    End If
End Function

'-------------------------------------------------------------------------------
' 機能名: エラーメッセージの表示（J3）。append=True で追記
' 引数  : wsMonthly（月次シート）、message（表示内容）、append（追記フラグ）
'-------------------------------------------------------------------------------
Private Sub ReportErrorToMonthlySheetLocal(ByRef wsMonthly As Worksheet, ByVal message As String, Optional ByVal append As Boolean = False)
    On Error Resume Next
    If wsMonthly Is Nothing Then Exit Sub
    With wsMonthly.Range(ERR_CELL_ADDR)
        If append And Len(.Value) > 0 Then
            .Value = CStr(.Value) & vbLf & message
        Else
            .Value = message
        End If
        .WrapText = True
    End With
    On Error GoTo 0
End Sub
