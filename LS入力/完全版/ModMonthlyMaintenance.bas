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
' シート名・列番号は ModAppConfig.bas の Enum を使用

 ' 行番号は ModAppConfig.bas の MonthlySheetRow Enum を使用
 ' エラー表示セルは共通定数 ERR_CELL_ADDR を使用（ModAppConfig.bas）

'===============================================================================
' 機能名: 月次データの全クリア＋カレンダー更新
' 概要  : 月次シートの転記データ（値/時間）を全消去し、塗りつぶしも解除します。
'         その後、「データ登録」シートの対象日（D4優先→D3）と同じ月で、
'         B列にカレンダー（日付）を再作成します。実行前に確認ダイアログを表示します。
' 引数  : showConfirm - True の場合は処理前に確認ダイアログを表示します。
'         False の場合は確認を省略します。
'===============================================================================
Public Sub ClearMonthlyDataAndRefreshCalendar(Optional ByVal showConfirm As Boolean = True)
    Dim appState As ApplicationState

    Dim wsMonthly As Worksheet
    Dim wsData As Worksheet
    Dim wasProtected As Boolean
    Dim pwd As String
    Dim targetDate As Date
    Dim ret As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' --- アプリ状態の退避と負荷軽減 ---
    SaveAndSetApplicationState appState

    ' --- シート取得 ---
    Set wsMonthly = GetSheet(Sheet_Monthly)
    Set wsData = GetSheet(Sheet_DataEntry)

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
    If Not DetermineTargetDate(wsData, targetDate) Then
        ReportErrorToMonthlySheet "対象日付が取得できません（" & DATA_ENTRY_DATE_CELL & " または D3 を設定してください）", True
        GoTo CleanUp
    End If

    ' --- 転記データ領域の一括クリア（値/塗りつぶし）---
    '     対象月の末日行までを対象とし、それ以降（合計・メモ行）は削除しない
    Dim daysInMonth As Long, lastDayRow As Long
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))
    lastDayRow = MonthlyRow_DataStart + daysInMonth - 1
    ClearAllMonthlyTransferArea wsMonthly, lastDayRow

    ' --- カレンダー更新の確認 ---
    ret = vbYes
    If showConfirm Then
        ret = MsgBox( _
            "対象月のカレンダー（日付列）を更新します。" & vbCrLf & _
            "対象月: " & Format$(targetDate, "m/dd(aaa)") & vbCrLf & vbCrLf & _
            "よろしいですか？", _
            vbYesNo + vbQuestion, "カレンダー更新の確認")
    End If
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
    RestoreApplicationState appState
    Exit Sub

ErrorHandler:
    ' --- 簡易エラー報告（J3に追記） ---
    On Error Resume Next
    ReportErrorToMonthlySheet "月次クリア/カレンダー更新エラー: " & Err.Description, True
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

    lastCol = wsMonthly.Cells(MonthlyRow_Header, wsMonthly.Columns.Count).End(xlToLeft).Column
    If lastCol < MonthlyCol_Min Then lastCol = MonthlyCol_Min

    ' 値クリア＋塗りつぶし解除
    On Error Resume Next
    Set rng = wsMonthly.Range(wsMonthly.Cells(MonthlyRow_DataStart, MonthlyCol_Min), _
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

    rowStart = MonthlyRow_DataStart
    firstDate = DateSerial(Year(targetDate), Month(targetDate), 1)
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))

    ' 必要行数分、日付を設定
    For r = 0 To daysInMonth - 1
        With wsMonthly.Cells(rowStart + r, MonthlyCol_Date)
            .Value = firstDate + r
            .NumberFormatLocal = "mm/dd(aaa)"
            .Interior.Pattern = xlNone
        End With
    Next r

    ' 末日以降（合計・メモ行など）は保持するため、ここでは何もしない
End Sub

