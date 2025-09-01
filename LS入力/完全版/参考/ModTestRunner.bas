Option Explicit
'===============================================================================
' モジュール名: ModTestRunner
'
' 【概要】主要マクロを一括実行するテスト支援モジュール。
'          `LS入力/tests/test_patterns.md` に列挙したテストを手動で
'          実行する際の補助機能を提供する。
'===============================================================================

' エントリーポイント ---------------------------------------------------------
Public Sub RunAllTests()
    Test_GetOutlookSchedule
    Test_ClearInputData
    Test_TransferDataToMonthlySheet
    Test_ClearMonthlyDataAndRefreshCalendar
    Test_DoubleClickClear
    MsgBox "全テスト完了", vbInformation
End Sub

' 個別テスト ---------------------------------------------------------------
Public Sub Test_GetOutlookSchedule()
    On Error GoTo ErrHandler
    Debug.Print "[Test] GetOutlookSchedule"
    Call GetOutlookSchedule()
    Debug.Print "    -> 完了"
    Exit Sub
ErrHandler:
    Debug.Print "    -> エラー: " & Err.Description
End Sub

Public Sub Test_ClearInputData()
    On Error GoTo ErrHandler
    Debug.Print "[Test] ClearInputData"
    Call ClearInputData()
    Debug.Print "    -> 完了"
    Exit Sub
ErrHandler:
    Debug.Print "    -> エラー: " & Err.Description
End Sub

Public Sub Test_TransferDataToMonthlySheet()
    On Error GoTo ErrHandler
    Debug.Print "[Test] TransferDataToMonthlySheet"
    Call TransferDataToMonthlySheet()
    Debug.Print "    -> 完了"
    Exit Sub
ErrHandler:
    Debug.Print "    -> エラー: " & Err.Description
End Sub

Public Sub Test_ClearMonthlyDataAndRefreshCalendar()
    On Error GoTo ErrHandler
    Debug.Print "[Test] ClearMonthlyDataAndRefreshCalendar"
    Call ClearMonthlyDataAndRefreshCalendar()
    Debug.Print "    -> 完了"
    Exit Sub
ErrHandler:
    Debug.Print "    -> エラー: " & Err.Description
End Sub

Public Sub Test_DoubleClickClear()
    On Error GoTo ErrHandler
    Dim cancel As Boolean
    Debug.Print "[Test] Worksheet_BeforeDoubleClick"
    Call Worksheet_BeforeDoubleClick(GetSheet(Sheet_DataAcquire).Range("B10"), cancel)
    Debug.Print "    -> 完了"
    Exit Sub
ErrHandler:
    Debug.Print "    -> エラー: " & Err.Description
End Sub

