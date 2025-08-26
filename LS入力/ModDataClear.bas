Attribute VB_Name = "ModDataClear"
'========================================
' Module : ModDataClear
' 機能   : 入力欄クリア（データ取得・データ登録）
' 対象   : Excel 2016+ / Windows
'========================================
Option Explicit

'=== アプリ状態の保存用（Excel 2016互換：XlCalculationはLongで保持） ===
Private Type ApplicationState
    ScreenUpdating As Boolean
    EnableEvents   As Boolean
    Calculation    As Long
End Type

'=== シート保護の復元用 ===
Private Type SheetProtectionInfo
    IsProtected As Boolean
    Password    As String
End Type

'=== 対象シート/範囲の定数（必要に応じて変更してください） ===
Private Const ACQUISITION_SHEET_NAME As String = "データ取得"
Private Const DATA_SHEET_NAME        As String = "データ登録"

' データ登録シート：日付セル
Private Const DATE_CELL_GETOUT    As String = "C4"      'データ取得の任意日付の削除
Private Const DATE_CELL_PRIORITY  As String = "D4"      'データ登録の任意日付の削除
Private Const DATE_CELL_WORKTIME As String = "E19"      'データ登録の勤務時間の削除
' クリア範囲
Private Const CLEAR_RANGE_ACQ   As String = "C8:F22"  ' データ取得
Private Const CLEAR_RANGE_ACQ2  As String = "H8:H22"  ' データ取得
Private Const CLEAR_RANGE_DATA  As String = "F8:F17"  ' データ登録

'========================================
' 公開手続き：入力データの一括クリア
'========================================
Public Sub ClearInputData()
    Dim wsAcq As Worksheet, wsData As Worksheet
    Dim protInfoAcq As SheetProtectionInfo
    Dim protInfoData As SheetProtectionInfo
    Dim prevState As ApplicationState

    SaveAndSetApplicationState prevState
    On Error GoTo ErrorHandler

    If MsgBox( _
        "「" & ACQUISITION_SHEET_NAME & "」「" & DATA_SHEET_NAME & "」の入力値をクリアします。" & vbCrLf & _
        "よろしいですか。", _
        vbYesNo + vbQuestion + vbDefaultButton2, "クリアの確認") = vbNo Then
        GoTo CleanUp
    End If

    Set wsAcq = ThisWorkbook.Sheets(ACQUISITION_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' シート保護の一時解除（必要時）
    If Not UnprotectSheetIfNeeded(wsAcq, protInfoAcq) Then GoTo CleanUp
    If Not UnprotectSheetIfNeeded(wsData, protInfoData) Then GoTo CleanUp

    '--- クリア処理 ---
    ' データ取得
    wsAcq.Range(CLEAR_RANGE_ACQ).ClearContents
    wsAcq.Range(CLEAR_RANGE_ACQ2).ClearContents
    wsAcq.Range(DATE_CELL_GETOUT).ClearContents

    ' データ登録
    wsData.Range(DATE_CELL_PRIORITY).ClearContents
    wsData.Range(CLEAR_RANGE_DATA).ClearContents
    wsData.Range(DATE_CELL_WORKTIME).ClearContents

    MsgBox "入力値をクリアしました。", vbInformation, "完了"

CleanUp:
    ' シート保護の復元 / アプリ状態の復元
    RestoreSheetProtection wsAcq, protInfoAcq
    RestoreSheetProtection wsData, protInfoData
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    MsgBox "クリア処理でエラーが発生しました: " & Err.description, vbCritical, "エラー"
    Resume CleanUp
End Sub

'========================================
' 内部ヘルパー
'========================================
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

' 保護解除（必要時）。パスワード不明でも空文字→入力依頼の順で試行
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    On Error Resume Next
    ws.Unprotect ""                 ' まずは空パスで試行
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
    ' 処理前に保護されていた場合のみ再保護する
    If protInfo.IsProtected Then
        On Error Resume Next
        ' UserInterfaceOnly:=True を指定し、マクロからの操作は許可する
        ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub


