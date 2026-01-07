Option Explicit
'
'===============================================================================
' モジュール名: ModCommonUtils
' 共通のアプリ状態管理およびシート保護関連のユーティリティ
'===============================================================================

Public Type ApplicationState
    ScreenUpdating As Boolean    ' 画面更新の状態
    EnableEvents   As Boolean    ' イベント発生の状態
    Calculation    As Long       ' 計算モード
End Type

Public Type SheetProtectionInfo
    IsProtected As Boolean       ' 元の保護状態
    Password    As String        ' 解除に使用したパスワード
End Type

Public Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
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

Public Sub RestoreApplicationState(ByRef prevState As ApplicationState)
    With Application
        .Calculation = prevState.Calculation
        .EnableEvents = prevState.EnableEvents
        .ScreenUpdating = prevState.ScreenUpdating
    End With
End Sub

Public Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
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
    protInfo.Password = InputBox("シート【" & ws.Name & "】の保護パスワードを入力してください。", "保護解除")
    If protInfo.Password = "" Then
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    If ws Is Nothing Then Exit Sub
    If protInfo.IsProtected Then
        On Error Resume Next
        ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

'-------------------------------------------------------------------------------
' 機能名: 対象日（D4優先→D3）の取得
' 引数  : wsData（データ登録シート）、targetDate（取得した日付）
' 戻り値: 取得できた場合 True
'-------------------------------------------------------------------------------
Public Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    Const DATE_CELL_NORMAL As String = "D3"
    DetermineTargetDate = False
    If IsDate(wsData.Range(DATA_ENTRY_DATE_CELL).Value) Then
        targetDate = CDate(wsData.Range(DATA_ENTRY_DATE_CELL).Value)
        DetermineTargetDate = True
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).Value)
        DetermineTargetDate = True
    End If
End Function

'-------------------------------------------------------------------------------
' 機能名: エラーメッセージの表示（J3）。append=True で追記
' 引数  : message（表示内容）、append（追記フラグ）
'-------------------------------------------------------------------------------
Public Sub ReportErrorToMonthlySheet(ByVal message As String, Optional ByVal append As Boolean = False)
    Dim wsMonthly As Worksheet
    On Error Resume Next
    Set wsMonthly = GetSheet(Sheet_Monthly)
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

'-------------------------------------------------------------------------------
' 機能名: 月次シートのエラーセルをクリア
' 引数  : なし
'-------------------------------------------------------------------------------
Public Sub ClearErrorCellOnMonthlySheet()
    Dim wsMonthly As Worksheet
    On Error Resume Next
    Set wsMonthly = GetSheet(Sheet_Monthly)
    If wsMonthly Is Nothing Then Exit Sub
    wsMonthly.Range(ERR_CELL_ADDR).ClearContents
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' 機能名: エラー詳細メッセージの取得
' 引数  : errNo（発生したエラー番号）, errDesc（エラー内容）
' 戻り値: 整形したエラーメッセージ
'-------------------------------------------------------------------------------
Public Function GetErrorDetails(ByVal errNo As Long, ByVal errDesc As String) As String
    Dim displayNo As Long
    Dim msg As String

    If errNo >= vbObjectError And errNo < 0 Then
        displayNo = errNo - vbObjectError
    Else
        displayNo = errNo
    End If

    msg = "エラー番号: " & CStr(displayNo)
    If Len(errDesc) > 0 Then
        msg = msg & vbCrLf & "内容: " & errDesc
    End If
    GetErrorDetails = msg
End Function

