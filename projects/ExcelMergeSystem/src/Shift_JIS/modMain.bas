'==================================================
' modMain - メイン処理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' モジュールレベル変数（グローバル変数からの変更）
'--------------------------------------------------
Private m_ConfigData As Object
Private m_LogCollection As Collection
Private m_IsProcessing As Boolean

'--------------------------------------------------
' 設定データへのアクセサ（カプセル化）
'--------------------------------------------------
Public Property Get ConfigData() As Object
    Set ConfigData = m_ConfigData
End Property

Public Property Set ConfigData(ByVal value As Object)
    Set m_ConfigData = value
End Property

Public Property Get LogCollection() As Collection
    Set LogCollection = m_LogCollection
End Property

Public Property Set LogCollection(ByVal value As Collection)
    Set m_LogCollection = value
End Property

'--------------------------------------------------
' メインエントリーポイント（外部から呼び出される）
'--------------------------------------------------
Public Sub ExecuteMerge(ByVal strFile1 As String, ByVal strFile2 As String)
    
    Dim startTime As Date
    Dim result As Boolean
    Dim appState As Object
    
    ' 二重実行防止
    If m_IsProcessing Then
        MsgBox "処理が既に実行中です。", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    m_IsProcessing = True
    result = False
    
    ' 初期化
    startTime = Now
    Set appState = SaveApplicationState()
    
    ' アプリケーション設定（パフォーマンス向上）
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' ログ初期化
    Call InitializeLog
    Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Call LogMessage(APP_TITLE & " v" & APP_VERSION & " 処理開始", LOG_LEVEL_INFO)
    Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Call LogMessage("Excel1: " & GetFileName(strFile1), LOG_LEVEL_INFO)
    Call LogMessage("Excel2: " & GetFileName(strFile2), LOG_LEVEL_INFO)
    
    ' 設定読込
    If Not LoadConfiguration() Then
        Call LogMessage("設定ファイルの読み込みに失敗しました", LOG_LEVEL_ERROR)
        GoTo Cleanup
    End If
    
    ' ファイル検証
    If Not ValidateFiles(strFile1, strFile2) Then
        Call LogMessage("ファイル検証エラー", LOG_LEVEL_ERROR)
        GoTo Cleanup
    End If
    
    ' データ処理実行
    result = ProcessMerge(strFile1, strFile2)
    
    If result Then
        Call LogMessage("処理完了 処理時間: " & _
            Format(Now - startTime, TIMESTAMP_FORMAT_TIME), LOG_LEVEL_INFO)
        Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Else
        Call LogMessage("処理失敗", LOG_LEVEL_ERROR)
    End If
    
Cleanup:
    ' アプリケーション状態復元
    Call RestoreApplicationState(appState)
    
    ' メモリ解放
    Call CleanupResources
    
    ' 処理完了メッセージ
    If result Then
        MsgBox "処理が完了しました。" & vbCrLf & _
               "出力フォルダを確認してください。", _
               vbInformation, APP_TITLE
    Else
        MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
               "ログを確認してください。", _
               vbExclamation, APP_TITLE
    End If
    
    m_IsProcessing = False
    
    ' 安全な自己終了（他のブックに影響を与えない）
    Call SafeCloseThisWorkbook
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage("システムエラー: " & Err.Description & _
                   " (エラー番号: " & Err.Number & ")", LOG_LEVEL_ERROR)
    result = False
    Resume Cleanup
    
End Sub

'--------------------------------------------------
' 結合処理メイン
'--------------------------------------------------
Private Function ProcessMerge(ByVal file1 As String, ByVal file2 As String) As Boolean
    
    Dim data1 As Object
    Dim data2 As Object
    Dim mergedData As Object
    Dim outputPath As String
    
    On Error GoTo ErrorHandler
    
    ProcessMerge = False
    Set data1 = Nothing
    Set data2 = Nothing
    Set mergedData = Nothing
    
    ' Excel1読込
    Call LogMessage("Excel1読込開始...", LOG_LEVEL_INFO)
    Set data1 = LoadExcelData(file1, "Excel1")
    If data1 Is Nothing Then
        Call LogMessage("Excel1のデータ読込に失敗しました", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' Excel2読込
    Call LogMessage("Excel2読込開始...", LOG_LEVEL_INFO)
    Set data2 = LoadExcelData(file2, "Excel2")
    If data2 Is Nothing Then
        Call LogMessage("Excel2のデータ読込に失敗しました", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' データ結合
    Call LogMessage("データ結合処理開始...", LOG_LEVEL_INFO)
    Set mergedData = MergeData(data1, data2)
    If mergedData Is Nothing Then
        Call LogMessage("データ結合に失敗しました", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' 結果を追加情報として保存
    mergedData("File1Name") = GetFileName(file1)
    mergedData("File2Name") = GetFileName(file2)
    
    ' 出力
    Call LogMessage("ファイル出力開始...", LOG_LEVEL_INFO)
    outputPath = GenerateOutput(mergedData)
    
    ProcessMerge = (outputPath <> "")
    
    If ProcessMerge Then
        Call LogMessage("出力ファイル: " & outputPath, LOG_LEVEL_INFO)
    End If
    
CleanupLocal:
    ' ローカルオブジェクトのメモリ解放
    Set data1 = Nothing
    Set data2 = Nothing
    Set mergedData = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ProcessMerge Error: " & Err.Description, LOG_LEVEL_ERROR)
    ProcessMerge = False
    Resume CleanupLocal
    
End Function

'--------------------------------------------------
' アプリケーション状態の保存
'--------------------------------------------------
Private Function SaveApplicationState() As Object
    Dim state As Object
    Set state = CreateObject("Scripting.Dictionary")
    
    With Application
        state("ScreenUpdating") = .ScreenUpdating
        state("DisplayAlerts") = .DisplayAlerts
        state("Calculation") = .Calculation
        state("EnableEvents") = .EnableEvents
    End With
    
    Set SaveApplicationState = state
End Function

'--------------------------------------------------
' アプリケーション状態の復元
'--------------------------------------------------
Private Sub RestoreApplicationState(ByVal state As Object)
    If state Is Nothing Then
        ' デフォルト状態に復元
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End With
    Else
        With Application
            .ScreenUpdating = state("ScreenUpdating")
            .DisplayAlerts = state("DisplayAlerts")
            .Calculation = state("Calculation")
            .EnableEvents = state("EnableEvents")
        End With
    End If
End Sub

'--------------------------------------------------
' リソースクリーンアップ
'--------------------------------------------------
Private Sub CleanupResources()
    On Error Resume Next
    
    ' モジュールレベル変数のクリア
    Set m_ConfigData = Nothing
    Set m_LogCollection = Nothing
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' グローバルリソースのクリーンアップ（ThisWorkbookから呼び出し）
'--------------------------------------------------
Public Sub CleanupGlobalResources()
    Call CleanupResources
End Sub

'--------------------------------------------------
' 安全な自己終了
'--------------------------------------------------
Private Sub SafeCloseThisWorkbook()
    On Error Resume Next
    
    Dim wb As Workbook
    
    ' 他のブックが開いているか確認
    If Workbooks.Count > 1 Then
        ' 他のブックがある場合は自分だけ閉じる
        ThisWorkbook.Close SaveChanges:=False
    Else
        ' 自分だけの場合は何もしない（Excelが終了してしまうため）
        ' ユーザーが手動で閉じる
    End If
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' ファイル名取得
'--------------------------------------------------
Public Function GetFileName(ByVal fullPath As String) As String
    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function
