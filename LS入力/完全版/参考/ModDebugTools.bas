Option Explicit
'===============================================================================
' モジュール名: ModDebugTools
' デバッグ用のユーティリティ。高度な診断と簡易チェックを提供。
'===============================================================================

'-------------------------------------------------------------------------------
' 機能名: RunDeepDiagnostics
' 引数  : なし
' 役割  : ログファイルにワークブックの情報を書き出す強力なデバッグ用コード
'-------------------------------------------------------------------------------
Public Sub RunDeepDiagnostics()
    Dim f As Integer, logPath As String
    logPath = ThisWorkbook.Path & "\\debug.log"
    f = FreeFile
    On Error GoTo ErrHandler

    Open logPath For Append As #f
    Print #f, "===== Diagnostics " & Now & " ====="
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Print #f, "Sheet: " & ws.Name & ", UsedRange: " & ws.UsedRange.Address
    Next ws
    Print #f, "==============================="
    Close #f

    Debug.Print "診断ログを出力しました: " & logPath
    Exit Sub
ErrHandler:
    Debug.Print "RunDeepDiagnostics でエラーが発生しました: " & Err.Number & " " & Err.Description
    On Error Resume Next
    If f > 0 Then Close #f
End Sub

'-------------------------------------------------------------------------------
' 機能名: QuickDebugCheck
' 引数  : なし
' 役割  : 現在のブックとシート情報を即時ウィンドウに出力する簡易チェック
'-------------------------------------------------------------------------------
Public Sub QuickDebugCheck()
    If ActiveWorkbook Is Nothing Then
        Debug.Print "アクティブなブックがありません"
        Exit Sub
    End If
    Debug.Print "ブック: " & ActiveWorkbook.Name & ", シート数: " & ActiveWorkbook.Worksheets.Count
    Debug.Print "アクティブシート: " & ActiveSheet.Name
End Sub

