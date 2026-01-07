'==================================================
' ThisWorkbook - ワークブックイベント処理
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' ワークブックオープン時の処理
'--------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' このワークブックを非表示にする（処理エンジンとして動作）
    If Not ThisWorkbook Is Nothing Then
        Application.Windows(ThisWorkbook.Name).Visible = False
    End If
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' ワークブッククローズ前の処理
'--------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' クリーンアップ処理があれば実行
    On Error Resume Next
    
    ' グローバルリソースの解放
    Call CleanupGlobalResources
    
    On Error GoTo 0
End Sub
