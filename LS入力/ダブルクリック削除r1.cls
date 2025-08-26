VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ==============================================
' Excel VBA - セル削除マクロ
' ==============================================
' 機能: B列をダブルクリックした行のC～F列とH列をクリア
' 対象: Excel 2016 / Windows 11
' 作成: 「JJ-07」2025/08
' ==============================================

' ==============================================
' 設定定数
' ==============================================
Private Const START_ROW As Long = 6                   ' データ開始行（この行未満では動作しない）
Private Const RIGHT_CLEAR_COLS As Long = 4            ' C列から右に削除する列数（C～F列=4列）
Private Const SELECT_CLEAR_COL As String = "H"        ' 追加で削除するH列
Private Const PROTECT_PASSWORD As String = ""         ' シート保護パスワード（使用しない場合は空文字）
Private Const CLEAR_CONSTANTS_ONLY As Boolean = False ' True: 定数のみクリア, False: 全てクリア

' ==============================================
' メインイベント: ダブルクリック処理
' ==============================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    ' ------------------------------------------
    ' 1. ダブルクリック対象の検証
    ' ------------------------------------------
    ' B列以外は処理対象外
    If Target.Column <> 2 Then Exit Sub
    
    ' 複数セル選択時は処理対象外（結合セルは許可）
    If Target.CountLarge > 1 And Not Target.MergeCells Then Exit Sub
    
    ' 対象行の決定（結合セルの場合は先頭行を使用）
    Dim targetRow As Long
    targetRow = IIf(Target.MergeCells, Target.MergeArea.Row, Target.Row)
    
    ' データ開始行未満は処理対象外
    If targetRow < START_ROW Then Exit Sub
    
    ' ------------------------------------------
    ' 2. 削除処理の実行
    ' ------------------------------------------
    ' 既定のダブルクリック編集をキャンセル
    Cancel = True
    ' 削除対象範囲の設定
    Dim clearRange As Range
    Call SetClearRange(targetRow, clearRange)
    ' シート保護の一時解除
    Call UnprotectSheet
    ' イベント無効化（再帰防止）
    Application.EnableEvents = False
    
    ' セル内容の削除実行
    Call ExecuteClearContents(clearRange)
    
    ' ------------------------------------------
    ' 3. 後処理
    ' ------------------------------------------
CleanUp:
    ' シート保護の復元
    Call ProtectSheet
    ' イベント有効化
    Application.EnableEvents = True
    
    Exit Sub

ErrorHandler:
    ' エラー発生時も必ず後処理を実行
    Resume CleanUp
End Sub

' ==============================================
' 削除範囲設定処理
' ==============================================
Private Sub SetClearRange(ByVal targetRow As Long, ByRef clearRange As Range)
    ' C～F列の範囲を設定（C列=3列目から4列分）
    Set clearRange = Me.Cells(targetRow, 3).Resize(1, RIGHT_CLEAR_COLS)
    
    ' H列を追加（H列=8列目）
    Set clearRange = Union(clearRange, Me.Cells(targetRow, 8))
End Sub

' ==============================================
' セル内容削除実行処理
' ==============================================
Private Sub ExecuteClearContents(ByVal clearRange As Range)
    If CLEAR_CONSTANTS_ONLY Then
        ' 定数セルのみクリア（数式は保持）
        On Error Resume Next ' 定数が存在しない場合のエラーを無視
        clearRange.SpecialCells(xlCellTypeConstants).ClearContents
        On Error GoTo 0      ' エラー処理を元に戻す
    Else
        ' 全セル内容をクリア（書式は保持）
        clearRange.ClearContents
    End If
End Sub

' ==============================================
' シート保護解除処理
' ==============================================
Private Sub UnprotectSheet()
    If Len(PROTECT_PASSWORD) > 0 Then
        Me.Unprotect Password:=PROTECT_PASSWORD
    End If
End Sub

' ==============================================
' シート保護復元処理
' ==============================================
Private Sub ProtectSheet()
    If Len(PROTECT_PASSWORD) > 0 Then
        ' UserInterfaceOnly:=True でVBA操作は許可、ユーザー操作は保護
        Me.Protect Password:=PROTECT_PASSWORD, UserInterfaceOnly:=True
    End If
End Sub

' ==============================================
' memo
' ==============================================
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "Sheet1"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True

