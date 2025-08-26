VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==============================================
' B列をダブルクリックした行の右3列（C:E）をクリア
' - 書式は残し、値/数式のみ削除（ClearContents）
' - ヘッダ除外・シート保護（任意）に対応
' - Excel 2016 / Windows 11 / 日本語環境想定
' ----------------------------------------------
' ★貼り付け場所：対象シートのコードモジュール（例：Sheet1）
'   VBE → プロジェクトでシートをダブルクリック → このコードを貼る
' ==============================================
Option Explicit

'=== 設定（必要に応じて変更） ===
Private Const START_ROW As Long = 6          ' データ開始行（これ未満の行では動かない）
Private Const RIGHT_CLEAR_COLS As Long = 3   ' 右に何列クリアするか（既定=3 → C:E）
Private Const PROTECT_PASSWORD As String = "" ' シート保護のパスワード（未使用なら空のまま）
Private Const CLEAR_CONSTANTS_ONLY As Boolean = False ' True: 定数セルのみクリア（数式は残す）

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo EH

    ' B列かつ単一セル（※結合セルは許容：行は結合範囲の先頭に合わせる）
    If Target.Column <> 2 Then Exit Sub
    If Target.CountLarge > 1 And Not Target.MergeCells Then Exit Sub

    Dim r As Long
    r = IIf(Target.MergeCells, Target.MergeArea.Row, Target.Row)
    If r < START_ROW Then Exit Sub

    ' ここで既定のダブルクリック編集を止める
    Cancel = True

    Dim rngClear As Range
    Set rngClear = Me.Cells(r, Target.Column + 1).Resize(1, RIGHT_CLEAR_COLS)

    ' シート保護対応（必要なら）
    If Len(PROTECT_PASSWORD) > 0 Then
        Me.Unprotect password:=PROTECT_PASSWORD
    End If

    Application.EnableEvents = False ' 再入防止

    If CLEAR_CONSTANTS_ONLY Then
        On Error Resume Next ' 定数が無いとSpecialCellsでエラー → 無視
        rngClear.SpecialCells(xlCellTypeConstants).ClearContents
        On Error GoTo EH
    Else
        rngClear.ClearContents
    End If

CleanUp:
    If Len(PROTECT_PASSWORD) > 0 Then
        ' UserInterfaceOnly:=True でユーザーの保護は維持しつつ、VBA操作は許可
        Me.Protect password:=PROTECT_PASSWORD, UserInterfaceOnly:=True
    End If
    Application.EnableEvents = True
    Exit Sub

EH:
    ' 例外時もイベントを復帰
    Resume CleanUp
End Sub
