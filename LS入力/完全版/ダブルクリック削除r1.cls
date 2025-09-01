Option Explicit
'===============================================================================
' モジュール名: Sheet1 (ダブルクリック削除r1.cls)
'
' 【概要】      B列のセルがダブルクリックされた際に、該当行の指定された列の
'               データをクリアする機能を提供します。
' 【作成】    「JJ-07」2025/08
' 【対象環境】  Excel 2016 / Windows 11
' 【主要機能】
' ・B列のダブルクリックイベントをトリガーとして動作
' ・指定行のC～F列およびH列のセル内容をクリア
' ・処理対象の開始行やクリア対象列を定数で管理
' ・シート保護に対応（パスワードが設定されている場合は一時的に解除）
'===============================================================================

'===============================================================================
' 【定数宣言セクション】
' モジュール内で使用する設定値を定義
' ※動作条件を変更する場合は、このセクションを修正してください
'===============================================================================

' --- 処理対象範囲の定義 ---
Private Const START_ROW As Long = 6                   ' データクリアの対象となる開始行（これより上の行は無視）
Private Const RIGHT_CLEAR_COLS As Long = 4            ' C列から右方向にクリアする列数（C,D,E,F列 = 4列）
Private Const SELECT_CLEAR_COL As String = "H"        ' 個別に追加でクリアする列（H列）

' --- 動作オプション ---
Private Const PROTECT_PASSWORD As String = ""         ' シート保護パスワード（保護を使用しない場合は空文字のまま）
Private Const CLEAR_CONSTANTS_ONLY As Boolean = False ' True: 数式以外の値のみクリア, False: 数式も含め全てクリア

'===============================================================================
' 【メインイベントプロシージャ】
' ワークシートのダブルクリックイベントを捕捉して処理を実行
'===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' --- 変数宣言 ---
    Dim targetRow As Long       ' 処理対象となる行番号
    Dim clearRange As Range     ' クリア対象のセル範囲

    On Error GoTo ErrorHandler

    ' --- ステップ1：ダブルクリック対象の検証 ---

    ' --- 有効性チェック：B列以外がクリックされた場合は処理を終了 ---
    If Target.Column <> 2 Then Exit Sub

    ' --- 有効性チェック：複数セルが選択されている場合は処理を終了（結合セルは例外） ---
    If Target.CountLarge > 1 And Not Target.MergeCells Then Exit Sub

    ' --- 有効性チェック：データ開始行より上は処理を終了 ---
    ' ※重要：結合セルの場合も考慮し、結合範囲の先頭行を取得
    targetRow = IIf(Target.MergeCells, Target.MergeArea.Row, Target.Row)
    If targetRow < START_ROW Then Exit Sub

    ' --- ステップ2：削除処理の実行準備 ---

    ' 既定のダブルクリック動作（セル編集モード）をキャンセル
    Cancel = True

    ' クリア対象のセル範囲を特定
    Call SetClearRange(targetRow, clearRange)

    ' 必要に応じてシート保護を一時解除
    Call UnprotectSheet()

    ' ※重要：イベントの連鎖や意図しない再帰呼び出しを防ぐため、イベントを一時的に無効化
    Application.EnableEvents = False
    
    ' 対象セルの内容をクリア
    Call ExecuteClearContents(clearRange)

    GoTo CleanUp

' --- エラー発生時の処理 ---
ErrorHandler:
    MsgBox "クリア処理でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Resume CleanUp

' --- 共通後処理 ---
CleanUp:
    Application.EnableEvents = True
    Call ProtectSheet()
End Sub

'===============================================================================
' 【内部ヘルパープロシージャ】
' メイン処理から呼び出される補助的な機能
'===============================================================================

'===============================================================================
' 【機能名】削除対象範囲の設定
' 【概要】  指定された行番号に基づき、クリア対象となるセル範囲オブジェクトを生成する。
' 【引数】  targetRow (Long): [入力] 処理対象の行番号
'           clearRange (Range): [出力] 生成されたセル範囲オブジェクト
'===============================================================================
Private Sub SetClearRange(ByVal targetRow As Long, ByRef clearRange As Range)
    ' C列から定数で指定された列数分の範囲を設定
    Set clearRange = Me.Cells(targetRow, 3).Resize(1, RIGHT_CLEAR_COLS)

    ' 定数で指定された追加の列を範囲に含める
    Set clearRange = Union(clearRange, Me.Cells(targetRow, 8))
End Sub

'===============================================================================
' 【機能名】セル内容削除の実行
' 【概要】  指定されたセル範囲の内容をクリアする。
'           定数`CLEAR_CONSTANTS_ONLY`の設定に応じて動作を切り替える。
' 【引数】  clearRange (Range): クリア対象のセル範囲
'===============================================================================
Private Sub ExecuteClearContents(ByVal clearRange As Range)
    ' --- 削除方法の分岐 ---
    If CLEAR_CONSTANTS_ONLY Then
        ' --- 定数（手入力された値）のみをクリア（数式は残す） ---
        On Error Resume Next ' 対象範囲に定数がなくてもエラーにしない
        clearRange.SpecialCells(xlCellTypeConstants).ClearContents
        On Error GoTo 0      ' エラーハンドリングを元に戻す
    Else
        ' --- 全てのセル内容（値、数式）をクリア（書式は残す） ---
        clearRange.ClearContents
    End If
End Sub

'===============================================================================
' 【機能名】シート保護解除
' 【概要】  定数`PROTECT_PASSWORD`に値が設定されている場合、シートの保護を解除する。
'===============================================================================
Private Sub UnprotectSheet()
    ' --- パスワードが設定されている場合のみ解除を実行 ---
    If Len(PROTECT_PASSWORD) > 0 Then
        Me.Unprotect Password:=PROTECT_PASSWORD
    End If
End Sub

'===============================================================================
' 【機能名】シート保護復元
' 【概要】  定数`PROTECT_PASSWORD`に値が設定されている場合、シートを再保護する。
'===============================================================================
Private Sub ProtectSheet()
    ' --- パスワードが設定されている場合のみ保護を実行 ---
    If Len(PROTECT_PASSWORD) > 0 Then
        ' ※重要：UserInterfaceOnly:=True を指定することで、
        ' ユーザーのUI操作は禁止しつつ、VBAマクロからの操作は許可する状態にする。
        Me.Protect Password:=PROTECT_PASSWORD, UserInterfaceOnly:=True
    End If
End Sub

'===============================================================================
' 【VBAプロジェクト情報】
' (システムが自動生成する内部情報のため、編集不要)
'===============================================================================
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "Sheet1"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True