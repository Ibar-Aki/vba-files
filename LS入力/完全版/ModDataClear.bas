Option Explicit
'===============================================================================
' モジュール名: ModDataClear

' 【概要】「データ取得」「データ登録」両シートの入力欄をクリアする機能を提供します。
' 【作成】「JJ-07」2025/08
' 【対象環境】Excel 2016+ / Windows
' 【主要機能】
' ・ユーザーに確認ダイアログを表示後、対象範囲のデータを一括でクリア
' ・処理中の画面更新やイベントを停止し、高速かつ安全に実行
' ・シート保護されている場合は、パスワードを要求して一時的に解除し、処理後に再保護
'===============================================================================


'===============================================================================
' 【データ構造（Type宣言セクション）】
' 処理に必要な情報をまとめた構造体
'===============================================================================


'===============================================================================
' 【定数宣言セクション】
' モジュール内で使用する定数を定義
' ※シート名やセル番地を変更する場合は、このセクションを修正してください
'===============================================================================

 ' シート名は ModAppConfig.bas の SheetName Enum を使用

' --- クリア対象セル/範囲定数 ---
' データ取得シート
Private Const DATE_CELL_GETOUT    As String = "C4"      ' 任意日付セル
Private Const CLEAR_RANGE_ACQ     As String = "C8:F22"  ' データ範囲1
Private Const CLEAR_RANGE_ACQ2    As String = "H8:H22"  ' データ範囲2
' データ登録シート
Private Const DATE_CELL_WORKTIME  As String = "E24"     ' 勤務時間セル
Private Const CLEAR_RANGE_DATA    As String = "F8:F22"  ' データ範囲

'===============================================================================
' 【メイン処理】
' ユーザーが実行する入力データの一括クリア処理
'===============================================================================
Public Sub ClearInputData()
    ' --- 変数宣言 ---
    Dim wsAcq As Worksheet              ' 「データ取得」シートオブジェクト
    Dim wsData As Worksheet             ' 「データ登録」シートオブジェクト
    Dim protInfoAcq As SheetProtectionInfo  ' 「データ取得」シートの保護情報
    Dim protInfoData As SheetProtectionInfo ' 「データ登録」シートの保護情報
    Dim prevState As ApplicationState   ' Excelの実行前状態

    ' --- ステップ1：Excel状態の保存と高速化設定 ---
    SaveAndSetApplicationState prevState
    On Error GoTo ErrorHandler

    ' --- ステップ2：ユーザーへの最終確認 ---
    If MsgBox( _
        "「" & GetSheetName(Sheet_DataAcquire) & "」「" & GetSheetName(Sheet_DataEntry) & "」の入力値をクリアします。" & vbCrLf & _
        "よろしいですか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, "クリアの確認") = vbNo Then
        GoTo CleanUp ' 「いいえ」が選択された場合は後処理へ
    End If

    ' --- ステップ3：ワークシートオブジェクトの取得 ---
    Set wsAcq = GetSheet(Sheet_DataAcquire)
    Set wsData = GetSheet(Sheet_DataEntry)

    ' --- ステップ4：シート保護の一時解除 ---
    ' ※保護されている場合、パスワード入力が求められます
    If Not UnprotectSheetIfNeeded(wsAcq, protInfoAcq) Then GoTo CleanUp
    If Not UnprotectSheetIfNeeded(wsData, protInfoData) Then GoTo CleanUp

    ' --- ステップ5：指定範囲のデータクリア実行 ---
    ' データ取得シートの入力値をクリア
    wsAcq.Range(CLEAR_RANGE_ACQ).ClearContents
    wsAcq.Range(CLEAR_RANGE_ACQ2).ClearContents
    wsAcq.Range(DATE_CELL_GETOUT).ClearContents

    ' データ登録シートの入力値をクリア
    wsData.Range(DATA_ENTRY_DATE_CELL).ClearContents
    wsData.Range(CLEAR_RANGE_DATA).ClearContents
    wsData.Range(DATE_CELL_WORKTIME).ClearContents

    ' --- ステップ6：完了メッセージの表示 ---
    MsgBox "クリア完了", vbInformation, "完了"

CleanUp:
    ' --- 最終処理：シート保護とExcel状態を元に戻す ---
    If Not wsAcq Is Nothing Then
        RestoreSheetProtection wsAcq, protInfoAcq
    End If
    If Not wsData Is Nothing Then
        RestoreSheetProtection wsData, protInfoData
    End If
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    ' --- エラー発生時の処理 ---
    MsgBox "クリア処理でエラーが発生しました: " & Err.description, vbCritical, "エラー"
    Resume CleanUp ' エラー発生時も必ずCleanUp処理を実行
End Sub
