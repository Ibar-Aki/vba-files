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

' --- アプリケーション状態保存用 ---
' ※Excel 2016互換のため、XlCalculation列挙型はLongで保持
Private Type ApplicationState
    ScreenUpdating As Boolean    ' 画面更新の状態
    EnableEvents   As Boolean    ' イベント発火の状態
    Calculation    As Long       ' 計算モードの状態
End Type

' --- シート保護情報保存用 ---
Private Type SheetProtectionInfo
    IsProtected As Boolean       ' 元の保護状態 (True/False)
    Password    As String        ' 解除に使用したパスワード
End Type

'===============================================================================
' 【定数宣言セクション】
' モジュール内で使用する定数を定義
' ※シート名やセル番地を変更する場合は、このセクションを修正してください
'===============================================================================

' --- シート名定数 ---
Private Const ACQUISITION_SHEET_NAME As String = "データ取得" ' クリア対象シート1
Private Const DATA_SHEET_NAME        As String = "データ登録" ' クリア対象シート2

' --- クリア対象セル/範囲定数 ---
' データ取得シート
Private Const DATE_CELL_GETOUT    As String = "C4"      ' 任意日付セル
Private Const CLEAR_RANGE_ACQ     As String = "C8:F22"  ' データ範囲1
Private Const CLEAR_RANGE_ACQ2    As String = "H8:H22"  ' データ範囲2
' データ登録シート
Private Const DATE_CELL_PRIORITY  As String = "D4"      ' 任意日付セル
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
        "「" & ACQUISITION_SHEET_NAME & "」「" & DATA_SHEET_NAME & "」の入力値をクリアします。" & vbCrLf & _
        "よろしいですか。", _
        vbYesNo + vbQuestion + vbDefaultButton2, "クリアの確認") = vbNo Then
        GoTo CleanUp ' 「いいえ」が選択された場合は後処理へ
    End If

    ' --- ステップ3：ワークシートオブジェクトの取得 ---
    Set wsAcq = ThisWorkbook.Sheets(ACQUISITION_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)

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
    wsData.Range(DATE_CELL_PRIORITY).ClearContents
    wsData.Range(CLEAR_RANGE_DATA).ClearContents
    wsData.Range(DATE_CELL_WORKTIME).ClearContents

    ' --- ステップ6：完了メッセージの表示 ---
    MsgBox "入力値をクリアしました。", vbInformation, "完了"

CleanUp:
    ' --- 最終処理：シート保護とExcel状態を元に戻す ---
    RestoreSheetProtection wsAcq, protInfoAcq
    RestoreSheetProtection wsData, protInfoData
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    ' --- エラー発生時の処理 ---
    MsgBox "クリア処理でエラーが発生しました: " & Err.description, vbCritical, "エラー"
    Resume CleanUp ' エラー発生時も必ずCleanUp処理を実行
End Sub

'===============================================================================
' 【内部ヘルパー関数・プロシージャ】
' メイン処理から呼び出される補助的な機能
'===============================================================================

'===============================================================================
' 【機能名】アプリケーション状態の保存と高速化設定
' 【概要】  現在のExcel設定を構造体に保存し、処理高速化のために設定を変更する
' 【引数】  prevState: ApplicationState構造体（保存先）
'===============================================================================
Private Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    ' --- 現在の状態を構造体に保存 ---
    With prevState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
    End With
    ' --- 処理中のパフォーマンス向上のため設定を変更 ---
    With Application
        .ScreenUpdating = False              ' 画面描画を停止
        .EnableEvents = False                ' イベント発生を停止
        .Calculation = xlCalculationManual   ' 計算を手動に
    End With
End Sub

'===============================================================================
' 【機能名】アプリケーション状態の復元
' 【概要】  保存しておいたExcel設定に復元する
' 【引数】  prevState: 保存しておいた状態を持つApplicationState構造体
'===============================================================================
Private Sub RestoreApplicationState(ByRef prevState As ApplicationState)
     ' --- 保存された状態にアプリケーション設定を復元 ---
    With Application
        .Calculation = prevState.Calculation        ' 計算モードを復元
        .EnableEvents = prevState.EnableEvents      ' イベント設定を復元
        .ScreenUpdating = prevState.ScreenUpdating  ' 画面描画を最後に有効化
    End With
End Sub

'===============================================================================
' 【機能名】シート保護の解除（必要時）
' 【概要】  シートが保護されている場合、解除を試みる。
'           まず空パスワードで試行し、失敗した場合はユーザーに入力を求める。
' 【引数】  ws: 対象ワークシート
'           protInfo: 保護情報を格納する構造体
' 【戻り値】Boolean: 解除に成功した場合はTrue、失敗・キャンセルの場合はFalse
'===============================================================================
Private Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    ' --- 現在の保護状態を記録 ---
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    ' --- 保護されていない場合は、何もしないで正常終了 ---
    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    ' --- 保護されている場合の解除処理 ---
    On Error Resume Next ' パスワード間違いのエラーをハンドルするため
    ws.Unprotect ""      ' まずは空パスワードで試行
    If Err.Number = 0 Then
        ' 空パスワードで解除成功
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    ' --- 空パスワード失敗時、ユーザーに入力を促す ---
    Err.Clear
    protInfo.Password = InputBox("シート「" & ws.Name & "」の保護パスワードを入力してください。", "シート保護解除")
    If protInfo.Password = "" Then
        ' キャンセルされたか、空文字が入力された場合は失敗とする
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ' --- 入力されたパスワードで解除を試行 ---
    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0) ' エラーが発生しなければ成功
    On Error GoTo 0
End Function

'===============================================================================
' 【機能名】シート保護の復元
' 【概要】  処理開始前にシートが保護されていた場合、元の状態に再保護する
' 【引数】  ws: 対象ワークシート
'           protInfo: 保存しておいた保護情報
'===============================================================================
Private Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    ' --- 処理前に保護されていた場合のみ再保護を実行 ---
    If protInfo.IsProtected Then
        On Error Resume Next
        ' ※重要：UserInterfaceOnly:=True を指定し、ユーザー操作のみを禁止する
        ' これにより、マクロからの操作は引き続き許可される
        ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

