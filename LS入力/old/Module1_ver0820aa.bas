'================================================================================
' Outlook予定取得マクロ (Version 1.4 - 出力列の調整)
' 修正内容:
' ・件名に応じた自動分類機能を追加 (v1.3)
' ・会議時間、分類、区分の出力列を調整 (v1.4)
'================================================================================
Sub GetOutlookSchedule()
    '============================================================
    ' ■ 1. 初期設定と定数宣言
    '============================================================
    ' --- ユーザー設定項目 (環境に合わせて変更してください) ---
    Const TARGET_SHEET_NAME As String = "データ取得"   ' マクロを実行するシート名
    Const DATE_INPUT_CELL As String = "C3"         ' 日付が入力されているセル番地
    Const OUTPUT_HEADER_ROW As Long = 7            ' 出力先のヘッダー行番号
    Const OUTPUT_START_COLUMN As String = "C"      ' 出力先の開始列

    ' --- 追加機能に関する設定項目 ---
    Const DEST_SHEET_NAME As String = "データ登録"   ' 転記先のシート名
    Const SOURCE_CELL As String = "C4"             ' 転記元のセル番地 (データ取得シート)
    Const DEST_CELL As String = "D4"               ' 転記先のセル番地 (データ登録シート)

    ' --- 変数宣言 ---
    Dim ws As Worksheet, wsDest As Worksheet
    Dim wasProtected As Boolean
    Dim olApp As Object, olNs As Object, olFolder As Object, olItems As Object
    Dim olRestrictedItems As Object, olApt As Object
    Dim targetDate As Date, cellValue As Variant
    Dim outputRow As Long, lastOutputRow As Long, actualCount As Long
    
    ' --- 自動分類機能で利用する変数 ---
    Dim keyMatrixRange As Range, classListRange As Range
    Dim keyMatrixKubunRange As Range, classListKubunRange As Range

    ' ■ エラー発生時はErrorHandlerセクションへジャンプ
    On Error GoTo ErrorHandler
    ' ■ 処理中の画面描画を停止してちらつき防止と高速化
    Application.ScreenUpdating = False

    '============================================================
    ' ■ 2. 実行前チェックと準備
    '============================================================
    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    ' --- シート保護状態を記録し、必要に応じて解除 ---
    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect
        If Err.Number <> 0 Then
            On Error GoTo ErrorHandler
            Dim userPassword As String
            userPassword = InputBox("シートがパスワードで保護されています。パスワードを入力してください:", "パスワード入力")
            If userPassword = "" Then GoTo CleanUp
            ws.Unprotect Password:=userPassword
        End If
        On Error GoTo ErrorHandler
    End If

    ' --- 日付入力セルのチェック ---
    cellValue = ws.Range(DATE_INPUT_CELL).Value
    If IsEmpty(cellValue) Or cellValue = "" Then
        MsgBox "セル " & DATE_INPUT_CELL & " が空欄です。日付を入力してください。", vbExclamation, "入力エラー"
        GoTo CleanUp
    End If
    targetDate = CDate(cellValue)
    
    ' --- 自動分類用の名前付き範囲をチェック ---
    On Error Resume Next
    Set keyMatrixRange = ThisWorkbook.Names("KeyMatrix").RefersToRange
    Set classListRange = ThisWorkbook.Names("ClassList").RefersToRange
    Set keyMatrixKubunRange = ThisWorkbook.Names("KeyMatrix_区分").RefersToRange
    Set classListKubunRange = ThisWorkbook.Names("ClassList_区分").RefersToRange
    On Error GoTo ErrorHandler
    If keyMatrixRange Is Nothing Or classListRange Is Nothing Or keyMatrixKubunRange Is Nothing Or classListKubunRange Is Nothing Then
        MsgBox "自動分類に必要な名前付き範囲（KeyMatrix, ClassListなど）が見つかりません。" & vbCrLf & "処理を中止します。", vbCritical
        GoTo CleanUp
    End If

    '============================================================
    ' ■ 3. Excelシートの出力範囲をクリア
    '============================================================
    outputRow = OUTPUT_HEADER_ROW + 1
    lastOutputRow = ws.Cells(ws.Rows.Count, OUTPUT_START_COLUMN).End(xlUp).Row

    If lastOutputRow >= outputRow Then
        Dim startColNum As Long
        startColNum = ws.Range(OUTPUT_START_COLUMN & "1").Column
        ' ★変更: クリア範囲をH列（開始列+5）まで拡張
        ws.Range(ws.Cells(outputRow, startColNum), ws.Cells(lastOutputRow, startColNum + 5)).ClearContents
    End If
    
    ' ★変更: ヘッダーをH列まで設定 (G列は空白)
    With ws.Cells(OUTPUT_HEADER_ROW, OUTPUT_START_COLUMN).Resize(1, 6)
        .Value = Array("時間", "件名", "会議時間", "分類", "", "区分")
        .Font.Bold = True
    End With

    '============================================================
    ' ■ 4. Outlookアプリケーションへの接続
    '============================================================
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo ErrorHandler
    If olApp Is Nothing Then Err.Raise vbObjectError, , "Outlookに接続できません"

    '============================================================
    ' ■ 5. 指定日の予定をOutlookから取得
    '============================================================
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(9) ' 9はolFolderCalendar
    Set olItems = olFolder.items
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True

    Dim filterString As String
    filterString = "[Start] <= '" & Format(targetDate, "yyyy/MM/dd 23:59") & "' AND [End] >= '" & Format(targetDate, "yyyy/MM/dd 00:00") & "'"
    Set olRestrictedItems = olItems.Restrict(filterString)

    '============================================================
    ' ■ 6. 取得した予定をExcelシートへ出力
    '============================================================
    If olRestrictedItems.Count = 0 Then
        ws.Cells(outputRow, OUTPUT_START_COLUMN).Value = "予定はありません"
    Else
        actualCount = 0
        For Each olApt In olRestrictedItems
            actualCount = actualCount + 1
            Dim subject As String
            subject = olApt.Subject
            
            ' --- 基本情報の出力 ---
            ws.Cells(outputRow, "C").Value = Format(olApt.Start, "hhmm") & "-" & Format(olApt.End, "hhmm")
            ws.Cells(outputRow, "D").Value = subject
            
            ' --- ★変更: 会議時間の計算と出力を E列 に ---
            Dim totalMinutes As Long, hours As Long, minutes As Long
            totalMinutes = DateDiff("n", olApt.Start, olApt.End)
            hours = totalMinutes \ 60
            minutes = totalMinutes Mod 60
            With ws.Cells(outputRow, "E")
                .NumberFormat = "@"
                .Value = Format(hours, "00") & Format(minutes, "00")
            End With

            ' --- ★変更: 自動分類処理の出力先を F列 と H列 に ---
            ws.Cells(outputRow, "F").Value = GetClassification(subject, keyMatrixRange, classListRange)
            ws.Cells(outputRow, "H").Value = GetClassification(subject, keyMatrixKubunRange, classListKubunRange)
            
            outputRow = outputRow + 1
        Next olApt
    End If
    MsgBox Format(targetDate, "yyyy年mm月dd日") & " の予定を " & actualCount & " 件取得しました。", vbInformation, "処理完了"

    '============================================================
    ' ■ 7. データ転記処理
    '============================================================
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET_NAME)
    On Error GoTo ErrorHandler

    If Not wsDest Is Nothing Then
        If Not IsEmpty(ws.Range(SOURCE_CELL).Value) And ws.Range(SOURCE_CELL).Value <> "" Then
            wsDest.Range(DEST_CELL).Value = ws.Range(SOURCE_CELL).Value
        End If
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "エラー"
    GoTo CleanUp

CleanUp:
    If Not ws Is Nothing And wasProtected Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    Set olApt = Nothing: Set olRestrictedItems = Nothing: Set olItems = Nothing
    Set olFolder = Nothing: Set olNs = Nothing: Set olApp = Nothing
    Set ws = Nothing: Set wsDest = Nothing
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------
' ■ Helper Function: 件名とキーワードリストを照合し、対応する分類名を返す
' (この部分のコードに変更はありません)
'--------------------------------------------------------------------------------
Private Function GetClassification(subjectText As String, keyRange As Range, classRange As Range) As String
    Dim keyArray As Variant, classArray As Variant
    Dim i As Long, j As Long
    
    If keyRange.Count = 1 Then
        ReDim keyArray(1 To 1, 1 To 1): keyArray(1, 1) = keyRange.Value
    Else
        keyArray = keyRange.Value
    End If
    
    If classRange.Count = 1 Then
        ReDim classArray(1 To 1, 1 To 1): classArray(1, 1) = classRange.Value
    Else
        classArray = classRange.Value
    End If

    For i = LBound(keyArray, 1) To UBound(keyArray, 1)
        For j = LBound(keyArray, 2) To UBound(keyArray, 2)
            Dim keyword As String: keyword = CStr(keyArray(i, j))
            If keyword <> "" Then
                If InStr(1, UCase(subjectText), UCase(keyword)) > 0 Then
                    GetClassification = CStr(classArray(i, 1))
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    GetClassification = ""
End Function

' 実行用のマクロ
Sub ExecuteOutlookSchedule()
    Call GetOutlookSchedule
End Sub