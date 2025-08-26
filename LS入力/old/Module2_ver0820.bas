
Attribute VB_Name = "Module2"
'================================================================================
' 機能: 「データ登録」シートから「月次データ」シートへデータを転記・集計する
'       区分＋作番の組み合わせで転記先を特定する方式に改善
' 作成者: Gemini (改良版)
' バージョン: 2.0 (区分＋作番対応、シート保護対応)
'================================================================================

Option Explicit

'--- Windows APIの宣言 (64bit/32bit両対応) ---
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If

' --- 定数宣言 ---
Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_UNICODETEXT As Long = 13

Private Const DATA_SHEET_NAME As String = "データ登録"
Private Const ACQUISITION_SHEET_NAME As String = "データ取得"
Private Const MONTHLY_SHEET_NAME As String = "月次データ"

Private Const DATE_CELL_PRIORITY As String = "D4"
Private Const DATE_CELL_NORMAL As String = "D3"

Private Const DATA_START_ROW As Long = 8
Private Const MONTHLY_WORKNO_ROW As Long = 8      ' 作番行
Private Const MONTHLY_HEADER_ROW As Long = 9      ' 区分行（ヘッダー）
Private Const MONTHLY_DATA_START_ROW As Long = 10
Private Const MONTHLY_DATA_END_ROW As Long = 40

'================================================================================
' ■ メイン処理: ボタンに登録するプロシージャ (データ転記)
'================================================================================
Sub TransferDataToMonthlySheet()

    ' --- 変数宣言 ---
    Dim wsData As Worksheet, wsMonthly As Worksheet
    Dim wasMonthlyProtected As Boolean
    Dim targetDate As Date
    Dim lastDataRow As Long, i As Long, transferCount As Long
    Dim categoryWorkNoDic As Object  ' Key: "区分|作番", Value: 列番号
    Dim targetRow As Long
    Dim previewText As String
    Dim addDecision As Object  ' Key: "区分|作番", Value: True=列追加許可 / False=追加拒否
    Dim workNumber As String, category As String, timeValue As Variant
    Dim targetCol As Long, combinedKey As String
    Set addDecision = CreateObject("Scripting.Dictionary")

    ' --- 事前準備 ---
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ' ワークシートオブジェクトを設定
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    On Error GoTo ErrorHandler
    If wsData Is Nothing Or wsMonthly Is Nothing Then
        MsgBox "「" & DATA_SHEET_NAME & "」または「" & MONTHLY_SHEET_NAME & "」シートが見つかりません。", vbCritical
        GoTo CleanUp
    End If
    
    ' --- シート保護の解除 ---
    wasMonthlyProtected = wsMonthly.ProtectContents
    If wasMonthlyProtected Then
        If Not UnprotectSheet(wsMonthly) Then GoTo CleanUp
    End If

    ' --- 1. 登録日の取得と検証 ---
    If IsDate(wsData.Range(DATE_CELL_PRIORITY).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_PRIORITY).Value)
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).Value)
    Else
        MsgBox "登録日（D3セル）または任意日付（D4セル）が有効な日付ではありません。", vbExclamation
        GoTo CleanUp
    End If
    
    ' --- 2. 転記元データの最終行を取得 ---
    lastDataRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    If lastDataRow < DATA_START_ROW Then
        MsgBox "転記するデータがありません。", vbInformation
        GoTo CleanUp
    End If
    
    ' --- 3. 転記先ヘッダー（区分＋作番）の読み込み ---
    Set categoryWorkNoDic = CreateObject("Scripting.Dictionary")
    Dim lastHeaderCol As Long, c As Long
    lastHeaderCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    
    For c = 3 To lastHeaderCol ' C列から開始
        Dim categoryName As String, workNoName As String
        categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).Value))
        workNoName = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).Value))
        
        If categoryName <> "" Then
            combinedKey = categoryName & "|" & workNoName
            If Not categoryWorkNoDic.Exists(combinedKey) Then
                categoryWorkNoDic.Add combinedKey, c
            End If
        End If
    Next c
    
    ' --- 4. 転記先の日付行を検索 ---
    targetRow = FindMatchingDateRow(wsMonthly, targetDate)
    If targetRow = 0 Then
        MsgBox "転記先シート「" & MONTHLY_SHEET_NAME & "」に、登録日「" & Format(targetDate, "m/d") & "」が見つかりませんでした。" & vbCrLf & _
               "B10:B" & MONTHLY_DATA_END_ROW & "の範囲に日付が存在するか確認してください。", vbExclamation
        GoTo CleanUp
    End If
    
    ' --- 5. 転記内容のプレビューを作成 ---
    previewText = "以下の内容でデータを転記します。よろしいですか？" & vbCrLf & vbCrLf
    previewText = previewText & "登録日: " & Format(targetDate, "yyyy/mm/dd") & " (" & wsMonthly.Cells(targetRow, "B").text & ")" & vbCrLf & vbCrLf
    previewText = previewText & "---------------------------------------------------------------" & vbCrLf
    previewText = previewText & "作番" & vbTab & " | " & "区分" & vbTab & " | " & "時間" & vbCrLf
    previewText = previewText & "---------------------------------------------------------------" & vbCrLf
    
    Dim previewDic As Object
    Set previewDic = CreateObject("Scripting.Dictionary")
    
    For i = DATA_START_ROW To lastDataRow
        If Not (IsEmpty(wsData.Cells(i, "C").Value) And IsEmpty(wsData.Cells(i, "D").Value)) Then
            Dim workNo As String, cat As String, timeVal As Variant, mins As Double, previewKey As String
            workNo = Trim$(CStr(wsData.Cells(i, "C").Value))
            cat = Trim$(CStr(wsData.Cells(i, "D").Value))
            timeVal = wsData.Cells(i, "E").Value
            mins = ConvertToMinutes(timeVal)
            
            If mins > 0 And workNo <> "" And cat <> "" Then
                previewKey = workNo & "|" & cat
                If previewDic.Exists(previewKey) Then
                    previewDic(previewKey) = previewDic(previewKey) + mins
                Else
                    previewDic.Add previewKey, mins
                End If
            End If
        End If
    Next i
    
    If previewDic.Count = 0 Then
        MsgBox "転記する有効な時間データがありません。", vbInformation
        GoTo CleanUp
    End If

    Dim key As Variant
    For Each key In previewDic.Keys
        Dim keyParts As Variant
        keyParts = Split(key, "|")
        previewText = previewText & keyParts(0) & vbTab & " | " & keyParts(1) & vbTab & " | " & MinutesToHHMMString(previewDic(key)) & vbCrLf
    Next key
    
    If MsgBox(previewText, vbYesNo + vbQuestion, "転記内容のプレビュー") = vbNo Then
        MsgBox "処理を中断しました。", vbInformation
        GoTo CleanUp
    End If

    ' --- 6. データをクリップボードへコピー ---
    Dim clipboardText As String
    clipboardText = ""
    For i = DATA_START_ROW To lastDataRow
        If Not (IsEmpty(wsData.Cells(i, "C").Value) And IsEmpty(wsData.Cells(i, "D").Value) And IsEmpty(wsData.Cells(i, "E").Value)) Then
            clipboardText = clipboardText & wsData.Cells(i, "C").text & vbTab & _
                                          wsData.Cells(i, "D").text & vbTab & vbTab & _
                                          wsData.Cells(i, "E").text & vbCrLf
        End If
    Next i

    If clipboardText <> "" Then
        CopyTextToClipboard clipboardText
    End If

    ' --- 7. データ転記処理（区分＋作番での照合） ---
    transferCount = 0

    
    For i = DATA_START_ROW To lastDataRow
        If Not (IsEmpty(wsData.Cells(i, "C").Value) And IsEmpty(wsData.Cells(i, "D").Value)) Then

            ' 必須チェック
            If IsEmpty(wsData.Cells(i, "C").Value) Or IsEmpty(wsData.Cells(i, "D").Value) Then GoTo Skip_Row
            timeValue = wsData.Cells(i, "E").Value
            If IsEmpty(timeValue) Or Not IsNumeric(timeValue) Then GoTo Skip_Row

            workNumber = Trim$(CStr(wsData.Cells(i, "C").Value))
            category = Trim$(CStr(wsData.Cells(i, "D").Value))
            combinedKey = category & "|" & workNumber

            ' 組み合わせキーでの照合・新規列追加処理
            If Not categoryWorkNoDic.Exists(combinedKey) Then
                If Not addDecision.Exists(combinedKey) Then
                    targetCol = EnsureCategoryWorkNoColumn( _
                                  category, workNumber, wsMonthly, categoryWorkNoDic)
                    addDecision.Add combinedKey, (targetCol > 0)
                End If

                ' 追加を選ばなかった組み合わせはスキップ
                If addDecision(combinedKey) = False Then GoTo Skip_Row
            End If

            targetCol = categoryWorkNoDic(combinedKey)

            ' 転記（シリアル値をそのまま）
            wsMonthly.Cells(targetRow, targetCol).Value = timeValue

            transferCount = transferCount + 1
Skip_Row:
        End If
    Next i
    
    ' --- 8. 事後処理 ---
    If transferCount = 0 Then
        MsgBox "転記対象となるデータがありませんでした。", vbInformation
    Else
        MsgBox transferCount & "件のデータを転記し、クリップボードにコピーしました。", vbInformation, "処理完了"
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
           
CleanUp:
    ' --- シートの再保護 ---
    If Not wsMonthly Is Nothing Then
        If wasMonthlyProtected Then wsMonthly.Protect
    End If
    
    ' --- オブジェクトの解放 ---
    Set wsData = Nothing
    Set wsMonthly = Nothing
    Set categoryWorkNoDic = Nothing
    Set previewDic = Nothing
    Set addDecision = Nothing
    Application.ScreenUpdating = True

End Sub

'================================================================================
' ■ データ消去用マクロ
'================================================================================
Sub ClearInputData()
    Dim wsAcquisition As Worksheet, wsData As Worksheet
    Dim wasAcquisitionProtected As Boolean, wasDataProtected As Boolean
    
    On Error GoTo ClearErrorHandler
    
    ' 確認メッセージ
    If MsgBox("「" & ACQUISITION_SHEET_NAME & "」シートと「" & DATA_SHEET_NAME & "」シートの入力データをクリアします。" & vbCrLf & "よろしいですか？", vbYesNo + vbQuestion, "クリアの確認") = vbNo Then
        MsgBox "処理を中断しました。", vbInformation
        Exit Sub
    End If
    
    ' シートの取得
    Set wsAcquisition = ThisWorkbook.Sheets(ACQUISITION_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    
    ' --- シート保護の解除 ---
    wasAcquisitionProtected = wsAcquisition.ProtectContents
    If wasAcquisitionProtected Then
        If Not UnprotectSheet(wsAcquisition) Then GoTo ClearCleanup
    End If
    
    wasDataProtected = wsData.ProtectContents
    If wasDataProtected Then
        If Not UnprotectSheet(wsData) Then GoTo ClearCleanup
    End If
    
    ' データのクリア
    wsAcquisition.Range("C8:E22").ClearContents
    wsData.Range("D4").ClearContents
    wsData.Range("F8:F17").ClearContents
    
    MsgBox "データのクリアが完了しました。", vbInformation
    
    GoTo ClearCleanup

ClearErrorHandler:
    MsgBox "クリア処理中にエラーが発生しました。" & vbCrLf & _
           "シート名が正しいか確認してください。" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
           
ClearCleanup:
    ' --- シートの再保護 ---
    If Not wsAcquisition Is Nothing Then
        If wasAcquisitionProtected Then wsAcquisition.Protect
    End If
    If Not wsData Is Nothing Then
        If wasDataProtected Then wsData.Protect
    End If
    
    ' --- オブジェクトの解放 ---
    Set wsAcquisition = Nothing
    Set wsData = Nothing
End Sub

'================================================================================
' ■ ヘルパー関数群
'================================================================================

' --- 機能: Excelの時刻形式(小数)とHHMM形式(整数/文字列)の両方を分に変換する
Private Function ConvertToMinutes(ByVal timeValue As Variant) As Double
    ConvertToMinutes = 0
    If IsEmpty(timeValue) Or Not IsNumeric(timeValue) Then Exit Function

    If InStr(1, CStr(timeValue), ".") > 0 Then
        ConvertToMinutes = timeValue * 1440
    Else
        Dim hhmmStr As String, hours As Long, minutesPart As Long
        hhmmStr = CStr(CLng(timeValue))
        Select Case Len(hhmmStr)
            Case 1, 2: hours = 0: minutesPart = CLng(hhmmStr)
            Case 3, 4: hours = CLng(Left(hhmmStr, Len(hhmmStr) - 2)): minutesPart = CLng(Right(hhmmStr, 2))
            Case Else: Exit Function
        End Select
        If minutesPart < 60 Then ConvertToMinutes = (hours * 60) + minutesPart
    End If
End Function

' --- 機能: 合計分をHHMM形式の文字列に変換する (例: 90 -> "0130")
Private Function MinutesToHHMMString(ByVal totalMinutes As Double) As String
    Dim hours As Long
    Dim minutesPart As Long
    If totalMinutes <= 0 Then
        MinutesToHHMMString = "0000"
    Else
        hours = Int(totalMinutes / 60)
        minutesPart = Round(totalMinutes - (hours * 60), 0)
        If minutesPart = 60 Then
            hours = hours + 1
            minutesPart = 0
        End If
        MinutesToHHMMString = Format(hours, "00") & Format(minutesPart, "00")
    End If
End Function

' --- 機能: 指定された日付をB10:B40の範囲から検索し、行番号を返す
Private Function FindMatchingDateRow(ws As Worksheet, targetDate As Date) As Long
    Dim i As Long, cellDate As Date
    For i = MONTHLY_DATA_START_ROW To MONTHLY_DATA_END_ROW
        If IsDate(ws.Cells(i, "B").Value) Then
            cellDate = CDate(ws.Cells(i, "B").Value)
            If Int(cellDate) = Int(targetDate) Then
                FindMatchingDateRow = i
                Exit Function
            End If
        End If
    Next i
    FindMatchingDateRow = 0
End Function

' --- 機能: 文字列をクリップボードにコピーする (64bit/32bit対応)
Private Sub CopyTextToClipboard(ByVal text As String)
    #If VBA7 Then
        Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    #Else
        Dim hGlobalMemory As Long, lpGlobalMemory As Long
    #End If
    Dim lngSize As Long

    If text = "" Then Exit Sub

    lngSize = (Len(text) + 1) * 2
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, lngSize)
    If hGlobalMemory = 0 Then Exit Sub

    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        lstrcpyW lpGlobalMemory, StrPtr(text)
        GlobalUnlock hGlobalMemory

        If OpenClipboard(0&) <> 0 Then
            EmptyClipboard
            SetClipboardData CF_UNICODETEXT, hGlobalMemory
            CloseClipboard
        End If
    End If
End Sub

' --- 機能: シートの保護を解除する (パスワードを要求)
Private Function UnprotectSheet(ws As Worksheet) As Boolean
    UnprotectSheet = False ' 初期値は失敗
    On Error Resume Next
    
    ' パスワードなしで試行
    ws.Unprotect ""
    If Err.Number = 0 Then
        UnprotectSheet = True
        Exit Function
    End If
    
    ' パスワードが必要な場合
    Err.Clear
    Dim password As String
    password = InputBox("シート「" & ws.Name & "」は保護されています。" & vbCrLf & "パスワードを入力してください。", "保護の解除")
    
    ' キャンセルされた場合
    If password = "" Then
        MsgBox "処理を中断しました。", vbInformation
        Exit Function
    End If
    
    ' 入力されたパスワードで試行
    ws.Unprotect password
    If Err.Number <> 0 Then
        MsgBox "パスワードが違います。処理を中断しました。", vbCritical
        Exit Function
    End If
    
    On Error GoTo 0
    UnprotectSheet = True
End Function

'================================================================================
' ■ 区分＋作番の組み合わせ列を確保する関数（新規）
'    指定区分＋作番の列が無ければ、Yes/Noでヘッダー末尾に追加して列番号を返す
'    戻り値: 追加/既存なら列番号、No選択時は0
'================================================================================
Private Function EnsureCategoryWorkNoColumn(ByVal category As String, _
                                           ByVal workNo As String, _
                                           ByRef wsMonthly As Worksheet, _
                                           ByRef categoryWorkNoDic As Object) As Long
    Dim resp As VbMsgBoxResult
    Dim lastCol As Long, newCol As Long, prevCol As Long
    Dim combinedKey As String

    combinedKey = category & "|" & workNo

    ' 既に存在するならその列番号を返す
    If categoryWorkNoDic.Exists(combinedKey) Then
        EnsureCategoryWorkNoColumn = categoryWorkNoDic(combinedKey)
        Exit Function
    End If

    resp = MsgBox("区分「" & category & "」＋作番「" & workNo & "」の転記先が見つかりません。" & vbCrLf & _
                  "ヘッダー（" & MONTHLY_HEADER_ROW & "行目）に列を追加しますか？", _
                  vbYesNo + vbQuestion, "転記先（区分＋作番列）の追加")

    If resp = vbNo Then
        EnsureCategoryWorkNoColumn = 0
        Exit Function
    End If

    ' 追加先の列番号を決定（ヘッダー最終列の右隣）
    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    prevCol = IIf(newCol > 1, newCol - 1, newCol)

    ' 値の設定（ヘッダーと作番）
    wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).Value = category  ' 9行目：区分
    wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).Value = workNo    ' 8行目：作番

    ' 体裁（直前列の見た目を踏襲）
    On Error Resume Next
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(prevCol).ColumnWidth
    
    ' ヘッダー行（9行目）の体裁
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).WrapText
    End With
    
    ' 作番行（8行目）の体裁
    With wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).WrapText
    End With
    On Error GoTo 0

    ' 辞書へ登録して列番号を返す
    categoryWorkNoDic.Add combinedKey, newCol
    EnsureCategoryWorkNoColumn = newCol
End Function

