Attribute VB_Name = "BorderExtractor"
Option Explicit

' =========================================================
' マクロ名: ExtractBorderInfo v3
' 目的: 指定範囲から罫線の詳細情報を抽出し、
'       新しいレポートシートに一覧表示します。(実務運用版)
' 変更: 2026-01-04 出力準備中で止まる問題への対策(レポートシート事前作成)
' 特徴: 高速化(配列処理)、安全性(エラーハンドリング)、
'       可読性向上(定数変換)、結合セル対応、
'       フリーズ防止(InputBox範囲指定 + チャンク処理)
' =========================================================

' チャンク処理の間隔 (この数のセルごとにDoEventsを実行)
Const CHUNK_SIZE As Long = 500
' 一定時間ごとにUIを返す(秒)
Const YIELD_INTERVAL As Double = 0.2
' 進捗ログ出力 (Immediateウィンドウ)
Const DEBUG_LOG As Boolean = True
' RGB色を取得するか (Trueで詳細色を取得。重い場合はFalse推奨)
Const READ_RGB_COLOR As Boolean = False
' 斜線罫線をチェックするか (使っていない場合はFalse推奨)
Const CHECK_DIAGONAL As Boolean = False
' サマリ作成時の繰り返し行開始行 (この行以降は連続パターンをグループ化)
Const REPEAT_ROW_START As Long = 10

Sub ExtractBorderInfo()
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim rngTarget As Range
    Dim cell As Range
    Dim i As Integer
    Dim borderTypes As Variant
    Dim borderNames As Variant
    Dim bd As Border
    Dim lineStyle As Variant
    Dim weight As Variant
    Dim colorCode As Variant
    Dim colorIndex As Variant
    Dim lastCellAddress As String
    Dim lastBorderName As String
    Dim lastStage As String
    Dim borderRange As Range
    
    ' 結果格納用配列
    Dim resultData() As Variant
    Dim resultCount As Long
    

    
    ' チャンク処理用カウンタ
    Dim cellCounter As Long
    Dim lastYield As Double
    Dim totalCells As Double
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 1. ソースの設定
    Set wsSource = ActiveSheet
    
    ' 2. InputBox で処理範囲を選択させる
    On Error Resume Next
    Set rngTarget = Application.InputBox( _
        Prompt:="罫線情報を抽出する範囲を選択してください。" & vbCrLf & _
                "キャンセルで終了します。", _
        Title:="範囲選択", _
        Default:=wsSource.UsedRange.Address, _
        Type:=8)
    On Error GoTo ErrorHandler
    
    ' キャンセルされた場合
    If rngTarget Is Nothing Then
        MsgBox "処理がキャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' データがない場合のチェック
    If rngTarget.Cells.Count = 1 And rngTarget.Cells(1, 1).Value = "" Then
        MsgBox "選択範囲にデータが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 3. 配列の初期化
    Dim estimatedMaxRows As Long
    estimatedMaxRows = rngTarget.Cells.Count * 6
    If estimatedMaxRows > 1048576 Then estimatedMaxRows = 1048576

    ' 出力シートを先に作成 (終盤のフリーズ回避)
    Set wsReport = CreateReportSheet()
    
    ReDim resultData(1 To estimatedMaxRows, 1 To 6)
    resultCount = 0
    cellCounter = 0
    lastYield = Timer
    totalCells = rngTarget.Cells.Count
    
    ' チェックする罫線の種類を定義
    If CHECK_DIAGONAL Then
        borderTypes = Array(7, 8, 9, 10, 5, 6)
        borderNames = Array("左", "上", "下", "右", "右下がり斜線", "右上がり斜線")
    Else
        borderTypes = Array(7, 8, 9, 10)
        borderNames = Array("左", "上", "下", "右")
    End If
    
    ' 4. 処理開始 (画面更新停止)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "罫線情報を抽出中... (Escキーでキャンセル)"
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = False
    
    For Each cell In rngTarget
        cellCounter = cellCounter + 1
        
        ' チャンク処理: 一定セル数 or 一定時間ごとに DoEvents を実行
        If cellCounter Mod CHUNK_SIZE = 0 Or (Timer - lastYield) > YIELD_INTERVAL Then
            Application.StatusBar = "処理中... " & cellCounter & " / " & totalCells & _
                " セル処理済み | 直近: " & lastCellAddress & " [" & lastBorderName & "] " & lastStage
            DoEvents
            lastYield = Timer
        End If
        
        ' 結合セルの場合、最初のセルのみ処理する（重複防止）
        If cell.MergeCells Then
            If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then
                GoTo ContinueLoop
            End If
        End If
        
        ' 罫線取得は結合セルならMergeArea、通常はセル自身から取得
        If cell.MergeCells Then
            Set borderRange = cell.MergeArea
        Else
            Set borderRange = cell
        End If
        
        For i = LBound(borderTypes) To UBound(borderTypes)
            Set bd = borderRange.Borders(borderTypes(i))
            
            ' 罫線が存在するか確認
            lastCellAddress = cell.Address(False, False)
            lastBorderName = borderNames(i)
            lastStage = "LineStyle"
            lineStyle = bd.LineStyle
            If lineStyle <> xlNone Then
                resultCount = resultCount + 1
                
                If resultCount > UBound(resultData, 1) Then
                    MsgBox "結果がシートの行数制限を超えそうになったため、一部のデータがカットされました。", vbExclamation
                    GoTo OutputReport
                End If
                
                lastStage = "Weight"
                weight = bd.Weight
                
                lastStage = "Color"
                On Error Resume Next
                colorIndex = bd.ColorIndex
                If READ_RGB_COLOR Then colorCode = bd.Color
                On Error GoTo ErrorHandler
                resultData(resultCount, 1) = cell.Address(False, False)
                resultData(resultCount, 2) = borderNames(i)
                resultData(resultCount, 3) = GetLineStyleName(lineStyle)
                resultData(resultCount, 4) = GetWeightName(weight)
                If READ_RGB_COLOR Then
                    resultData(resultCount, 5) = GetColorName(colorCode)
                Else
                    resultData(resultCount, 5) = GetColorIndexName(colorIndex)
                End If
                resultData(resultCount, 6) = colorIndex
            End If
        Next i
        
ContinueLoop:
    Next cell

OutputReport:
    ' 5. レポート出力
    If resultCount > 0 Then
        lastStage = "Output:Start"
        Application.StatusBar = "出力準備中... " & resultCount & " 件"
        DoEvents
        
        ' ヘッダー作成
        lastStage = "Output:Header"
        Application.StatusBar = "出力中(ヘッダー)..."
        DoEvents
        With wsReport
            .Cells(1, 1).Value = "セル位置"
            .Cells(1, 2).Value = "罫線位置"
            .Cells(1, 3).Value = "線種"
            .Cells(1, 4).Value = "太さ"
            .Cells(1, 5).Value = "色"
            .Cells(1, 6).Value = "カラーインデックス"
            .Range("A1:F1").Font.Bold = True
            .Range("A1:F1").Interior.Color = RGB(220, 230, 241)
            
            ' データ一括貼り付け
            lastStage = "Output:Write"
            Application.StatusBar = "出力中(書き込み)... " & resultCount & " 件"
            DoEvents
            .Range("A2").Resize(resultCount, 6).Value = resultData
            
            lastStage = "Output:Format"
            Application.StatusBar = "出力中(整形)..."
            DoEvents
            .Columns.AutoFit
        End With
        
        MsgBox "抽出完了！ " & resultCount & " 件の罫線情報が見つかりました。", vbInformation
    Else
        ' 結果がない場合は作成済みシートを削除
        Application.DisplayAlerts = False
        wsReport.Delete
        Application.DisplayAlerts = True
        MsgBox "対象範囲に罫線は見つかりませんでした。", vbInformation
    End If

NormalExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    If Err.Number = 18 Then
        MsgBox "処理がキャンセルされました。", vbInformation
    Else
        MsgBox "予期せぬエラーが発生しました: " & Err.Description, vbCritical
    End If
    Resume NormalExit
End Sub

' ---------------------------------------------------------
' Helper: レポートシート作成
' ---------------------------------------------------------
Function CreateReportSheet() As Worksheet
    Dim baseName As String
    Dim sheetName As String
    Dim sheetIndex As Integer
    
    If Worksheets.Count >= 255 Then
        Err.Raise vbObjectError + 1000, , "シート数が上限に近いためレポートを作成できません。"
    End If
    
    baseName = "Border_" & Format(Now, "hhmmss")
    sheetName = baseName
    sheetIndex = 1
    Do While SheetExists(sheetName)
        sheetIndex = sheetIndex + 1
        sheetName = baseName & "_" & sheetIndex
    Loop
    
    Set CreateReportSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    CreateReportSheet.Name = sheetName
End Function

' シート存在チェック
Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' ---------------------------------------------------------
' Helper Functions
' ---------------------------------------------------------

' 線種の名称取得
Function GetLineStyleName(styleCode As Variant) As String
    Select Case styleCode
        Case xlContinuous: GetLineStyleName = "実線"
        Case xlDash: GetLineStyleName = "破線"
        Case xlDashDot: GetLineStyleName = "一点鎖線"
        Case xlDashDotDot: GetLineStyleName = "二点鎖線"
        Case xlDot: GetLineStyleName = "点線"
        Case xlDouble: GetLineStyleName = "二重線"
        Case xlSlantDashDot: GetLineStyleName = "斜め一点鎖線"
        Case xlLineStyleNone: GetLineStyleName = "なし"
        Case Else: GetLineStyleName = "その他(" & styleCode & ")"
    End Select
End Function

' 太さの名称取得
Function GetWeightName(weightCode As Variant) As String
    Select Case weightCode
        Case xlHairline: GetWeightName = "極細"
        Case xlThin: GetWeightName = "細"
        Case xlMedium: GetWeightName = "中"
        Case xlThick: GetWeightName = "太"
        Case Else: GetWeightName = "その他(" & weightCode & ")"
    End Select
End Function

' 色の名称取得 (簡易版)
Function GetColorName(colorCode As Variant) As String
    Select Case colorCode
        Case 0: GetColorName = "黒(自動)"
        Case 16777215: GetColorName = "白"
        Case 255: GetColorName = "赤"
        Case 65280: GetColorName = "緑"
        Case 16711680: GetColorName = "青"
        Case 65535: GetColorName = "黄"
        Case 16711935: GetColorName = "マゼンタ"
        Case 16776960: GetColorName = "シアン"
        Case Else: GetColorName = "Color(" & colorCode & ")"
    End Select
End Function

' ColorIndexの名称取得 (簡易版)
Function GetColorIndexName(colorIndex As Variant) As String
    Select Case colorIndex
        Case xlColorIndexAutomatic: GetColorIndexName = "自動"
        Case xlColorIndexNone: GetColorIndexName = "なし/テーマ"
        Case 1: GetColorIndexName = "黒"
        Case 2: GetColorIndexName = "白"
        Case 3: GetColorIndexName = "赤"
        Case 4: GetColorIndexName = "緑"
        Case 5: GetColorIndexName = "青"
        Case 6: GetColorIndexName = "黄"
        Case 7: GetColorIndexName = "マゼンタ"
        Case 8: GetColorIndexName = "シアン"
        Case Else: GetColorIndexName = "ColorIndex(" & colorIndex & ")"
    End Select
End Function

' ---------------------------------------------------------
' 罫線データの圧縮サマリ作成
' ---------------------------------------------------------
Sub SummarizeBorderReport()
    On Error GoTo SummaryErrorHandler
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim i As Long
    Dim addr As String
    Dim rowNum As Long
    Dim colNum As Long
    Dim key As String
    Dim rowMap As Object
    Dim rowDict As Object
    Dim colList As Collection
    Dim rowKeys As Variant
    Dim r As Long
    Dim signature As String
    Dim sigToPattern As Object
    Dim patternEntries As Collection
    Dim patternId As String
    Dim outRow As Long
    Dim patRow As Long
    Dim prevSig As String
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim groupPat As String
    
    Set wsSrc = GetBorderReportSheet()
    If wsSrc Is Nothing Then
        MsgBox "罫線データのシートが見つかりません。Border_で始まるシートをアクティブにして実行してください。", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "罫線データがありません。", vbExclamation
        Exit Sub
    End If
    
    data = wsSrc.Range("A2:F" & lastRow).Value2
    
    Set rowMap = CreateObject("Scripting.Dictionary")
    rowMap.CompareMode = 0
    
    For i = 1 To UBound(data, 1)
        ' フリーズ防止
        If i Mod 1000 = 0 Then
            Application.StatusBar = "サマリ作成中... " & i & " / " & UBound(data, 1)
            DoEvents
        End If
        
        addr = CStr(data(i, 1))
        If Len(addr) = 0 Then GoTo ContinueData
        
        rowNum = GetRowFromAddress(addr)
        colNum = GetColFromAddress(addr)
        
        key = CStr(data(i, 2)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 4)) & "|" & _
              CStr(data(i, 5)) & "|" & CStr(data(i, 6))
        
        If Not rowMap.Exists(rowNum) Then
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = 0
            rowMap.Add rowNum, rowDict
        Else
            Set rowDict = rowMap(rowNum)
        End If
        
        If Not rowDict.Exists(key) Then
            Set colList = New Collection
            rowDict.Add key, colList
        Else
            Set colList = rowDict(key)
        End If
        colList.Add colNum
        
ContinueData:
    Next i
    
    rowKeys = rowMap.Keys
    If UBound(rowKeys) < 0 Then
        MsgBox "罫線データがありません。", vbExclamation
        Exit Sub
    End If
    SortLongArray rowKeys
    
    Set sigToPattern = CreateObject("Scripting.Dictionary")
    sigToPattern.CompareMode = 0
    
    Set wsOut = CreateSummarySheet("Border_Summary")
    
    wsOut.Cells(1, 1).Value = "RowRange"
    wsOut.Cells(1, 2).Value = "PatternID"
    wsOut.Cells(1, 3).Value = "RowCount"
    wsOut.Cells(1, 4).Value = "Note"
    
    outRow = 2
    prevSig = ""
    groupStart = -1
    groupEnd = -1
    groupPat = ""
    
    For r = LBound(rowKeys) To UBound(rowKeys)
        rowNum = CLng(rowKeys(r))
        Set rowDict = rowMap(rowNum)
        
        Set patternEntries = New Collection
        signature = BuildRowSignature(rowDict, patternEntries)
        
        If Not sigToPattern.Exists(signature) Then
            patternId = "P" & (sigToPattern.Count + 1)
            sigToPattern.Add signature, Array(patternId, patternEntries)
        Else
            patternId = sigToPattern(signature)(0)
        End If
        
        If rowNum >= REPEAT_ROW_START Then
            If groupStart = -1 Then
                groupStart = rowNum
                groupEnd = rowNum
                groupPat = patternId
                prevSig = signature
            ElseIf signature = prevSig And rowNum = groupEnd + 1 Then
                groupEnd = rowNum
            Else
                WriteRowGroup wsOut, outRow, groupStart, groupEnd, groupPat, "Rows>= " & REPEAT_ROW_START
                outRow = outRow + 1
                groupStart = rowNum
                groupEnd = rowNum
                groupPat = patternId
                prevSig = signature
            End If
        Else
            WriteRowGroup wsOut, outRow, rowNum, rowNum, patternId, ""
            outRow = outRow + 1
        End If
    Next r
    
    If groupStart <> -1 Then
        WriteRowGroup wsOut, outRow, groupStart, groupEnd, groupPat, "Rows>= " & REPEAT_ROW_START
        outRow = outRow + 1
    End If
    
    patRow = outRow + 2
    wsOut.Cells(patRow, 1).Value = "PatternID"
    wsOut.Cells(patRow, 2).Value = "罫線位置"
    wsOut.Cells(patRow, 3).Value = "線種"
    wsOut.Cells(patRow, 4).Value = "太さ"
    wsOut.Cells(patRow, 5).Value = "色"
    wsOut.Cells(patRow, 6).Value = "カラーインデックス"
    wsOut.Cells(patRow, 7).Value = "列範囲"
    patRow = patRow + 1
    
    Dim sigKey As Variant
    Dim entry As Variant
    For Each sigKey In sigToPattern.Keys
        patternId = sigToPattern(sigKey)(0)
        Set patternEntries = sigToPattern(sigKey)(1)
        For Each entry In patternEntries
            wsOut.Cells(patRow, 1).Value = patternId
            wsOut.Cells(patRow, 2).Value = entry(0)
            wsOut.Cells(patRow, 3).Value = entry(1)
            wsOut.Cells(patRow, 4).Value = entry(2)
            wsOut.Cells(patRow, 5).Value = entry(3)
            wsOut.Cells(patRow, 6).Value = entry(4)
            wsOut.Cells(patRow, 7).Value = entry(5)
            patRow = patRow + 1
        Next entry
    Next sigKey
    
    wsOut.Columns.AutoFit
    Application.StatusBar = False
    MsgBox "サマリ作成完了: " & wsOut.Name, vbInformation
    Exit Sub

SummaryErrorHandler:
    Application.StatusBar = False
    MsgBox "サマリ作成中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub WriteRowGroup(wsOut As Worksheet, outRow As Long, startRow As Long, endRow As Long, patternId As String, note As String)
    If startRow = endRow Then
        wsOut.Cells(outRow, 1).Value = startRow
        wsOut.Cells(outRow, 3).Value = 1
    Else
        wsOut.Cells(outRow, 1).Value = startRow & "-" & endRow
        wsOut.Cells(outRow, 3).Value = endRow - startRow + 1
    End If
    wsOut.Cells(outRow, 2).Value = patternId
    wsOut.Cells(outRow, 4).Value = note
End Sub

Private Function BuildRowSignature(rowDict As Object, patternEntries As Collection) As String
    Dim keys As Variant
    Dim k As Variant
    Dim colList As Collection
    Dim colArr As Variant
    Dim ranges As String
    Dim parts As Variant
    Dim sig As String
    
    keys = rowDict.Keys
    SortStringArray keys
    
    For Each k In keys
        Set colList = rowDict(k)
        colArr = CollectionToLongArray(colList)
        SortLongArray colArr
        ranges = ColsToColRanges(colArr)
        parts = Split(CStr(k), "|")
        sig = sig & CStr(k) & ":" & ranges & ";"
        patternEntries.Add Array(parts(0), parts(1), parts(2), parts(3), parts(4), ranges)
    Next k
    
    BuildRowSignature = sig
End Function

Private Function ColsToColRanges(colArr As Variant) As String
    Dim i As Long
    Dim startCol As Long
    Dim prevCol As Long
    Dim part As String
    Dim result As String
    
    If IsEmpty(colArr) Then
        ColsToColRanges = ""
        Exit Function
    End If
    
    startCol = CLng(colArr(LBound(colArr)))
    prevCol = startCol
    
    For i = LBound(colArr) + 1 To UBound(colArr)
        If CLng(colArr(i)) = prevCol + 1 Then
            prevCol = CLng(colArr(i))
        Else
            part = ColRangeToText(startCol, prevCol)
            If Len(result) = 0 Then
                result = part
            Else
                result = result & ", " & part
            End If
            startCol = CLng(colArr(i))
            prevCol = startCol
        End If
    Next i
    
    part = ColRangeToText(startCol, prevCol)
    If Len(result) = 0 Then
        result = part
    Else
        result = result & ", " & part
    End If
    
    ColsToColRanges = result
End Function

Private Function ColRangeToText(startCol As Long, endCol As Long) As String
    Dim startText As String
    Dim endText As String
    startText = ColNumberToLetters(startCol)
    endText = ColNumberToLetters(endCol)
    If startCol = endCol Then
        ColRangeToText = startText
    Else
        ColRangeToText = startText & ":" & endText
    End If
End Function

Private Function CollectionToLongArray(colList As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To colList.Count - 1)
    For i = 1 To colList.Count
        arr(i - 1) = CLng(colList(i))
    Next i
    CollectionToLongArray = arr
End Function

Private Sub SortLongArray(ByRef arr As Variant)
    If IsEmpty(arr) Then Exit Sub
    QuickSortLong arr, LBound(arr), UBound(arr)
End Sub

Private Sub QuickSortLong(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim mid As Variant
    Dim temp As Variant
    low = first
    high = last
    mid = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < mid
            low = low + 1
        Loop
        Do While arr(high) > mid
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSortLong arr, first, high
    If low < last Then QuickSortLong arr, low, last
End Sub

Private Sub SortStringArray(ByRef arr As Variant)
    If IsEmpty(arr) Then Exit Sub
    QuickSortString arr, LBound(arr), UBound(arr)
End Sub

Private Sub QuickSortString(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim mid As Variant
    Dim temp As Variant
    low = first
    high = last
    mid = arr((first + last) \ 2)
    Do While low <= high
        Do While CStr(arr(low)) < CStr(mid)
            low = low + 1
        Loop
        Do While CStr(arr(high)) > CStr(mid)
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSortString arr, first, high
    If low < last Then QuickSortString arr, low, last
End Sub

Private Function GetRowFromAddress(addr As String) As Long
    Dim s As String
    Dim i As Long
    s = Replace(addr, "$", "")
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            GetRowFromAddress = CLng(Mid$(s, i))
            Exit Function
        End If
    Next i
    GetRowFromAddress = 0
End Function

Private Function GetColFromAddress(addr As String) As Long
    Dim s As String
    Dim i As Long
    Dim letters As String
    s = Replace(addr, "$", "")
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            letters = Left$(s, i - 1)
            Exit For
        End If
    Next i
    GetColFromAddress = ColLettersToNumber(letters)
End Function

Private Function ColLettersToNumber(colLetters As String) As Long
    Dim i As Long
    Dim result As Long
    Dim ch As Integer
    For i = 1 To Len(colLetters)
        ch = Asc(UCase$(Mid$(colLetters, i, 1))) - Asc("A") + 1
        result = result * 26 + ch
    Next i
    ColLettersToNumber = result
End Function

Private Function ColNumberToLetters(colNum As Long) As String
    Dim result As String
    Dim n As Long
    n = colNum
    Do While n > 0
        result = Chr$(((n - 1) Mod 26) + Asc("A")) & result
        n = (n - 1) \ 26
    Loop
    ColNumberToLetters = result
End Function

Private Function GetBorderReportSheet() As Worksheet
    Dim ws As Worksheet
    If IsBorderReportSheet(ActiveSheet) Then
        Set GetBorderReportSheet = ActiveSheet
        Exit Function
    End If
    For Each ws In Worksheets
        If IsBorderReportSheet(ws) Then
            Set GetBorderReportSheet = ws
            Exit Function
        End If
    Next ws
    Set GetBorderReportSheet = Nothing
End Function

Private Function IsBorderReportSheet(ws As Worksheet) As Boolean
    On Error Resume Next
    IsBorderReportSheet = (CStr(ws.Cells(1, 1).Value) = "セル位置" And CStr(ws.Cells(1, 2).Value) = "罫線位置")
    On Error GoTo 0
End Function

Private Function CreateSummarySheet(baseName As String) As Worksheet
    Dim name As String
    Dim idx As Long
    name = baseName
    idx = 1
    Do While SheetExists(name)
        idx = idx + 1
        name = baseName & "_" & idx
    Loop
    Set CreateSummarySheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    CreateSummarySheet.Name = name
End Function
