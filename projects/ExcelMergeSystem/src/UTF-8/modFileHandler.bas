'==================================================
' modFileHandler - ファイル処理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' Excelデータ読込
'--------------------------------------------------
Public Function LoadExcelData(ByVal filePath As String, _
                            ByVal configSection As String) As Object
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dataDict As Object
    Dim headerRows As Long
    Dim dataStartRow As Long
    Dim idColumn As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim idColNum As Long
    
    On Error GoTo ErrorHandler
    
    Set LoadExcelData = Nothing
    Set wb = Nothing
    Set dataDict = Nothing
    
    ' 設定値取得
    headerRows = GetConfigValueLong(configSection & "_HeaderRows", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_HEADER_ROWS, DEFAULT_EXCEL2_HEADER_ROWS))
    dataStartRow = GetConfigValueLong(configSection & "_DataStartRow", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_DATA_START_ROW, DEFAULT_EXCEL2_DATA_START_ROW))
    idColumn = GetConfigValue(configSection & "_IDColumn", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_ID_COLUMN, DEFAULT_EXCEL2_ID_COLUMN))
    
    ' 列番号に変換
    idColNum = ColumnLetterToNumber(idColumn)
    If idColNum = 0 Then
        Call LogMessage("無効な列指定: " & idColumn, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' ファイルオープン
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    
    If wb.Worksheets.Count = 0 Then
        Call LogMessage("ワークシートが存在しません: " & filePath, LOG_LEVEL_ERROR)
        wb.Close False
        Exit Function
    End If
    
    Set ws = wb.Worksheets(1)
    
    ' データ範囲特定
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, idColNum).End(xlUp).Row, dataStartRow)
    lastCol = Application.Max(ws.Cells(headerRows, ws.Columns.Count).End(xlToLeft).Column, 1)
    
    ' 最大行数チェック
    If (lastRow - dataStartRow + 1) > MAX_PROCESSING_ROWS Then
        Call LogMessage("データ行数が制限を超えています (" & _
                       (lastRow - dataStartRow + 1) & "行 > " & MAX_PROCESSING_ROWS & "行)", LOG_LEVEL_WARNING)
    End If
    
    ' Dictionary作成
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    ' ヘッダー情報取得
    dataDict("Headers") = GetHeaders(ws, headerRows, lastCol)
    dataDict("HeaderRows") = headerRows
    dataDict("LastCol") = lastCol
    dataDict("IDColumn") = idColumn
    dataDict("IDColumnNum") = idColNum  ' 列番号も保存
    
    ' データ取得
    Set dataDict("Data") = GetDataWithID(ws, dataStartRow, lastRow, idColNum, lastCol)
    
    ' ファイル情報
    dataDict("FileName") = GetFileName(filePath)
    dataDict("RowCount") = Application.Max(0, lastRow - dataStartRow + 1)
    
    ' クローズ
    wb.Close False
    Set wb = Nothing
    
    Set LoadExcelData = dataDict
    
    Call LogMessage(configSection & " 読込完了: " & _
                   dataDict("RowCount") & "件", LOG_LEVEL_INFO)
    
    Exit Function
    
ErrorHandler:
    If Not wb Is Nothing Then
        wb.Close False
        Set wb = Nothing
    End If
    Set LoadExcelData = Nothing
    Call LogMessage("LoadExcelData Error: " & Err.Description & _
                   " (File: " & filePath & ")", LOG_LEVEL_ERROR)
    
End Function

'--------------------------------------------------
' ヘッダー取得（セル結合対応）
'--------------------------------------------------
Private Function GetHeaders(ByVal ws As Worksheet, _
                          ByVal headerRows As Long, _
                          ByVal lastCol As Long) As Variant
    
    Dim headers() As String
    Dim i As Long
    Dim j As Long
    Dim cell As Range
    Dim mergeArea As Range
    
    On Error GoTo ErrorHandler
    
    ' 配列を事前に初期化
    ReDim headers(1 To headerRows, 1 To lastCol)
    
    For i = 1 To headerRows
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If cell.MergeCells Then
                ' セル結合の場合は左上の値を使用
                Set mergeArea = cell.MergeArea
                headers(i, j) = CStr(mergeArea.Cells(1, 1).Value)
            Else
                headers(i, j) = CStr(cell.Value)
            End If
        Next j
    Next i
    
    GetHeaders = headers
    Exit Function
    
ErrorHandler:
    ' エラー時も初期化済み配列を返す
    Call LogMessage("GetHeaders Error: " & Err.Description, LOG_LEVEL_ERROR)
    GetHeaders = headers
    
End Function

'--------------------------------------------------
' データ取得（識別コード付き）
'--------------------------------------------------
Private Function GetDataWithID(ByVal ws As Worksheet, _
                             ByVal startRow As Long, _
                             ByVal endRow As Long, _
                             ByVal idCol As Long, _
                             ByVal lastCol As Long) As Object
    
    Dim dataDict As Object
    Dim duplicates As Object
    Dim rowDict As Object
    Dim i As Long
    Dim idValue As String
    Dim rowData As Variant
    Dim validCount As Long
    Dim key As Variant
    
    On Error GoTo ErrorHandler
    
    Set dataDict = CreateObject("Scripting.Dictionary")
    Set duplicates = CreateObject("Scripting.Dictionary")
    
    validCount = 0
    
    For i = startRow To endRow
        idValue = Trim(CStr(ws.Cells(i, idCol).Value))
        
        ' 空白IDはスキップ
        If idValue <> "" Then
            ' 重複チェック
            If dataDict.Exists(idValue) Then
                If Not duplicates.Exists(idValue) Then
                    duplicates(idValue) = CStr(dataDict(idValue)("Row")) & "," & CStr(i)
                Else
                    duplicates(idValue) = duplicates(idValue) & "," & CStr(i)
                End If
            Else
                ' データ格納
                rowData = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Value
                
                Set rowDict = CreateObject("Scripting.Dictionary")
                rowDict("Data") = rowData
                rowDict("Row") = i
                Set dataDict(idValue) = rowDict
                
                validCount = validCount + 1
            End If
        End If
        
        ' 進捗表示（1000件ごと）
        If i Mod 1000 = 0 Then
            DoEvents
        End If
    Next i
    
    ' 重複があれば警告
    If duplicates.Count > 0 Then
        For Each key In duplicates.Keys
            Call LogMessage("識別コード重複: " & key & _
                          " (行: " & duplicates(key) & ")", LOG_LEVEL_WARNING)
        Next key
    End If
    
    Call LogMessage("有効データ件数: " & validCount & "件", LOG_LEVEL_INFO)
    
    Set GetDataWithID = dataDict
    Exit Function
    
ErrorHandler:
    Call LogMessage("GetDataWithID Error: " & Err.Description, LOG_LEVEL_ERROR)
    If dataDict Is Nothing Then
        Set dataDict = CreateObject("Scripting.Dictionary")
    End If
    Set GetDataWithID = dataDict
    
End Function

'--------------------------------------------------
' 出力ファイル生成
'--------------------------------------------------
Public Function GenerateOutput(ByVal mergedData As Object) As String
    
    Dim wbOut As Workbook
    Dim wsData As Worksheet
    Dim outputPath As String
    Dim row As Long
    Dim col As Long
    Dim headers1 As Variant
    Dim headers2 As Variant
    Dim id As Variant
    Dim rowData As Variant
    Dim i As Long
    Dim lastHeaderRow As Long
    Dim includeLogSheet As Boolean
    Dim excel2IdColNum As Long
    
    On Error GoTo ErrorHandler
    
    GenerateOutput = ""
    Set wbOut = Nothing
    
    ' 新規ワークブック作成
    Set wbOut = Workbooks.Add
    Set wsData = wbOut.Worksheets(1)
    wsData.Name = "結合データ"
    
    ' ヘッダー作成
    headers1 = mergedData("Headers1")
    headers2 = mergedData("Headers2")
    
    ' Excel2の識別コード列番号を取得
    excel2IdColNum = mergedData("Excel2IDColumnNum")
    If excel2IdColNum = 0 Then excel2IdColNum = 1  ' デフォルトは1列目
    
    ' 最終ヘッダー行を特定
    lastHeaderRow = Application.Max(UBound(headers1, 1), UBound(headers2, 1))
    
    ' ヘッダー出力（セル結合解除して最終行のみ出力）
    row = 1
    col = 1
    
    ' Excel1ヘッダー
    For i = 1 To UBound(headers1, 2)
        wsData.Cells(row, col).Value = headers1(UBound(headers1, 1), i)
        col = col + 1
    Next i
    
    ' Excel2ヘッダー（識別コード列を除く）
    For i = 1 To UBound(headers2, 2)
        ' 識別コード列はスキップ（設定から動的に取得）
        If i <> excel2IdColNum Then
            wsData.Cells(row, col).Value = headers2(UBound(headers2, 1), i)
            col = col + 1
        End If
    Next i
    
    ' ヘッダー書式設定
    With wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, col - 1))
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BG
        .Borders.LineStyle = xlContinuous
    End With
    
    ' データ出力
    row = 2
    For Each id In mergedData("MergedRows").Keys
        rowData = mergedData("MergedRows")(id)
        
        For col = 1 To UBound(rowData, 2)
            wsData.Cells(row, col).Value = rowData(1, col)
        Next col
        
        row = row + 1
        
        ' 進捗表示（100件ごと）
        If row Mod 100 = 0 Then
            DoEvents
        End If
    Next id
    
    ' 列幅自動調整
    wsData.Cells.EntireColumn.AutoFit
    
    ' ログシート作成
    includeLogSheet = GetConfigValueBool(CFG_INCLUDE_LOG_SHEET, DEFAULT_INCLUDE_LOG_SHEET)
    If includeLogSheet Then
        Call CreateLogSheet(wbOut, mergedData("Statistics"))
    End If
    
    ' ファイル保存
    outputPath = GetOutputPath()
    wbOut.SaveAs outputPath, xlOpenXMLWorkbook
    wbOut.Close False
    Set wbOut = Nothing
    
    GenerateOutput = outputPath
    Call LogMessage("出力完了: " & outputPath, LOG_LEVEL_INFO)
    
    Exit Function
    
ErrorHandler:
    If Not wbOut Is Nothing Then
        wbOut.Close SaveChanges:=False
        Set wbOut = Nothing
    End If
    GenerateOutput = ""
    Call LogMessage("GenerateOutput Error: " & Err.Description, LOG_LEVEL_ERROR)
    
End Function

'--------------------------------------------------
' 列文字を列番号に変換
'--------------------------------------------------
Public Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    On Error GoTo ErrorHandler
    
    colLetter = UCase(Trim(colLetter))
    
    If colLetter = "" Then
        ColumnLetterToNumber = 0
        Exit Function
    End If
    
    ' 数値の場合はそのまま返す
    If IsNumeric(colLetter) Then
        ColumnLetterToNumber = CLng(colLetter)
        Exit Function
    End If
    
    ' 列文字を番号に変換
    ColumnLetterToNumber = Range(colLetter & "1").Column
    Exit Function
    
ErrorHandler:
    ColumnLetterToNumber = 0
End Function
