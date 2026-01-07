'==================================================
' modFileHandler - t@CW[
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' Excelf[^Ǎ
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
    
    ' ݒl擾
    headerRows = GetConfigValueLong(configSection & "_HeaderRows", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_HEADER_ROWS, DEFAULT_EXCEL2_HEADER_ROWS))
    dataStartRow = GetConfigValueLong(configSection & "_DataStartRow", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_DATA_START_ROW, DEFAULT_EXCEL2_DATA_START_ROW))
    idColumn = GetConfigValue(configSection & "_IDColumn", _
                    IIf(configSection = "Excel1", DEFAULT_EXCEL1_ID_COLUMN, DEFAULT_EXCEL2_ID_COLUMN))
    
    ' ԍɕϊ
    idColNum = ColumnLetterToNumber(idColumn)
    If idColNum = 0 Then
        Call LogMessage("ȗw: " & idColumn, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' t@CI[v
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    
    If wb.Worksheets.Count = 0 Then
        Call LogMessage("[NV[g݂܂: " & filePath, LOG_LEVEL_ERROR)
        wb.Close False
        Exit Function
    End If
    
    Set ws = wb.Worksheets(1)
    
    ' f[^͈͓
    lastRow = Application.Max(ws.Cells(ws.Rows.Count, idColNum).End(xlUp).Row, dataStartRow)
    lastCol = Application.Max(ws.Cells(headerRows, ws.Columns.Count).End(xlToLeft).Column, 1)
    
    ' ős`FbN
    If (lastRow - dataStartRow + 1) > MAX_PROCESSING_ROWS Then
        Call LogMessage("f[^s𒴂Ă܂ (" & _
                       (lastRow - dataStartRow + 1) & "s > " & MAX_PROCESSING_ROWS & "s)", LOG_LEVEL_WARNING)
    End If
    
    ' Dictionary쐬
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    ' wb_[擾
    dataDict("Headers") = GetHeaders(ws, headerRows, lastCol)
    dataDict("HeaderRows") = headerRows
    dataDict("LastCol") = lastCol
    dataDict("IDColumn") = idColumn
    dataDict("IDColumnNum") = idColNum  ' ԍۑ
    
    ' f[^擾
    Set dataDict("Data") = GetDataWithID(ws, dataStartRow, lastRow, idColNum, lastCol)
    
    ' t@C
    dataDict("FileName") = GetFileName(filePath)
    dataDict("RowCount") = Application.Max(0, lastRow - dataStartRow + 1)
    
    ' N[Y
    wb.Close False
    Set wb = Nothing
    
    Set LoadExcelData = dataDict
    
    Call LogMessage(configSection & " Ǎ: " & _
                   dataDict("RowCount") & "", LOG_LEVEL_INFO)
    
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
' wb_[擾iZΉj
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
    
    ' zOɏ
    ReDim headers(1 To headerRows, 1 To lastCol)
    
    For i = 1 To headerRows
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If cell.MergeCells Then
                ' Z̏ꍇ͍̒lgp
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
    ' G[ςݔzԂ
    Call LogMessage("GetHeaders Error: " & Err.Description, LOG_LEVEL_ERROR)
    GetHeaders = headers
    
End Function

'--------------------------------------------------
' f[^擾iʃR[htj
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
        
        ' ID̓XLbv
        If idValue <> "" Then
            ' d`FbN
            If dataDict.Exists(idValue) Then
                If Not duplicates.Exists(idValue) Then
                    duplicates(idValue) = CStr(dataDict(idValue)("Row")) & "," & CStr(i)
                Else
                    duplicates(idValue) = duplicates(idValue) & "," & CStr(i)
                End If
            Else
                ' f[^i[
                rowData = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Value
                
                Set rowDict = CreateObject("Scripting.Dictionary")
                rowDict("Data") = rowData
                rowDict("Row") = i
                Set dataDict(idValue) = rowDict
                
                validCount = validCount + 1
            End If
        End If
        
        ' i\i1000Ɓj
        If i Mod 1000 = 0 Then
            DoEvents
        End If
    Next i
    
    ' dΌx
    If duplicates.Count > 0 Then
        For Each key In duplicates.Keys
            Call LogMessage("ʃR[hd: " & key & _
                          " (s: " & duplicates(key) & ")", LOG_LEVEL_WARNING)
        Next key
    End If
    
    Call LogMessage("Lf[^: " & validCount & "", LOG_LEVEL_INFO)
    
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
' o̓t@C
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
    
    ' VK[NubN쐬
    Set wbOut = Workbooks.Add
    Set wsData = wbOut.Worksheets(1)
    wsData.Name = "f[^"
    
    ' wb_[쐬
    headers1 = mergedData("Headers1")
    headers2 = mergedData("Headers2")
    
    ' Excel2̎ʃR[hԍ擾
    excel2IdColNum = mergedData("Excel2IDColumnNum")
    If excel2IdColNum = 0 Then excel2IdColNum = 1  ' ftHg1
    
    ' ŏIwb_[s
    lastHeaderRow = Application.Max(UBound(headers1, 1), UBound(headers2, 1))
    
    ' wb_[óiZčŏIŝݏój
    row = 1
    col = 1
    
    ' Excel1wb_[
    For i = 1 To UBound(headers1, 2)
        wsData.Cells(row, col).Value = headers1(UBound(headers1, 1), i)
        col = col + 1
    Next i
    
    ' Excel2wb_[iʃR[hj
    For i = 1 To UBound(headers2, 2)
        ' ʃR[h̓XLbviݒ肩瓮IɎ擾j
        If i <> excel2IdColNum Then
            wsData.Cells(row, col).Value = headers2(UBound(headers2, 1), i)
            col = col + 1
        End If
    Next i
    
    ' wb_[ݒ
    With wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, col - 1))
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BG
        .Borders.LineStyle = xlContinuous
    End With
    
    ' f[^o
    row = 2
    For Each id In mergedData("MergedRows").Keys
        rowData = mergedData("MergedRows")(id)
        
        For col = 1 To UBound(rowData, 2)
            wsData.Cells(row, col).Value = rowData(1, col)
        Next col
        
        row = row + 1
        
        ' i\i100Ɓj
        If row Mod 100 = 0 Then
            DoEvents
        End If
    Next id
    
    ' 񕝎
    wsData.Cells.EntireColumn.AutoFit
    
    ' OV[g쐬
    includeLogSheet = GetConfigValueBool(CFG_INCLUDE_LOG_SHEET, DEFAULT_INCLUDE_LOG_SHEET)
    If includeLogSheet Then
        Call CreateLogSheet(wbOut, mergedData("Statistics"))
    End If
    
    ' t@Cۑ
    outputPath = GetOutputPath()
    If Not ValidateOutputPath(outputPath) Then
        wbOut.Close False
        Set wbOut = Nothing
        GenerateOutput = ""
        Exit Function
    End If
    wbOut.SaveAs outputPath, xlOpenXMLWorkbook
    wbOut.Close False
    Set wbOut = Nothing
    
    GenerateOutput = outputPath
    Call LogMessage("o͊: " & outputPath, LOG_LEVEL_INFO)
    
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
' 񕶎ԍɕϊ
'--------------------------------------------------
Public Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    On Error GoTo ErrorHandler
    
    colLetter = UCase(Trim(colLetter))
    
    If colLetter = "" Then
        ColumnLetterToNumber = 0
        Exit Function
    End If
    
    ' l̏ꍇ͂̂܂ܕԂ
    If IsNumeric(colLetter) Then
        ColumnLetterToNumber = CLng(colLetter)
        Exit Function
    End If
    
    ' 񕶎ԍɕϊ
    Dim i As Long
    Dim charCode As Long
    Dim result As Long
    
    result = 0
    For i = 1 To Len(colLetter)
        charCode = Asc(Mid(colLetter, i, 1))
        If charCode < Asc("A") Or charCode > Asc("Z") Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        result = result * 26 + (charCode - Asc("A") + 1)
    Next i
    
    ColumnLetterToNumber = result
    Exit Function
    
ErrorHandler:
    ColumnLetterToNumber = 0
End Function
