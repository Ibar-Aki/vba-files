'==================================================
' Excel結合処理システム VBAコード集
' Version: 1.0
' Date: 2025/07/16
'==================================================

'==================================================
' ThisWorkbook モジュール
'==================================================
Private Sub Workbook_Open()
    ' ワークブックを非表示にする
    Application.Windows(ThisWorkbook.Name).Visible = False
End Sub

'==================================================
' modMain - メイン処理モジュール
'==================================================
Option Explicit

' グローバル変数
Public g_ConfigData As Object
Public g_LogCollection As Collection

'--------------------------------------------------
' メインエントリーポイント（外部から呼び出される）
'--------------------------------------------------
Public Sub ExecuteMerge(ByVal strFile1 As String, ByVal strFile2 As String)
    
    Dim startTime As Date
    Dim result As Boolean
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    
    ' 初期化
    startTime = Now
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' ログ初期化
    Call InitializeLog
    Call LogMessage("========== 処理開始 ==========", "INFO")
    Call LogMessage("Excel1: " & GetFileName(strFile1), "INFO")
    Call LogMessage("Excel2: " & GetFileName(strFile2), "INFO")
    
    ' 設定読込
    If Not LoadConfiguration() Then
        Call LogMessage("設定ファイルの読み込みに失敗しました", "ERROR")
        GoTo Cleanup
    End If
    
    ' ファイル検証
    If Not ValidateFiles(strFile1, strFile2) Then
        Call LogMessage("ファイル検証エラー", "ERROR")
        GoTo Cleanup
    End If
    
    ' データ処理実行
    result = ProcessMerge(strFile1, strFile2)
    
    If result Then
        Call LogMessage("処理完了 処理時間: " & _
            Format(Now - startTime, "hh:mm:ss"), "INFO")
        Call LogMessage("========== 処理終了 ==========", "INFO")
    Else
        Call LogMessage("処理失敗", "ERROR")
    End If
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    ' 処理完了メッセージ
    If result Then
        MsgBox "処理が完了しました。" & vbCrLf & _
               "出力フォルダを確認してください。", _
               vbInformation, "Excel結合処理"
    Else
        MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
               "ログを確認してください。", _
               vbExclamation, "Excel結合処理"
    End If
    
    ' 自己終了
    For Each wb In Workbooks
        If wb.Name = ThisWorkbook.Name Then
            wb.Close SaveChanges:=False
        End If
    Next wb
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage("システムエラー: " & Err.Description, "ERROR")
    Resume Cleanup
    
End Sub

'--------------------------------------------------
' 結合処理メイン
'--------------------------------------------------
Private Function ProcessMerge(ByVal file1 As String, ByVal file2 As String) As Boolean
    
    Dim data1 As Object, data2 As Object
    Dim mergedData As Object
    Dim outputPath As String
    
    On Error GoTo ErrorHandler
    
    ProcessMerge = False
    
    ' Excel1読込
    Call LogMessage("Excel1読込開始...", "INFO")
    Set data1 = LoadExcelData(file1, "Excel1")
    If data1 Is Nothing Then
        Exit Function
    End If
    
    ' Excel2読込
    Call LogMessage("Excel2読込開始...", "INFO")
    Set data2 = LoadExcelData(file2, "Excel2")
    If data2 Is Nothing Then
        Exit Function
    End If
    
    ' データ結合
    Call LogMessage("データ結合処理開始...", "INFO")
    Set mergedData = MergeData(data1, data2)
    If mergedData Is Nothing Then
        Exit Function
    End If
    
    ' 結果を追加情報として保存
    mergedData("File1Name") = GetFileName(file1)
    mergedData("File2Name") = GetFileName(file2)
    
    ' 出力
    Call LogMessage("ファイル出力開始...", "INFO")
    outputPath = GenerateOutput(mergedData)
    
    ProcessMerge = (outputPath <> "")
    
    If ProcessMerge Then
        Call LogMessage("出力ファイル: " & outputPath, "INFO")
    End If
    
    Exit Function
    
ErrorHandler:
    ProcessMerge = False
    Call LogMessage("ProcessMerge Error: " & Err.Description, "ERROR")
    
End Function

'--------------------------------------------------
' ファイル名取得
'--------------------------------------------------
Public Function GetFileName(ByVal fullPath As String) As String
    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        GetFileName = Mid(fullPath, pos + 1)
    Else
        GetFileName = fullPath
    End If
End Function

'==================================================
' modFileHandler - ファイル処理モジュール
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
    Dim lastRow As Long, lastCol As Long
    Dim idColNum As Long
    
    On Error GoTo ErrorHandler
    
    Set LoadExcelData = Nothing
    
    ' 設定値取得
    headerRows = CLng(g_ConfigData(configSection & "_HeaderRows"))
    dataStartRow = CLng(g_ConfigData(configSection & "_DataStartRow"))
    idColumn = g_ConfigData(configSection & "_IDColumn")
    
    ' 列番号に変換
    idColNum = Range(idColumn & "1").Column
    
    ' ファイルオープン
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    
    ' データ範囲特定
    lastRow = ws.Cells(ws.Rows.Count, idColNum).End(xlUp).Row
    If lastRow < dataStartRow Then
        lastRow = dataStartRow
    End If
    
    lastCol = ws.Cells(headerRows, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then
        lastCol = 1
    End If
    
    ' Dictionary作成
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    ' ヘッダー情報取得
    dataDict("Headers") = GetHeaders(ws, headerRows, lastCol)
    dataDict("HeaderRows") = headerRows
    dataDict("LastCol") = lastCol
    
    ' データ取得
    Set dataDict("Data") = GetDataWithID(ws, dataStartRow, lastRow, _
                                        idColNum, lastCol)
    
    ' ファイル情報
    dataDict("FileName") = GetFileName(filePath)
    dataDict("RowCount") = Application.Max(0, lastRow - dataStartRow + 1)
    dataDict("IDColumn") = idColumn
    
    ' クローズ
    wb.Close False
    
    Set LoadExcelData = dataDict
    
    Call LogMessage(configSection & " 読込完了: " & _
                   dataDict("RowCount") & "件", "INFO")
    
    Exit Function
    
ErrorHandler:
    If Not wb Is Nothing Then wb.Close False
    Set LoadExcelData = Nothing
    Call LogMessage("LoadExcelData Error: " & Err.Description & _
                   " (File: " & filePath & ")", "ERROR")
    
End Function

'--------------------------------------------------
' ヘッダー取得（セル結合対応）
'--------------------------------------------------
Private Function GetHeaders(ByVal ws As Worksheet, _
                          ByVal headerRows As Long, _
                          ByVal lastCol As Long) As Variant
    
    Dim headers() As String
    Dim i As Long, j As Long
    Dim cell As Range
    Dim mergeArea As Range
    
    On Error GoTo ErrorHandler
    
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
    Call LogMessage("GetHeaders Error: " & Err.Description, "ERROR")
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
    Dim i As Long
    Dim idValue As String
    Dim rowData As Variant
    Dim duplicates As Object
    Dim validCount As Long
    
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
                
                Set dataDict(idValue) = CreateObject("Scripting.Dictionary")
                dataDict(idValue)("Data") = rowData
                dataDict(idValue)("Row") = i
                validCount = validCount + 1
            End If
        End If
    Next i
    
    ' 重複があれば警告
    If duplicates.Count > 0 Then
        Dim key As Variant
        For Each key In duplicates.Keys
            Call LogMessage("識別コード重複: " & key & _
                          " (行: " & duplicates(key) & ")", "WARNING")
        Next key
    End If
    
    Call LogMessage("有効データ件数: " & validCount & "件", "INFO")
    
    Set GetDataWithID = dataDict
    Exit Function
    
ErrorHandler:
    Call LogMessage("GetDataWithID Error: " & Err.Description, "ERROR")
    Set GetDataWithID = dataDict
    
End Function

'--------------------------------------------------
' 出力ファイル生成
'--------------------------------------------------
Public Function GenerateOutput(ByVal mergedData As Object) As String
    
    Dim wbOut As Workbook
    Dim wsData As Worksheet
    Dim outputPath As String
    Dim row As Long, col As Long
    Dim headers1 As Variant, headers2 As Variant
    Dim id As Variant
    Dim i As Long
    Dim lastHeaderRow As Long
    
    On Error GoTo ErrorHandler
    
    GenerateOutput = ""
    
    ' 新規ワークブック作成
    Set wbOut = Workbooks.Add
    Set wsData = wbOut.Worksheets(1)
    wsData.Name = "結合データ"
    
    ' ヘッダー作成
    headers1 = mergedData("Headers1")
    headers2 = mergedData("Headers2")
    
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
        ' 識別コード列はスキップ（最初の列と仮定）
        If i > 1 Then
            wsData.Cells(row, col).Value = headers2(UBound(headers2, 1), i)
            col = col + 1
        End If
    Next i
    
    ' ヘッダー書式設定
    With wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, col - 1))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' データ出力
    row = 2
    For Each id In mergedData("MergedRows").Keys
        Dim rowData As Variant
        rowData = mergedData("MergedRows")(id)
        
        For col = 1 To UBound(rowData, 2)
            wsData.Cells(row, col).Value = rowData(1, col)
        Next col
        
        row = row + 1
    Next id
    
    ' 列幅自動調整
    wsData.Cells.EntireColumn.AutoFit
    
    ' ログシート作成
    If CBool(g_ConfigData("Output_IncludeLogSheet")) Then
        Call CreateLogSheet(wbOut, mergedData("Statistics"))
    End If
    
    ' ファイル保存
    outputPath = GetOutputPath()
    wbOut.SaveAs outputPath, xlOpenXMLWorkbook
    wbOut.Close
    
    GenerateOutput = outputPath
    Call LogMessage("出力完了: " & outputPath, "INFO")
    
    Exit Function
    
ErrorHandler:
    If Not wbOut Is Nothing Then
        wbOut.Close SaveChanges:=False
    End If
    GenerateOutput = ""
    Call LogMessage("GenerateOutput Error: " & Err.Description, "ERROR")
    
End Function

'==================================================
' modDataProcessor - データ処理モジュール
'==================================================
Option Explicit

'--------------------------------------------------
' データ結合処理
'--------------------------------------------------
Public Function MergeData(ByVal data1 As Object, _
                        ByVal data2 As Object) As Object
    
    Dim mergedDict As Object
    Dim id As Variant
    Dim only1 As New Collection
    Dim only2 As New Collection
    Dim matched As Long
    
    On Error GoTo ErrorHandler
    
    Set mergedDict = CreateObject("Scripting.Dictionary")
    
    ' 結果格納用
    Set mergedDict("MergedRows") = CreateObject("Scripting.Dictionary")
    mergedDict("Headers1") = data1("Headers")
    mergedDict("Headers2") = data2("Headers")
    
    matched = 0
    
    ' Excel1のデータを処理
    For Each id In data1("Data").Keys
        If data2("Data").Exists(id) Then
            ' 両方に存在
            mergedDict("MergedRows")(id) = MergeRow( _
                data1("Data")(id)("Data"), _
                data2("Data")(id)("Data"), _
                data1("LastCol"), _
                data2("LastCol"))
            matched = matched + 1
        Else
            ' Excel1のみ
            mergedDict("MergedRows")(id) = MergeRow( _
                data1("Data")(id)("Data"), _
                Empty, _
                data1("LastCol"), _
                data2("LastCol"))
            only1.Add id
        End If
    Next id
    
    ' Excel2のみのデータを処理
    For Each id In data2("Data").Keys
        If Not data1("Data").Exists(id) Then
            mergedDict("MergedRows")(id) = MergeRow( _
                Empty, _
                data2("Data")(id)("Data"), _
                data1("LastCol"), _
                data2("LastCol"))
            only2.Add id
        End If
    Next id
    
    ' 統計情報
    Set mergedDict("Statistics") = CreateObject("Scripting.Dictionary")
    mergedDict("Statistics")("Excel1Count") = data1("Data").Count
    mergedDict("Statistics")("Excel2Count") = data2("Data").Count
    mergedDict("Statistics")("MatchedCount") = matched
    mergedDict("Statistics")("Only1Count") = only1.Count
    mergedDict("Statistics")("Only2Count") = only2.Count
    mergedDict("Statistics")("Only1IDs") = CollectionToArray(only1)
    mergedDict("Statistics")("Only2IDs") = CollectionToArray(only2)
    
    Set MergeData = mergedDict
    
    ' ログ出力
    Call LogMessage("結合完了 - 一致: " & matched & _
                   " Excel1のみ: " & only1.Count & _
                   " Excel2のみ: " & only2.Count, "INFO")
    
    Exit Function
    
ErrorHandler:
    Set MergeData = Nothing
    Call LogMessage("MergeData Error: " & Err.Description, "ERROR")
    
End Function

'--------------------------------------------------
' 行データ結合
'--------------------------------------------------
Private Function MergeRow(ByVal row1 As Variant, _
                        ByVal row2 As Variant, _
                        ByVal col1 As Long, _
                        ByVal col2 As Long) As Variant
    
    Dim mergedRow As Variant
    Dim i As Long
    Dim totalCols As Long
    
    On Error GoTo ErrorHandler
    
    ' Excel2の識別コード列を除いた列数を計算
    totalCols = col1 + col2 - 1
    
    ' 結合配列作成
    ReDim mergedRow(1 To 1, 1 To totalCols)
    
    ' Excel1データコピー
    If Not IsEmpty(row1) Then
        For i = 1 To col1
            mergedRow(1, i) = row1(1, i)
        Next i
    Else
        ' Excel1が空の場合は空文字を設定
        For i = 1 To col1
            mergedRow(1, i) = ""
        Next i
    End If
    
    ' Excel2データコピー（識別コード列を除く）
    If Not IsEmpty(row2) Then
        For i = 2 To col2  ' 最初の列（識別コード）をスキップ
            mergedRow(1, col1 + i - 1) = row2(1, i)
        Next i
    Else
        ' Excel2が空の場合は空文字を設定
        For i = 2 To col2
            mergedRow(1, col1 + i - 1) = ""
        Next i
    End If
    
    MergeRow = mergedRow
    Exit Function
    
ErrorHandler:
    Call LogMessage("MergeRow Error: " & Err.Description, "ERROR")
    MergeRow = mergedRow
    
End Function

'--------------------------------------------------
' コレクションを配列に変換
'--------------------------------------------------
Private Function CollectionToArray(ByVal col As Collection) As Variant
    Dim arr() As String
    Dim i As Long
    
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    ReDim arr(1 To col.Count)
    
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    
    CollectionToArray = arr
    
End Function

'==================================================
' modLogger - ログ処理モジュール
'==================================================
Option Explicit

'--------------------------------------------------
' ログ初期化
'--------------------------------------------------
Public Sub InitializeLog()
    Set g_LogCollection = New Collection
End Sub

'--------------------------------------------------
' ログメッセージ追加
'--------------------------------------------------
Public Sub LogMessage(ByVal message As String, _
                     ByVal logLevel As String)
    
    Dim logEntry As Object
    Set logEntry = CreateObject("Scripting.Dictionary")
    
    logEntry("Timestamp") = Now
    logEntry("Level") = logLevel
    logEntry("Message") = message
    
    g_LogCollection.Add logEntry
    
    ' イミディエイトウィンドウに出力（デバッグ用）
    Debug.Print Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                " [" & logLevel & "] " & message
    
End Sub

'--------------------------------------------------
' ログシート作成
'--------------------------------------------------
Public Function CreateLogSheet(ByVal wb As Workbook, _
                             ByVal stats As Object) As Worksheet
    
    Dim ws As Worksheet
    Dim row As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' ログシート追加
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "処理ログ"
    
    ' ヘッダー設定
    With ws
        ' タイトル
        .Range("A1").Value = "Excel結合処理ログ"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        ' 基本情報セクション
        row = 3
        .Range("A" & row & ":B" & row).Font.Bold = True
        .Range("A" & row).Value = "項目"
        .Range("B" & row).Value = "内容"
        
        row = row + 1
        .Cells(row, 1).Value = "処理日時"
        .Cells(row, 2).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
        row = row + 1
        
        ' ファイル情報（ファイル名が利用可能な場合）
        If Not IsEmpty(g_ConfigData) Then
            If g_ConfigData.Exists("File1Name") Then
                .Cells(row, 1).Value = "Excel1ファイル"
                .Cells(row, 2).Value = g_ConfigData("File1Name")
                row = row + 1
            End If
            If g_ConfigData.Exists("File2Name") Then
                .Cells(row, 1).Value = "Excel2ファイル"
                .Cells(row, 2).Value = g_ConfigData("File2Name")
                row = row + 1
            End If
        End If
        
        ' 統計情報
        row = row + 1
        .Range("A" & row).Value = "処理結果"
        .Range("A" & row).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "Excel1データ件数"
        .Cells(row, 2).Value = stats("Excel1Count") & " 件"
        row = row + 1
        
        .Cells(row, 1).Value = "Excel2データ件数"
        .Cells(row, 2).Value = stats("Excel2Count") & " 件"
        row = row + 1
        
        .Cells(row, 1).Value = "結合済みデータ件数"
        .Cells(row, 2).Value = stats("MatchedCount") & " 件"
        row = row + 1
        
        .Cells(row, 1).Value = "Excel1のみデータ件数"
        .Cells(row, 2).Value = stats("Only1Count") & " 件"
        row = row + 1
        
        .Cells(row, 1).Value = "Excel2のみデータ件数"
        .Cells(row, 2).Value = stats("Only2Count") & " 件"
        row = row + 1
        
        ' 識別コードリスト
        row = row + 1
        If stats("Only1Count") > 0 Then
            .Cells(row, 1).Value = "Excel1のみ識別コード"
            If stats("Only1Count") <= 20 Then
                .Cells(row, 2).Value = Join(stats("Only1IDs"), ", ")
            Else
                ' 最初の20件のみ表示
                Dim tempArr() As String
                ReDim tempArr(1 To 20)
                For i = 1 To 20
                    tempArr(i) = stats("Only1IDs")(i)
                Next i
                .Cells(row, 2).Value = Join(tempArr, ", ") & " ... (他" & _
                                      (stats("Only1Count") - 20) & "件)"
            End If
            row = row + 1
        End If
        
        If stats("Only2Count") > 0 Then
            .Cells(row, 1).Value = "Excel2のみ識別コード"
            If stats("Only2Count") <= 20 Then
                .Cells(row, 2).Value = Join(stats("Only2IDs"), ", ")
            Else
                ' 最初の20件のみ表示
                ReDim tempArr(1 To 20)
                For i = 1 To 20
                    tempArr(i) = stats("Only2IDs")(i)
                Next i
                .Cells(row, 2).Value = Join(tempArr, ", ") & " ... (他" & _
                                      (stats("Only2Count") - 20) & "件)"
            End If
            row = row + 1
        End If
        
        ' 処理ログ
        row = row + 2
        .Cells(row, 1).Value = "処理ログ"
        .Range("A" & row & ":C" & row).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "時刻"
        .Cells(row, 2).Value = "レベル"
        .Cells(row, 3).Value = "メッセージ"
        .Range("A" & row & ":C" & row).Interior.Color = RGB(217, 217, 217)
        row = row + 1
        
        ' ログ出力
        Dim logEntry As Variant
        For Each logEntry In g_LogCollection
            .Cells(row, 1).Value = Format(logEntry("Timestamp"), "hh:mm:ss")
            .Cells(row, 2).Value = logEntry("Level")
            .Cells(row, 3).Value = logEntry("Message")
            
            ' レベルによって色分け
            Select Case logEntry("Level")
                Case "ERROR"
                    .Range("B" & row & ":C" & row).Font.Color = RGB(255, 0, 0)
                Case "WARNING"
                    .Range("B" & row & ":C" & row).Font.Color = RGB(255, 140, 0)
            End Select
            
            row = row + 1
        Next logEntry
        
        ' 列幅調整
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 60
        
        ' 罫線
        .Range("A3:B" & (row - 1)).Borders.LineStyle = xlContinuous
        
    End With
    
    Set CreateLogSheet = ws
    Exit Function
    
ErrorHandler:
    Call LogMessage("CreateLogSheet Error: " & Err.Description, "ERROR")
    Set CreateLogSheet = Nothing
    
End Function

'==================================================
' modConfig - 設定管理モジュール
'==================================================
Option Explicit

'--------------------------------------------------
' 設定ファイル読込
'--------------------------------------------------
Public Function LoadConfiguration() As Boolean
    
    Dim configPath As String
    Dim wbConfig As Workbook
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    LoadConfiguration = False
    
    ' 設定ファイルパス
    configPath = ThisWorkbook.Path & "\Config\MergeConfig.xlsx"
    
    ' 設定ファイル存在確認
    If Not FileExists(configPath) Then
        Call LogMessage("設定ファイルが見つかりません: " & configPath, "WARNING")
        ' デフォルト設定使用
        Call SetDefaultConfig
        LoadConfiguration = True
        Exit Function
    End If
    
    ' 設定読込
    Set g_ConfigData = CreateObject("Scripting.Dictionary")
    Set wbConfig = Workbooks.Open(configPath, ReadOnly:=True, UpdateLinks:=0)
    Set wsConfig = wbConfig.Worksheets("Config")
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' ヘッダー行をスキップ
        If wsConfig.Cells(i, 1).Value <> "" Then
            g_ConfigData(CStr(wsConfig.Cells(i, 1).Value)) = _
                CStr(wsConfig.Cells(i, 2).Value)
        End If
    Next i
    
    wbConfig.Close False
    
    ' 必須項目チェック
    If Not ValidateConfig() Then
        Call SetDefaultConfig
    End If
    
    Call LogMessage("設定ファイル読込完了", "INFO")
    LoadConfiguration = True
    
    Exit Function
    
ErrorHandler:
    If Not wbConfig Is Nothing Then wbConfig.Close False
    Call LogMessage("設定ファイル読込エラー: " & Err.Description, "ERROR")
    Call SetDefaultConfig
    LoadConfiguration = True
    
End Function

'--------------------------------------------------
' デフォルト設定
'--------------------------------------------------
Private Sub SetDefaultConfig()
    
    Set g_ConfigData = CreateObject("Scripting.Dictionary")
    
    ' Excel1設定
    g_ConfigData("Excel1_HeaderRows") = "3"
    g_ConfigData("Excel1_DataStartRow") = "4"
    g_ConfigData("Excel1_IDColumn") = "B"
    
    ' Excel2設定
    g_ConfigData("Excel2_HeaderRows") = "2"
    g_ConfigData("Excel2_DataStartRow") = "3"
    g_ConfigData("Excel2_IDColumn") = "A"
    
    ' 出力設定
    g_ConfigData("Output_FileNameFormat") = "結合データ_[DATE].xlsx"
    g_ConfigData("Output_IncludeLogSheet") = "TRUE"
    
    Call LogMessage("デフォルト設定を使用します", "INFO")
    
End Sub

'--------------------------------------------------
' 設定検証
'--------------------------------------------------
Private Function ValidateConfig() As Boolean
    
    Dim requiredKeys As Variant
    Dim i As Long
    
    requiredKeys = Array("Excel1_HeaderRows", "Excel1_DataStartRow", _
                        "Excel1_IDColumn", "Excel2_HeaderRows", _
                        "Excel2_DataStartRow", "Excel2_IDColumn")
    
    ValidateConfig = True
    
    For i = 0 To UBound(requiredKeys)
        If Not g_ConfigData.Exists(requiredKeys(i)) Then
            Call LogMessage("必須設定項目が不足: " & requiredKeys(i), "WARNING")
            ValidateConfig = False
            Exit Function
        End If
    Next i
    
End Function

'--------------------------------------------------
' 出力パス生成
'--------------------------------------------------
Public Function GetOutputPath() As String
    
    Dim outputDir As String
    Dim fileName As String
    Dim dateStr As String
    
    ' 出力ディレクトリ
    outputDir = ThisWorkbook.Path & "\Output"
    
    ' ディレクトリ作成
    If Not FolderExists(outputDir) Then
        MkDir outputDir
    End If
    
    ' ファイル名生成
    dateStr = Format(Now, "yyyymmdd_HHmmss")
    fileName = g_ConfigData("Output_FileNameFormat")
    fileName = Replace(fileName, "[DATE]", dateStr)
    
    GetOutputPath = outputDir & "\" & fileName
    
End Function

'--------------------------------------------------
' ファイル存在確認
'--------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

'--------------------------------------------------
' フォルダ存在確認
'--------------------------------------------------
Public Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
End Function

'==================================================
' modValidator - 検証処理モジュール
'==================================================
Option Explicit

'--------------------------------------------------
' ファイル検証
'--------------------------------------------------
Public Function ValidateFiles(ByVal file1 As String, _
                            ByVal file2 As String) As Boolean
    
    ValidateFiles = False
    
    ' ファイル存在確認
    If Not FileExists(file1) Then
        Call LogMessage("Excel1が見つかりません: " & file1, "ERROR")
        Exit Function
    End If
    
    If Not FileExists(file2) Then
        Call LogMessage("Excel2が見つかりません: " & file2, "ERROR")
        Exit Function
    End If
    
    ' 拡張子確認
    If LCase(Right(file1, 5)) <> ".xlsx" Then
        Call LogMessage("Excel1は.xlsx形式である必要があります", "ERROR")
        Exit Function
    End If
    
    If LCase(Right(file2, 5)) <> ".xlsx" Then
        Call LogMessage("Excel2は.xlsx形式である必要があります", "ERROR")
        Exit Function
    End If
    
    ' 同一ファイルチェック
    If file1 = file2 Then
        Call LogMessage("同じファイルが指定されています", "ERROR")
        Exit Function
    End If
    
    ValidateFiles = True
    Call LogMessage("ファイル検証OK", "INFO")
    
End Function