'==================================================
' modLogger - ログ処理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' ログ初期化
'--------------------------------------------------
Public Sub InitializeLog()
    Set LogCollection = New Collection
End Sub

'--------------------------------------------------
' ログメッセージ追加
'--------------------------------------------------
Public Sub LogMessage(ByVal message As String, _
                     ByVal logLevel As String)
    
    Dim logEntry As Object
    Dim formattedTime As String
    
    On Error Resume Next
    
    Set logEntry = CreateObject("Scripting.Dictionary")
    
    logEntry("Timestamp") = Now
    logEntry("Level") = logLevel
    logEntry("Message") = message
    
    ' ログコレクションに追加
    If Not LogCollection Is Nothing Then
        LogCollection.Add logEntry
    End If
    
    ' イミディエイトウィンドウに出力（デバッグ用）
    formattedTime = Format(Now, TIMESTAMP_FORMAT_FULL)
    Debug.Print formattedTime & " [" & logLevel & "] " & message
    
    On Error GoTo 0
    
End Sub

'--------------------------------------------------
' ログシート作成
'--------------------------------------------------
Public Function CreateLogSheet(ByVal wb As Workbook, _
                             ByVal stats As Object) As Worksheet
    
    Dim ws As Worksheet
    Dim row As Long
    Dim i As Long
    Dim key As Variant
    Dim logEntry As Variant
    Dim tempArr() As String
    Dim displayCount As Long
    Dim remainingCount As Long
    
    On Error GoTo ErrorHandler
    
    ' ログシート追加
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "処理ログ"
    
    ' ヘッダー設定
    With ws
        ' タイトル
        .Range("A1").Value = APP_TITLE & " 処理ログ"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        ' 基本情報セクション
        row = 3
        .Range("A" & row & ":B" & row).Font.Bold = True
        .Range("A" & row).Value = "項目"
        .Range("B" & row).Value = "内容"
        
        row = row + 1
        .Cells(row, 1).Value = "処理日時"
        .Cells(row, 2).Value = Format(Now, TIMESTAMP_FORMAT_FULL)
        row = row + 1
        
        .Cells(row, 1).Value = "システムバージョン"
        .Cells(row, 2).Value = APP_VERSION
        row = row + 1
        
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
        
        ' 識別コードリスト（Excel1のみ）
        row = row + 1
        If stats("Only1Count") > 0 Then
            .Cells(row, 1).Value = "Excel1のみ識別コード"
            
            If stats("Only1Count") <= MAX_DISPLAY_IDS Then
                .Cells(row, 2).Value = Join(stats("Only1IDs"), ", ")
            Else
                ' 最初のN件のみ表示
                tempArr = GetFirstN(stats("Only1IDs"), MAX_DISPLAY_IDS)
                remainingCount = stats("Only1Count") - MAX_DISPLAY_IDS
                .Cells(row, 2).Value = Join(tempArr, ", ") & " ... (他" & remainingCount & "件)"
            End If
            row = row + 1
        End If
        
        ' 識別コードリスト（Excel2のみ）
        If stats("Only2Count") > 0 Then
            .Cells(row, 1).Value = "Excel2のみ識別コード"
            
            If stats("Only2Count") <= MAX_DISPLAY_IDS Then
                .Cells(row, 2).Value = Join(stats("Only2IDs"), ", ")
            Else
                ' 最初のN件のみ表示
                tempArr = GetFirstN(stats("Only2IDs"), MAX_DISPLAY_IDS)
                remainingCount = stats("Only2Count") - MAX_DISPLAY_IDS
                .Cells(row, 2).Value = Join(tempArr, ", ") & " ... (他" & remainingCount & "件)"
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
        .Range("A" & row & ":C" & row).Interior.Color = COLOR_HEADER_BG
        row = row + 1
        
        ' ログ出力
        If Not LogCollection Is Nothing Then
            For Each logEntry In LogCollection
                .Cells(row, 1).Value = Format(logEntry("Timestamp"), TIMESTAMP_FORMAT_TIME)
                .Cells(row, 2).Value = logEntry("Level")
                .Cells(row, 3).Value = logEntry("Message")
                
                ' レベルによって色分け
                Select Case logEntry("Level")
                    Case LOG_LEVEL_ERROR
                        .Range("B" & row & ":C" & row).Font.Color = COLOR_ERROR_TEXT
                    Case LOG_LEVEL_WARNING
                        .Range("B" & row & ":C" & row).Font.Color = COLOR_WARNING_TEXT
                End Select
                
                row = row + 1
            Next logEntry
        End If
        
        ' 列幅調整
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 60
        
        ' 罫線
        If row > 4 Then
            .Range("A3:C" & (row - 1)).Borders.LineStyle = xlContinuous
        End If
        
    End With
    
    Set CreateLogSheet = ws
    Exit Function
    
ErrorHandler:
    Call LogMessage("CreateLogSheet Error: " & Err.Description, LOG_LEVEL_ERROR)
    Set CreateLogSheet = Nothing
    
End Function

'--------------------------------------------------
' ログをテキストファイルに出力
'--------------------------------------------------
Public Sub ExportLogToFile(ByVal filePath As String)
    Dim fileNum As Integer
    Dim logEntry As Variant
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    Print #fileNum, APP_TITLE & " v" & APP_VERSION & " - 処理ログ"
    Print #fileNum, "出力日時: " & Format(Now, TIMESTAMP_FORMAT_FULL)
    Print #fileNum, String(60, "-")
    Print #fileNum, ""
    
    If Not LogCollection Is Nothing Then
        For Each logEntry In LogCollection
            Print #fileNum, Format(logEntry("Timestamp"), TIMESTAMP_FORMAT_FULL) & _
                           " [" & logEntry("Level") & "] " & logEntry("Message")
        Next logEntry
    End If
    
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Sub
