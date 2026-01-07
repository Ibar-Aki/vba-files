'==================================================
' modDataProcessor - データ処理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' データ結合処理
'--------------------------------------------------
Public Function MergeData(ByVal data1 As Object, _
                        ByVal data2 As Object) As Object
    
    Dim mergedDict As Object
    Dim id As Variant
    Dim only1 As Collection
    Dim only2 As Collection
    Dim matched As Long
    Dim excel2IdColNum As Long
    
    On Error GoTo ErrorHandler
    
    Set mergedDict = CreateObject("Scripting.Dictionary")
    Set only1 = New Collection
    Set only2 = New Collection
    
    ' 結果格納用
    Set mergedDict("MergedRows") = CreateObject("Scripting.Dictionary")
    mergedDict("Headers1") = data1("Headers")
    mergedDict("Headers2") = data2("Headers")
    
    ' Excel2の識別コード列番号を保存（出力時に使用）
    excel2IdColNum = data2("IDColumnNum")
    mergedDict("Excel2IDColumnNum") = excel2IdColNum
    
    matched = 0
    
    ' Excel1のデータを処理
    For Each id In data1("Data").Keys
        If data2("Data").Exists(id) Then
            ' 両方に存在
            mergedDict("MergedRows")(id) = MergeRow( _
                data1("Data")(id)("Data"), _
                data2("Data")(id)("Data"), _
                data1("LastCol"), _
                data2("LastCol"), _
                excel2IdColNum)
            matched = matched + 1
        Else
            ' Excel1のみ
            mergedDict("MergedRows")(id) = MergeRow( _
                data1("Data")(id)("Data"), _
                Empty, _
                data1("LastCol"), _
                data2("LastCol"), _
                excel2IdColNum)
            only1.Add id
        End If
        
        ' 進捗表示
        If matched Mod 500 = 0 Then
            DoEvents
        End If
    Next id
    
    ' Excel2のみのデータを処理
    For Each id In data2("Data").Keys
        If Not data1("Data").Exists(id) Then
            mergedDict("MergedRows")(id) = MergeRow( _
                Empty, _
                data2("Data")(id)("Data"), _
                data1("LastCol"), _
                data2("LastCol"), _
                excel2IdColNum)
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
                   " Excel2のみ: " & only2.Count, LOG_LEVEL_INFO)
    
    ' クリーンアップ
    Set only1 = Nothing
    Set only2 = Nothing
    
    Exit Function
    
ErrorHandler:
    Set MergeData = Nothing
    Set only1 = Nothing
    Set only2 = Nothing
    Call LogMessage("MergeData Error: " & Err.Description, LOG_LEVEL_ERROR)
    
End Function

'--------------------------------------------------
' 行データ結合
'--------------------------------------------------
Private Function MergeRow(ByVal row1 As Variant, _
                        ByVal row2 As Variant, _
                        ByVal col1 As Long, _
                        ByVal col2 As Long, _
                        ByVal excel2IdCol As Long) As Variant
    
    Dim mergedRow As Variant
    Dim i As Long
    Dim j As Long
    Dim totalCols As Long
    
    On Error GoTo ErrorHandler
    
    ' Excel2の識別コード列を除いた列数を計算
    totalCols = col1 + col2 - 1
    
    ' 結合配列作成（事前に初期化）
    ReDim mergedRow(1 To 1, 1 To totalCols)
    
    ' 初期値として空文字を設定
    For i = 1 To totalCols
        mergedRow(1, i) = ""
    Next i
    
    ' Excel1データコピー
    If Not IsEmpty(row1) Then
        For i = 1 To col1
            mergedRow(1, i) = row1(1, i)
        Next i
    End If
    
    ' Excel2データコピー（識別コード列を除く）
    If Not IsEmpty(row2) Then
        j = col1 + 1
        For i = 1 To col2
            ' 識別コード列はスキップ（動的な列番号を使用）
            If i <> excel2IdCol Then
                mergedRow(1, j) = row2(1, i)
                j = j + 1
            End If
        Next i
    End If
    
    MergeRow = mergedRow
    Exit Function
    
ErrorHandler:
    ' エラー時も初期化済み配列を返す
    Call LogMessage("MergeRow Error: " & Err.Description, LOG_LEVEL_ERROR)
    MergeRow = mergedRow
    
End Function

'--------------------------------------------------
' コレクションを配列に変換
'--------------------------------------------------
Public Function CollectionToArray(ByVal col As Collection) As Variant
    Dim arr() As String
    Dim i As Long
    
    If col Is Nothing Then
        CollectionToArray = Array()
        Exit Function
    End If
    
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

'--------------------------------------------------
' 配列の最初のN件を取得
'--------------------------------------------------
Public Function GetFirstN(ByVal sourceArray As Variant, ByVal n As Long) As Variant
    Dim result() As String
    Dim i As Long
    Dim upperBound As Long
    
    On Error GoTo ErrorHandler
    
    If Not IsArray(sourceArray) Then
        GetFirstN = Array()
        Exit Function
    End If
    
    On Error Resume Next
    upperBound = UBound(sourceArray)
    If Err.Number <> 0 Then
        GetFirstN = Array()
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    If upperBound < 1 Then
        GetFirstN = Array()
        Exit Function
    End If
    
    ' 実際に取得する件数を決定
    n = Application.Min(n, upperBound)
    
    ReDim result(1 To n)
    
    For i = 1 To n
        result(i) = CStr(sourceArray(i))
    Next i
    
    GetFirstN = result
    Exit Function
    
ErrorHandler:
    GetFirstN = Array()
End Function
