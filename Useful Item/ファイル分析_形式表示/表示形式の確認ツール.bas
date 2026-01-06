Option Explicit

'===============================================================================
' モジュール名: ModFormatInventory (v3 - 改善版)
' 【概要】指定範囲のセルの表示形式を一覧化する機能を提供します。
' 【変更】2026-01-04 フリーズ対策(レポートシート事前作成、EnableEvents制御、
'         YIELD_INTERVAL、DEBUG_LOG追加)
' 【特徴】InputBox範囲指定、チャンク処理、エラーハンドリング対応
'===============================================================================

' チャンク処理の間隔
Const CHUNK_SIZE As Long = 500
' 一定時間ごとにUIを返す(秒)
Const YIELD_INTERVAL As Double = 0.2
' 進捗ログ出力 (Immediateウィンドウ)
Const DEBUG_LOG As Boolean = False

' レポートシート名ベース
Const REPORT_SHEET_BASE As String = "表示形式レポート"

Sub CreateFormatInventoryReport()
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim rngTarget As Range
    Dim c As Range
    Dim dict As Object
    Dim cats As Object
    Dim samples As Object
    Dim key As Variant
    Dim cellCounter As Long
    Dim lastYield As Double
    Dim totalCells As Double
    Dim lastCellAddress As String
    Dim lastStage As String
    
    On Error GoTo ErrorHandler
    
    ' 1. 辞書オブジェクトの初期化
    Set dict = CreateObject("Scripting.Dictionary")
    Set cats = CreateObject("Scripting.Dictionary")
    Set samples = CreateObject("Scripting.Dictionary")
    
    ' 2. ソースシートの設定
    Set wsSource = ActiveSheet
    
    ' 3. InputBox で処理範囲を選択させる
    On Error Resume Next
    Set rngTarget = Application.InputBox( _
        Prompt:="表示形式を調査する範囲を選択してください。" & vbCrLf & _
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
    
    ' 出力シートを先に作成 (終盤のフリーズ回避)
    Set wsReport = CreateReportSheet(REPORT_SHEET_BASE)
    
    ' 4. 処理開始 (画面更新停止)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "表示形式を抽出中... (Escキーでキャンセル)"
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = False
    
    cellCounter = 0
    lastYield = Timer
    totalCells = rngTarget.Cells.Count
    
    ' 5. 表示形式を収集
    For Each c In rngTarget
        cellCounter = cellCounter + 1
        
        ' チャンク処理: 一定セル数 or 一定時間ごとに DoEvents を実行
        If cellCounter Mod CHUNK_SIZE = 0 Or (Timer - lastYield) > YIELD_INTERVAL Then
            Application.StatusBar = "処理中... " & cellCounter & " / " & totalCells & _
                " セル処理済み | 直近: " & lastCellAddress & " " & lastStage
            If DEBUG_LOG Then Debug.Print "Progress: " & cellCounter & "/" & totalCells
            DoEvents
            lastYield = Timer
        End If
        
        lastCellAddress = c.Address(False, False)
        lastStage = "NumberFormat"
        
        key = c.NumberFormat
        
        ' カウント
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
            cats.Add key, InferCategory(CStr(key))
            samples.Add key, ""
        End If
        
        ' サンプルセル（最大20件まで）
        If SamplesCount(samples(key)) < 20 Then
            If Len(samples(key)) > 0 Then
                samples(key) = samples(key) & ", " & wsSource.Name & "!" & c.Address(False, False)
            Else
                samples(key) = wsSource.Name & "!" & c.Address(False, False)
            End If
        End If
    Next c
    
    ' 結果がない場合
    If dict.Count = 0 Then
        Application.DisplayAlerts = False
        wsReport.Delete
        Application.DisplayAlerts = True
        MsgBox "対象範囲にセルが見つかりませんでした。", vbInformation
        GoTo NormalExit
    End If
    
    ' 6. 結果を配列に変換してソート
    Dim arrKeys() As Variant
    Dim arrCounts() As Long
    Dim i As Long, j As Long
    
    ReDim arrKeys(0 To dict.Count - 1)
    ReDim arrCounts(0 To dict.Count - 1)
    
    i = 0
    For Each key In dict.Keys
        arrKeys(i) = key
        arrCounts(i) = dict(key)
        i = i + 1
    Next key
    
    ' バブルソート（件数の降順）
    For i = LBound(arrCounts) To UBound(arrCounts) - 1
        For j = i + 1 To UBound(arrCounts)
            If arrCounts(i) < arrCounts(j) Then
                SwapLong arrCounts(i), arrCounts(j)
                SwapVar arrKeys(i), arrKeys(j)
            End If
        Next j
    Next i
    
    ' 7. レポート出力
    lastStage = "Output:Start"
    Application.StatusBar = "出力準備中... " & dict.Count & " 種類"
    DoEvents
    
    ' ヘッダー作成
    lastStage = "Output:Header"
    Application.StatusBar = "出力中(ヘッダー)..."
    DoEvents
    With wsReport
        .Range("A1").Value = "表示形式"
        .Range("B1").Value = "推定カテゴリ"
        .Range("C1").Value = "件数"
        .Range("D1").Value = "代表セル"
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(220, 230, 241)
        
        ' データ出力
        lastStage = "Output:Write"
        Application.StatusBar = "出力中(書き込み)... " & dict.Count & " 種類"
        DoEvents
        Dim rowOut As Long
        rowOut = 2
        For i = LBound(arrKeys) To UBound(arrKeys)
            key = arrKeys(i)
            .Cells(rowOut, 1).Value = key
            .Cells(rowOut, 2).Value = cats(key)
            .Cells(rowOut, 3).Value = dict(key)
            .Cells(rowOut, 4).Value = samples(key)
            rowOut = rowOut + 1
        Next i
        
        lastStage = "Output:Format"
        Application.StatusBar = "出力中(整形)..."
        DoEvents
        .Columns("A:D").AutoFit
    End With
    
    MsgBox "抽出完了！ " & dict.Count & " 種類の表示形式が見つかりました。" & vbCrLf & _
           "合計セル数: " & TotalCount(dict), vbInformation

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
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    End If
    Resume NormalExit
End Sub

' ---------------------------------------------------------
' Helper: レポートシート作成
' ---------------------------------------------------------
Function CreateReportSheet(baseName As String) As Worksheet
    Dim sheetName As String
    Dim sheetIndex As Integer
    
    If Worksheets.Count >= 255 Then
        Err.Raise vbObjectError + 1000, , "シート数が上限に近いためレポートを作成できません。"
    End If
    
    sheetName = baseName & "_" & Format(Now, "hhmmss")
    sheetIndex = 1
    Do While SheetExists(sheetName)
        sheetIndex = sheetIndex + 1
        sheetName = baseName & "_" & Format(Now, "hhmmss") & "_" & sheetIndex
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

' 表示形式からカテゴリを推測
Private Function InferCategory(ByVal nf As String) As String
    Dim s As String
    s = LCase$(Trim$(nf))
    s = Replace(s, "  ", " ")
    
    ' テキスト
    If InStr(s, "@") > 0 And Not HasDateTokens(s) Then
        InferCategory = "テキスト"
        Exit Function
    End If
    
    ' 通貨
    If InStr(s, "$") > 0 Or InStr(s, "¥") > 0 Or InStr(s, "_(") > 0 Then
        InferCategory = "通貨"
        Exit Function
    End If
    
    ' パーセント
    If InStr(s, "%") > 0 Then
        InferCategory = "パーセント"
        Exit Function
    End If
    
    ' 分数
    If InStr(s, "/") > 0 And InStr(s, "m") = 0 And InStr(s, "d") = 0 Then
        InferCategory = "分数"
        Exit Function
    End If
    
    ' 指数
    If InStr(s, "e+") > 0 Or InStr(s, "e-") > 0 Then
        InferCategory = "指数"
        Exit Function
    End If
    
    ' 日付・時刻
    If HasDateTokens(s) Then
        If InStr(s, "h") > 0 Or InStr(s, "s") > 0 Or InStr(s, "午後") > 0 Then
            InferCategory = "日付時刻"
        Else
            InferCategory = "日付"
        End If
        Exit Function
    End If
    
    ' 時刻のみ
    If InStr(s, "h") > 0 Or InStr(s, "s") > 0 Or InStr(s, "午後") > 0 Then
        InferCategory = "時刻"
        Exit Function
    End If
    
    ' 桁区切り
    If InStr(s, "#,") > 0 Or InStr(s, "0,") > 0 Then
        InferCategory = "数値（桁区切り）"
        Exit Function
    End If
    
    ' 標準・数値
    If s = "general" Or s = "g/標準" Then
        InferCategory = "標準"
    ElseIf InStr(s, "0") > 0 Or InStr(s, "#") > 0 Then
        InferCategory = "数値"
    Else
        InferCategory = "その他"
    End If
End Function

' 日付トークンの有無をチェック
Private Function HasDateTokens(ByVal s As String) As Boolean
    HasDateTokens = (InStr(s, "y") > 0 Or InStr(s, "m") > 0 Or InStr(s, "d") > 0 Or _
                     InStr(s, "年") > 0 Or InStr(s, "月") > 0 Or InStr(s, "日") > 0)
End Function

' サンプルの件数をカウント
Private Function SamplesCount(ByVal csv As String) As Long
    Dim parts() As String
    Dim i As Long
    If Len(csv) = 0 Then
        SamplesCount = 0
    Else
        parts = Split(csv, ",")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(i))) > 0 Then
                SamplesCount = SamplesCount + 1
            End If
        Next i
    End If
End Function

' 合計件数を計算
Private Function TotalCount(ByVal d As Object) As Long
    Dim k As Variant, s As Long
    s = 0
    For Each k In d.Keys
        s = s + d(k)
    Next k
    TotalCount = s
End Function

' Long値のスワップ
Private Sub SwapLong(ByRef a As Long, ByRef b As Long)
    Dim t As Long: t = a: a = b: b = t
End Sub

' Variant値のスワップ
Private Sub SwapVar(ByRef a As Variant, ByRef b As Variant)
    Dim t As Variant: t = a: a = b: b = t
End Sub
