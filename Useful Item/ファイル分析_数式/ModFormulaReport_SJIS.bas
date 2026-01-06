Option Explicit
'===============================================================================
' モジュール名: ModFormulaReport (v3 - 改善版)
' 役割      : 指定範囲の数式セルを抽出し、専用シートに一覧として出力する
' 変更      : 2026-01-04 フリーズ対策(レポートシート事前作成、EnableEvents制御、
'             YIELD_INTERVAL、DEBUG_LOG追加)
' 特徴      : InputBox範囲指定、チャンク処理、配列一括出力、エラーハンドリング
'===============================================================================

' チャンク処理の間隔
Const CHUNK_SIZE As Long = 500
' 一定時間ごとにUIを返す(秒)
Const YIELD_INTERVAL As Double = 0.2
' 進捗ログ出力 (Immediateウィンドウ)
Const DEBUG_LOG As Boolean = False

' レポートシート名ベース
Const REPORT_SHEET_BASE As String = "数式レポート"

Sub CreateFormulaSummaryReport()
    Dim wsSource As Worksheet
    Dim wsReport As Worksheet
    Dim rngTarget As Range
    Dim targetCell As Range
    Dim resultData() As Variant
    Dim resultCount As Long
    Dim cellCounter As Long
    Dim lastYield As Double
    Dim totalCells As Double
    Dim lastCellAddress As String
    Dim lastStage As String
    
    On Error GoTo ErrorHandler
    
    ' 1. ソースシートの設定
    Set wsSource = ActiveSheet
    
    ' 2. InputBox で処理範囲を選択させる
    On Error Resume Next
    Set rngTarget = Application.InputBox( _
        Prompt:="数式を抽出する範囲を選択してください。" & vbCrLf & _
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
    
    ' 3. 配列の初期化
    Dim estimatedMaxRows As Long
    estimatedMaxRows = rngTarget.Cells.Count
    If estimatedMaxRows > 1048576 Then estimatedMaxRows = 1048576
    
    ' 出力シートを先に作成 (終盤のフリーズ回避)
    Set wsReport = CreateReportSheet(REPORT_SHEET_BASE)
    
    ReDim resultData(1 To estimatedMaxRows, 1 To 5)
    resultCount = 0
    cellCounter = 0
    lastYield = Timer
    totalCells = rngTarget.Cells.Count
    
    ' 4. 処理開始 (画面更新停止)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "数式を抽出中... (Escキーでキャンセル)"
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = False
    
    ' 5. 数式セルを抽出
    For Each targetCell In rngTarget
        cellCounter = cellCounter + 1
        
        ' チャンク処理: 一定セル数 or 一定時間ごとに DoEvents を実行
        If cellCounter Mod CHUNK_SIZE = 0 Or (Timer - lastYield) > YIELD_INTERVAL Then
            Application.StatusBar = "処理中... " & cellCounter & " / " & totalCells & _
                " セル処理済み | 直近: " & lastCellAddress & " " & lastStage
            If DEBUG_LOG Then Debug.Print "Progress: " & cellCounter & "/" & totalCells
            DoEvents
            lastYield = Timer
        End If
        
        ' 数式セルの場合のみ処理
        If targetCell.HasFormula Then
            lastCellAddress = targetCell.Address(False, False)
            lastStage = "Formula"
            
            resultCount = resultCount + 1
            
            If resultCount > UBound(resultData, 1) Then
                MsgBox "結果がシートの行数制限を超えそうになったため、一部のデータがカットされました。", vbExclamation
                GoTo OutputReport
            End If
            
            resultData(resultCount, 1) = "'" & Replace(targetCell.Worksheet.Name, "'", "''") & "'"
            resultData(resultCount, 2) = targetCell.Worksheet.Name
            resultData(resultCount, 3) = targetCell.Address(False, False)
            resultData(resultCount, 4) = targetCell.Text
            resultData(resultCount, 5) = targetCell.Formula
        End If
    Next targetCell

OutputReport:
    ' 6. レポート出力
    If resultCount > 0 Then
        lastStage = "Output:Start"
        Application.StatusBar = "出力準備中... " & resultCount & " 件"
        DoEvents
        
        ' ヘッダー作成
        lastStage = "Output:Header"
        Application.StatusBar = "出力中(ヘッダー)..."
        DoEvents
        With wsReport
            .Range("A1").Value = "シート参照"
            .Range("B1").Value = "シート名"
            .Range("C1").Value = "セル位置"
            .Range("D1").Value = "表示値"
            .Range("E1").Value = "数式"
            .Range("A1:E1").Font.Bold = True
            .Range("A1:E1").Interior.Color = RGB(220, 230, 241)
            
            ' データ一括貼り付け
            lastStage = "Output:Write"
            Application.StatusBar = "出力中(書き込み)... " & resultCount & " 件"
            DoEvents
            .Range("A2").Resize(resultCount, 5).Value = resultData
            
            lastStage = "Output:Format"
            Application.StatusBar = "出力中(整形)..."
            DoEvents
            .Columns("A:E").AutoFit
        End With
        
        MsgBox "抽出完了！ " & resultCount & " 件の数式が見つかりました。", vbInformation
    Else
        ' 結果がない場合は作成済みシートを削除
        Application.DisplayAlerts = False
        wsReport.Delete
        Application.DisplayAlerts = True
        MsgBox "対象範囲に数式は見つかりませんでした。", vbInformation
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
