Option Explicit

'===============================================================================
' モジュール名: ModFormatInventory
'
' 【概要】アクティブなブックに含まれる全てのワークシートを走査し、
'         使用されているセルの表示形式を一覧化する機能を提供します。
' 【作成】「業務支援ツール作成」2025/09
' 【対象環境】Excel 2016+ / Windows
' 【主要機能】
' ・ブック内の全使用セルから表示形式を収集し、重複を除いてリストアップ
' ・各表示形式が使用されているセル数をカウント
' ・結果を「FormatInventory」シートに件数の降順で出力
' ・処理の実行時間を計測し、完了時にメッセージで報告
'===============================================================================

'===============================================================================
' 【メイン処理】
' ブック全体の表示形式を棚卸しし、結果を専用シートに出力する
'===============================================================================
Public Sub ListAllNumberFormats()
    ' --- 変数宣言 ---
    Dim ws As Worksheet, outWs As Worksheet ' ループ用ワークシート、出力用ワークシート
    Dim dict As Object                      ' 表示形式(key)と出現回数(value)を格納
    Dim samples As Object                   ' 表示形式(key)とサンプルセルアドレス(value)を格納
    Dim cats As Object                      ' 表示形式(key)と推定カテゴリ(value)を格納
    Dim r As Range, c As Range, ur As Range ' ループ用のRangeオブジェクト
    Dim k As Variant                        ' Dictionaryのキーを格納するループ変数
    Dim rowOut As Long                      ' 出力シートの行カウンタ
    Dim nf As String, key As String         ' 取得した表示形式、正規化後の表示形式キー
    Dim t0 As Double, cnt As Long           ' 処理時間計測用、セルカウント用(未使用)
    
    ' --- ステップ1：Excel状態の保存と高速化設定 ---
    ' 処理中のパフォーマンス向上のため、画面更新や自動計算を一時的に停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    t0 = Timer ' 処理開始時間を記録
    
    ' --- ステップ2：データ格納用Dictionaryオブジェクトの初期化 ---
    Set dict = CreateObject("Scripting.Dictionary")      ' key: NumberFormatLocal, val: count
    Set samples = CreateObject("Scripting.Dictionary")   ' key: NumberFormatLocal, val: sample addresses (CSV, max 5)
    Set cats = CreateObject("Scripting.Dictionary")      ' key: NumberFormatLocal, val: inferred category
    
    ' --- ステップ3：全ワークシートを巡回し、表示形式を収集 ---
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next ' UsedRangeが取得できないシート(グラフシート等)をスキップ
        Set ur = ws.UsedRange
        On Error GoTo 0
        
        ' --- UsedRangeが存在するシートのみ処理 ---
        If Not ur Is Nothing Then
            For Each c In ur.Cells
                nf = CStr(c.NumberFormatLocal)
                
                ' --- 表示形式が設定されているセルのみを対象とする ---
                ' 空や同一書式の大量繰返しに対しても軽負荷化（使用範囲のみ走査）
                If Len(nf) > 0 Then
                    key = CanonNF(nf) ' 表示形式を正規化してキーとして使用
                    
                    ' --- 新しい表示形式の場合、各種Dictionaryに情報を追加 ---
                    If Not dict.Exists(key) Then
                        dict.Add key, 1
                        samples.Add key, ws.Name & "!" & c.Address(0, 0)
                        cats.Add key, InferCategory(nf)
                    Else
                        ' --- 既存の表示形式の場合、カウントを増やし、サンプルアドレスを追加 ---
                        dict(key) = dict(key) + 1
                        ' サンプルは最大5つまで追加
                        If SamplesCount(samples(key)) < 5 Then
                            samples(key) = samples(key) & ", " & ws.Name & "!" & c.Address(0, 0)
                        End If
                    End If
                End If
            Next c
        End If
        Set ur = Nothing ' メモリ解放
    Next ws
    
    ' --- ステップ4：出力用ワークシートの準備 ---
    On Error Resume Next
    Set outWs = ThisWorkbook.Worksheets("FormatInventory")
    On Error GoTo 0
    
    If outWs Is Nothing Then
        ' シートが存在しない場合は新規作成
        Set outWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        outWs.Name = "FormatInventory"
    Else
        ' シートが既に存在する場合は内容をクリア
        outWs.Cells.Clear
    End If
    
    ' --- ステップ5：結果の出力 ---
    With outWs
        ' --- 見出し行の作成 ---
        .Range("A1:E1").Value = Array("No.", "NumberFormatLocal（表示形式）", "推定カテゴリ", "件数", "代表セル（最大5件）")
        rowOut = 2 ' 出力開始行
        
        ' --- 件数で降順ソートするための準備 ---
        ' Dictionaryのキーと値を配列に格納
        Dim arrKeys() As Variant, arrCounts() As Long, i As Long, j As Long
        ReDim arrKeys(0 To dict.Count - 1)
        ReDim arrCounts(0 To dict.Count - 1)
        i = 0
        For Each k In dict.Keys
            arrKeys(i) = k
            arrCounts(i) = dict(k)
            i = i + 1
        Next k
        
        ' --- 簡易バブルソート（件数の降順） ---
        For i = LBound(arrCounts) To UBound(arrCounts) - 1
            For j = i + 1 To UBound(arrCounts)
                If arrCounts(i) < arrCounts(j) Then
                    SwapLong arrCounts(i), arrCounts(j) ' 件数を交換
                    SwapVar arrKeys(i), arrKeys(j)      ' 対応するキーも交換
                End If
            Next j
        Next i
        
        ' --- ソートされた結果をシートに書き出し ---
        For i = LBound(arrKeys) To UBound(arrKeys)
            k = arrKeys(i)
            .Cells(rowOut, 1).Value = rowOut - 1
            .Cells(rowOut, 2).Value = k ' 正規化後の表示（人間可読性重視でそのまま）
            .Cells(rowOut, 3).Value = cats(k)
            .Cells(rowOut, 4).Value = arrCounts(i)
            .Cells(rowOut, 5).Value = samples(k)
            rowOut = rowOut + 1
        Next i
        
        ' --- 体裁の調整 ---
        .Columns("A:E").AutoFit
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(230, 230, 230)
        .Activate
    End With
    
    ' --- ステップ6：Excel状態の復元と完了報告 ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "完了: 表示形式 " & dict.Count & " 種類、総セル数 " & TotalCount(dict) & _
           "（" & Format(Timer - t0, "0.0") & "s）", vbInformation, "処理完了"
End Sub

'===============================================================================
' 【内部ヘルパー関数・プロシージャ】
' メイン処理から呼び出される補助的な機能
'===============================================================================

'===============================================================================
' 【機能名】表示形式の簡易正規化
' 【概要】  全角/半角スペースや余分な空白を除去し、表記揺れを統一する
' 【引数】  nf: 元の表示形式文字列
' 【戻り値】String: 正規化された表示形式文字列
'===============================================================================
Private Function CanonNF(ByVal nf As String) As String
    Dim s As String
    s = Trim$(nf) ' 前後の空白を除去
    ' よくある冗長空白を削る（例："_ *"まわりは会計書式ゆえそのまま保持）
    s = Replace(s, "  ", " ") ' 連続する半角スペースを1つに
    CanonNF = s
End Function

'===============================================================================
' 【機能名】表示形式のカテゴリ推定
' 【概要】  表示形式の文字列から、そのカテゴリ（日付、数値など）を簡易的に判定する
' 【引数】  nf: 表示形式文字列
' 【戻り値】String: 推定されたカテゴリ名
'===============================================================================
Private Function InferCategory(ByVal nf As String) As String
    Dim s As String: s = LCase$(nf) ' 大文字小文字を区別せずに判定するため小文字に変換
    
    ' --- カテゴリ判定ロジック ---
    ' 明示的な文字列形式
    If s = "@" Or InStr(s, "@") > 0 And Not HasDateTokens(s) Then InferCategory = "文字列": Exit Function
    ' 標準形式
    If s = "general" Or s = "g/標準" Or InStr(s, "標準") > 0 Then InferCategory = "標準": Exit Function
    
    ' 日付・時刻関連のトークンを含むか
    If HasDateTokens(s) Then
        If InStr(s, "h") > 0 Or InStr(s, "時") > 0 Then
            InferCategory = "時刻/日時": Exit Function
        Else
            InferCategory = "日付": Exit Function
        End If
    End If
    
    ' パーセント形式
    If InStr(s, "%") > 0 Then InferCategory = "パーセント": Exit Function
    
    ' 通貨・会計形式
    If InStr(s, "[$") > 0 Or InStr(s, "\") > 0 Or InStr(s, "¥") > 0 Or InStr(s, "_(") > 0 Then
        InferCategory = "通貨/会計": Exit Function
    End If
    
    ' 分数形式
    If InStr(s, "?/?") > 0 Then InferCategory = "分数": Exit Function
    
    ' 指数形式
    If InStr(s, "e+0") > 0 Or InStr(s, "e-0") > 0 Then InferCategory = "指数": Exit Function
    
    ' 上記以外は数値と判定（桁区切りの有無で細分化）
    If InStr(s, "#,") > 0 Or InStr(s, "0,") > 0 Then
        InferCategory = "数値（桁区切り）"
    Else
        InferCategory = "数値"
    End If
End Function

'===============================================================================
' 【機能名】日付/時刻トークンの存在チェック
' 【概要】  与えられた文字列に日付や時刻を表す文字が含まれているかを判定する
' 【引数】  s: 判定対象の文字列 (小文字に変換済みを想定)
' 【戻り値】Boolean: 含まれていればTrue, なければFalse
'===============================================================================
Private Function HasDateTokens(ByVal s As String) As Boolean
    ' --- y, m, d, h, s や漢字、am/pmなどの日付・時刻関連の文字が含まれるかチェック ---
    HasDateTokens = (InStr(s, "y") > 0 Or InStr(s, "m") > 0 Or InStr(s, "d") > 0 Or _
                     InStr(s, "g") > 0 Or InStr(s, "年") > 0 Or InStr(s, "月") > 0 Or _
                     InStr(s, "日") > 0 Or InStr(s, "時") > 0 Or InStr(s, "分") > 0 Or _
                     InStr(s, "秒") > 0 Or InStr(s, "午前") > 0 Or InStr(s, "午後") > 0 Or _
                     InStr(s, "h") > 0 Or InStr(s, "s") > 0)
End Function


'===============================================================================
' 【小物ユーティリティ関数・プロシージャ】
' モジュール内で使用する汎用的な補助機能
'===============================================================================

'===============================================================================
' 【機能名】サンプル数のカウント
' 【概要】  カンマ区切りの文字列から要素数をカウントする
' 【引数】  csv: カンマ区切りの文字列
' 【戻り値】Long: 要素数
'===============================================================================
Private Function SamplesCount(ByVal csv As String) As Long
    If Len(csv) = 0 Then
        SamplesCount = 0
    Else
        SamplesCount = UBound(Split(csv, ",")) + 1
    End If
End Function

'===============================================================================
' 【機能名】長整数型(Long)変数の値交換
' 【概要】  2つのLong型変数の値を入れ替える
' 【引数】  a, b: 値を交換する変数 (参照渡し)
'===============================================================================
Private Sub SwapLong(ByRef a As Long, ByRef b As Long)
    Dim t As Long: t = a: a = b: b = t
End Sub

'===============================================================================
' 【機能名】Variant型変数の値交換
' 【概要】  2つのVariant型変数の値を入れ替える
' 【引数】  a, b: 値を交換する変数 (参照渡し)
'===============================================================================
Private Sub SwapVar(ByRef a As Variant, ByRef b As Variant)
    Dim t As Variant: t = a: a = b: b = t
End Sub

'===============================================================================
' 【機能名】Dictionaryの値の合計値計算
' 【概要】  Scripting.Dictionaryに格納された全ての値(数値)の合計を計算する
' 【引数】  d: 対象のDictionaryオブジェクト
' 【戻り値】Long: 計算された合計値
'===============================================================================
Private Function TotalCount(ByVal d As Object) As Long
    Dim k As Variant, s As Long
    s = 0 ' 合計値の初期化
    For Each k In d.Keys
        s = s + d(k)
    Next k
    TotalCount = s
End Function
