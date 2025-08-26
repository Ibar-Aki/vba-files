Option Explicit

'================================================================================
' Outlook予定取得マクロ (Version 1.3 - 分類/区分 判定機能を追加)
'
' 追加点（v1.3）:
' ・件名に含まれるキーワード行を KeyMatrix / KeyMatrix_区分 から検索し、
'   一致した行の ClassList / ClassList_区分 をそれぞれ F列(分類) / H列(区分) に出力
' ・Excel 2016向けに配列処理で高速化（式は書き込まない＝安定稼働）
'
' 前提の名前付き範囲（同ブック内・既存）：
'   - KeyMatrix        : キーワード行列（複数列可／空白可）
'   - ClassList        : 1列。KeyMatrix 各行に対応する分類名
'   - KeyMatrix_区分   : キーワード行列（複数列可／空白可）
'   - ClassList_区分   : 1列。KeyMatrix_区分 各行に対応する区分名
'================================================================================
Sub GetOutlookSchedule()

    '============================================================
    ' ■ 1. 初期設定と定数宣言
    '============================================================
    ' --- ユーザー設定項目 ---
    Const TARGET_SHEET_NAME As String = "データ取得"      ' マクロを実行するシート名
    Const DATE_INPUT_CELL   As String = "C3"              ' 日付が入力されているセル番地
    Const OUTPUT_HEADER_ROW As Long = 7                   ' ヘッダー行
    Const OUTPUT_START_COLUMN As String = "C"             ' 出力先の開始列（C）

    ' --- 追加機能に関する設定項目 ---
    Const DEST_SHEET_NAME As String = "データ登録"        ' 転記先のシート名
    Const SOURCE_CELL     As String = "C4"                ' 転記元（データ取得）
    Const DEST_CELL       As String = "D4"                ' 転記先（データ登録）

    ' --- 列番号（固定） ---
    Const COL_TIME As Long      = 3   ' C:時間
    Const COL_SUBJECT As Long   = 4   ' D:件名
    Const COL_DURATION As Long  = 5   ' E:会議時間（"HHMM"）
    Const COL_CLASS As Long     = 6   ' F:分類（追加）
    Const COL_RESERVED As Long  = 7   ' G:未使用（予約）
    Const COL_KUBUN As Long     = 8   ' H:区分（追加）

    ' 変数
    Dim ws As Worksheet, wsDest As Worksheet
    Dim wasProtected As Boolean
    Dim olApp As Object, olNs As Object, olFolder As Object
    Dim olItems As Object, olRestrictedItems As Object, olApt As Object
    Dim targetDate As Date, outputRow As Long, lastOutputRow As Long
    Dim actualCount As Long, cellValue As Variant

    ' ---- 分類・区分 判定用（名前付き範囲を配列に確保）----
    Dim rngKey As Range, rngClass As Range
    Dim rngKeyKbn As Range, rngClassKbn As Range
    Dim arrKey As Variant, arrClass As Variant
    Dim arrKeyKbn As Variant, arrClassKbn As Variant
    Dim enableClass As Boolean, enableKbn As Boolean
    Dim warnMsg As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    '============================================================
    ' ■ 2. 実行前チェックと準備
    '============================================================
    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    ' シート保護解除（必要時）
    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect
        If Err.Number <> 0 Then
            On Error GoTo ErrorHandler
            Dim userPassword As String
            userPassword = InputBox("シートがパスワードで保護されています。パスワードを入力してください:", "パスワード入力")
            If userPassword = "" Then
                MsgBox "パスワードが入力されませんでした。処理を中止します。", vbExclamation
                GoTo CleanUp
            End If
            ws.Unprotect Password:=userPassword
        End If
        On Error GoTo ErrorHandler
    End If

    ' 日付入力セル
    cellValue = ws.Range(DATE_INPUT_CELL).Value
    If IsEmpty(cellValue) Or cellValue = "" Then
        MsgBox "セル " & DATE_INPUT_CELL & " が空欄です。日付を入力してください。", vbExclamation, "入力エラー"
        GoTo CleanUp
    End If
    targetDate = CDate(cellValue)

    '============================================================
    ' ■ 3. 出力範囲クリア & ヘッダー
    '============================================================
    outputRow = OUTPUT_HEADER_ROW + 1
    lastOutputRow = ws.Cells(ws.Rows.Count, COL_TIME).End(xlUp).Row
    If lastOutputRow < OUTPUT_HEADER_ROW Then lastOutputRow = OUTPUT_HEADER_ROW

    If lastOutputRow >= outputRow Then
        ws.Range(ws.Cells(outputRow, COL_TIME), ws.Cells(lastOutputRow, COL_KUBUN)).ClearContents
    End If

    ' ヘッダ設定
    ws.Cells(OUTPUT_HEADER_ROW, COL_TIME).Value = "時間"
    ws.Cells(OUTPUT_HEADER_ROW, COL_SUBJECT).Value = "件名"
    ws.Cells(OUTPUT_HEADER_ROW, COL_DURATION).Value = "会議時間"
    ws.Cells(OUTPUT_HEADER_ROW, COL_CLASS).Value = "分類"
    ws.Cells(OUTPUT_HEADER_ROW, COL_KUBUN).Value = "区分"
    ws.Range(ws.Cells(OUTPUT_HEADER_ROW, COL_TIME), ws.Cells(OUTPUT_HEADER_ROW, COL_KUBUN)).Font.Bold = True

    '============================================================
    ' ■ 4. 名前付き範囲の取得（存在しなければ無効化）
    '============================================================
    enableClass = TryGetNamedRange("KeyMatrix", rngKey, warnMsg) _
                  And TryGetNamedRange("ClassList", rngClass, warnMsg)
    If enableClass Then
        If rngClass.Columns.Count <> 1 Or rngClass.Rows.Count <> rngKey.Rows.Count Then
            warnMsg = warnMsg & vbCrLf & "ClassList は1列で、KeyMatrix と同じ行数である必要があります。分類判定をスキップします。"
            enableClass = False
        End If
    End If
    If enableClass Then
        arrKey = To2DArray(rngKey)
        arrClass = To2DArray(rngClass)
    End If

    enableKbn = TryGetNamedRange("KeyMatrix_区分", rngKeyKbn, warnMsg) _
                And TryGetNamedRange("ClassList_区分", rngClassKbn, warnMsg)
    If enableKbn Then
        If rngClassKbn.Columns.Count <> 1 Or rngClassKbn.Rows.Count <> rngKeyKbn.Rows.Count Then
            warnMsg = warnMsg & vbCrLf & "ClassList_区分 は1列で、KeyMatrix_区分 と同じ行数である必要があります。区分判定をスキップします。"
            enableKbn = False
        End If
    End If
    If enableKbn Then
        arrKeyKbn = To2DArray(rngKeyKbn)
        arrClassKbn = To2DArray(rngClassKbn)
    End If

    '============================================================
    ' ■ 5. Outlook 接続
    '============================================================
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    If olApp Is Nothing Then Err.Raise vbObjectError, , "Outlookに接続できません"

    '============================================================
    ' ■ 6. 指定日の予定を取得
    '============================================================
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(9) ' 9=olFolderCalendar
    Set olItems = olFolder.Items
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True

    Dim filterString As String
    filterString = "[Start] <= '" & Format(targetDate, "yyyy/MM/dd 23:59") & "' AND [End] >= '" & Format(targetDate, "yyyy/MM/dd 00:00") & "'"
    Set olRestrictedItems = olItems.Restrict(filterString)

    '============================================================
    ' ■ 7. Excelへ出力（分類・区分の判定を追加）
    '============================================================
    If olRestrictedItems.Count = 0 Then
        ws.Cells(outputRow, COL_TIME).Value = "予定はありません"
        MsgBox Format(targetDate, "yyyy年mm月dd日") & " の予定はありませんでした。", vbInformation, "処理完了"
    Else
        actualCount = 0
        Dim subj As String, className As String, kubunName As String
        For Each olApt In olRestrictedItems
            actualCount = actualCount + 1

            ' 時間・件名・会議時間
            ws.Cells(outputRow, COL_TIME).Value = Format(olApt.Start, "hhmm") & "-" & Format(olApt.End, "hhmm")
            subj = NzString(olApt.Subject)
            ws.Cells(outputRow, COL_SUBJECT).Value = subj

            Dim totalMinutes As Long, hours As Long, minutes As Long
            totalMinutes = DateDiff("n", olApt.Start, olApt.End)
            hours = totalMinutes \ 60
            minutes = totalMinutes Mod 60
            With ws.Cells(outputRow, COL_DURATION)
                .NumberFormat = "@"
                .Value = Format(hours, "00") & Format(minutes, "00")
            End With

            ' 分類（F）／区分（H）
            className = ""
            If enableClass Then className = ResolveClassByKeyMatrix(subj, arrKey, arrClass)
            ws.Cells(outputRow, COL_CLASS).Value = className

            kubunName = ""
            If enableKbn Then kubunName = ResolveClassByKeyMatrix(subj, arrKeyKbn, arrClassKbn)
            ws.Cells(outputRow, COL_KUBUN).Value = kubunName

            outputRow = outputRow + 1
        Next olApt

        Dim doneMsg As String
        doneMsg = Format(targetDate, "yyyy年mm月dd日") & " の予定を " & actualCount & " 件取得しました。"
        If Len(warnMsg) > 0 Then
            doneMsg = doneMsg & vbCrLf & "（注意）" & vbCrLf & Trim$(warnMsg)
        End If
        MsgBox doneMsg, vbInformation, "処理完了"
    End If

    '============================================================
    ' ■ 8. データ転記処理（C4 → データ登録!D4）
    '============================================================
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET_NAME)
    On Error GoTo ErrorHandler

    If Not wsDest Is Nothing Then
        If NzString(ws.Range(SOURCE_CELL).Value) <> "" Then
            ' ※必要なら wsDest の保護解除/再保護を追加
            wsDest.Range(DEST_CELL).Value = ws.Range(SOURCE_CELL).Value
        End If
    End If

    GoTo CleanUp

'======================== エラーハンドラ/終了処理 ==============================
ErrorHandler:
    Dim errorTitle As String, errorMsg As String
    errorTitle = "エラーが発生しました"
    Select Case Err.Number
        Case 9
            errorMsg = "シート「" & TARGET_SHEET_NAME & "」または「" & DEST_SHEET_NAME & "」が見つかりませんでした。"
        Case 13
            errorMsg = "セル「" & DATE_INPUT_CELL & "」の値を日付として認識できません。"
        Case 287, -2147467259, -2147221233
            errorMsg = "Outlookへのアクセスで問題が発生しました。"
        Case vbObjectError
            errorMsg = "Outlookアプリケーションの起動に失敗しました。"
        Case Else
            errorMsg = "予期しないエラーが発生しました。" & vbCrLf & _
                       "エラー番号: " & Err.Number & vbCrLf & _
                       "エラー内容: " & Err.Description
    End Select
    MsgBox errorMsg, vbCritical, errorTitle
    '（エラーでも後始末へ）

CleanUp:
    If Not ws Is Nothing And wasProtected Then
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If

    On Error Resume Next
    Set olApt = Nothing
    Set olRestrictedItems = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    Set wsDest = Nothing
    Set ws = Nothing
    Application.ScreenUpdating = True
End Sub

' 実行用（ボタン登録など）
Sub ExecuteOutlookSchedule()
    Call GetOutlookSchedule
End Sub

'======================== 補助関数 ============================

' 名前付き範囲を取得。存在しなければ False を返し、warn に追記
Private Function TryGetNamedRange(ByVal nameStr As String, ByRef rng As Range, ByRef warn As String) As Boolean
    On Error Resume Next
    Set rng = Nothing
    Set rng = ThisWorkbook.Names(nameStr).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        TryGetNamedRange = False
        warn = warn & IIf(Len(warn) > 0, vbCrLf, "") & "名前付き範囲 """ & nameStr & """ が見つかりません。"
    Else
        TryGetNamedRange = True
    End If
End Function

' Range→2次元Variant配列に安全変換（単一セルでも(1,1)始まりにする）
Private Function To2DArray(ByVal rng As Range) As Variant
    Dim v As Variant
    If rng.Cells.Count = 1 Then
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = rng.Value
        To2DArray = v
    Else
        To2DArray = rng.Value
    End If
End Function

' Null/Empty を空文字にする
Private Function NzString(ByVal v As Variant) As String
    If IsError(v) Then
        NzString = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzString = ""
    Else
        NzString = CStr(v)
    End If
End Function

' 件名に対する行マッチ→対応クラス名を返す（未一致は ""）
'   keysArr  : (1..R, 1..C) のキーワード行列。空白セルは無視
'   classArr : (1..R, 1..1) のクラス名縦ベクトル
Private Function ResolveClassByKeyMatrix(ByVal subjectText As String, _
                                         ByRef keysArr As Variant, _
                                         ByRef classArr As Variant) As String
    Dim r As Long, c As Long
    Dim rows As Long, cols As Long
    Dim kw As String

    If IsEmpty(keysArr) Then Exit Function
    rows = UBound(keysArr, 1)
    cols = UBound(keysArr, 2)

    For r = 1 To rows
        ' 行内に1つでもヒットがあれば採用（最初の一致を返す）
        For c = 1 To cols
            kw = NzString(keysArr(r, c))
            If Len(kw) > 0 Then
                If InStr(1, subjectText, kw, vbTextCompare) > 0 Then
                    ResolveClassByKeyMatrix = NzString(classArr(r, 1))
                    Exit Function
                End If
            End If
        Next c
    Next r

    ' 未一致は空文字
    ResolveClassByKeyMatrix = ""
End Function
