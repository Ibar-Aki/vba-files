'================================================================================
' 機能: 「データ登録」→「月次データ」へ転記（区分＋作番で列特定）
' 改訂: v2.2
'   - 【不具合#1】combinedKey のスコープ明示（全箇所で宣言）
'   - 【不具合#7】クリップボード: Forms.DataObject 優先＋WinAPIフォールバック＆リトライ/解放
'   - 【不具合#9】副作用対策: ScreenUpdating/EnableEvents/Calculation を退避→停止→復帰
'   - 追加機能: 自動列追加ポリシー（都度確認/全自動/一括拒否）
'   - 追加機能: 重複上書き検知（既存値あり→セルを黄色塗り＋メッセージ列へ注記）
'   - その他: 時刻パース強化(H:MM 文字列), 日付行の動的探索, 新列のデータ行に [h]:mm 書式
' 対象: Excel 2016 / Windows 11 / 日本語環境
'================================================================================
Option Explicit

'=========================
' WinAPI（64/32両対応）
'=========================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_UNICODETEXT As Long = 13

'=========================
' シート/レイアウト設定
'=========================
Private Const DATA_SHEET_NAME As String = "データ登録"
Private Const ACQUISITION_SHEET_NAME As String = "データ取得"
Private Const MONTHLY_SHEET_NAME As String = "月次データ"

Private Const DATE_CELL_PRIORITY As String = "D4" ' 任意日付
Private Const DATE_CELL_NORMAL  As String = "D3" ' 登録日

Private Const DATA_START_ROW As Long = 8
Private Const MONTHLY_WORKNO_ROW As Long = 8   ' 作番行
Private Const MONTHLY_HEADER_ROW As Long = 9   ' 区分行（ヘッダ）
Private Const MONTHLY_DATA_START_ROW As Long = 10

Private Const DATE_COL As String = "B"        ' 月次の日付列
Private Const MESSAGE_COL As String = "A"     ' メッセージ列（注記の出力先）

'=========================
' 動作ポリシー
'=========================
' 列追加ポリシー: 0=都度確認 / 1=全自動 / 2=一括拒否
' 列追加ポリシー（Enumの代わりにConstで定義?Excel 2016/VBE環境差対策）
Private Const AddPolicy_Prompt As Long = 0
Private Const AddPolicy_Auto   As Long = 1
Private Const AddPolicy_Reject As Long = 2

Private Const AUTO_ADD_POLICY As Long = AddPolicy_Prompt
Private Const ACCUMULATE_MODE As Boolean = True   ' True=加算, False=上書き
Private Const DRY_RUN As Boolean = False          ' True=プレビューのみ

' 重複上書き検知
Private Const DUP_HIGHLIGHT_COLOR As Long = vbYellow ' 既存値あり時の塗り色

'=========================
' メイン処理
'=========================
Public Sub TransferDataToMonthlySheet()
    Dim wsData As Worksheet, wsMonthly As Worksheet
    Dim wasProt As Boolean, usedPw As String
    Dim targetDate As Date, targetRow As Long
    Dim lastDataRow As Long, i As Long
    Dim dicMap As Object            ' Key: "区分|作番" → 列番号
    Dim agg As Object               ' Key: "区分|作番" → 合計分
    Dim combinedKey As String

    ' 速度/副作用対策（#9）
    Dim prevUpd As Boolean, prevEvt As Boolean, prevCalc As XlCalculation
    prevUpd = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevEvt = Application.EnableEvents:    Application.EnableEvents = False
    prevCalc = Application.Calculation:    Application.Calculation = xlCalculationManual

    On Error GoTo EH

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)

    ' 保護解除
    wasProt = wsMonthly.ProtectContents
    If wasProt Then
        If Not UnprotectSheetEx(wsMonthly, usedPw) Then GoTo Clean
    End If

    ' 日付の決定
    If IsDate(wsData.Range(DATE_CELL_PRIORITY).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_PRIORITY).Value)
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).Value)
    Else
        MsgBox "登録日（D3）または任意日付（D4）が無効です。", vbExclamation
        GoTo Clean
    End If

    ' 転記元最終行
    lastDataRow = wsData.Cells(wsData.rows.Count, "C").End(xlUp).Row
    If lastDataRow < DATA_START_ROW Then MsgBox "転記するデータがありません。", vbInformation: GoTo Clean

    ' ヘッダ辞書（区分|作番 → 列）
    Set dicMap = CreateObject("Scripting.Dictionary")
    Dim lastHeaderCol As Long, c As Long
    lastHeaderCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    For c = 3 To lastHeaderCol ' C列～
        Dim categoryName As String, workNoName As String
        categoryName = Trim$(CStr(wsMonthly.Cells(MONTHLY_HEADER_ROW, c).Value))
        workNoName = Trim$(CStr(wsMonthly.Cells(MONTHLY_WORKNO_ROW, c).Value))
        If categoryName <> "" Then
            combinedKey = categoryName & "|" & workNoName ' 【#1 明示宣言済】
            If Not dicMap.Exists(combinedKey) Then dicMap.Add combinedKey, c
        End If
    Next c

    ' 日付行（動的）
    targetRow = FindMatchingDateRowDyn(wsMonthly, targetDate)
    If targetRow = 0 Then
        MsgBox "月次シートのB列に対象日が見つかりません。", vbExclamation
        GoTo Clean
    End If

    ' 集計（分）
    Set agg = CreateObject("Scripting.Dictionary")
    For i = DATA_START_ROW To lastDataRow
        Dim workNo As String, cat As String, mins As Double
        workNo = Trim$(CStr(wsData.Cells(i, "C").Value))
        cat = Trim$(CStr(wsData.Cells(i, "D").Value))
        mins = ConvertToMinutesEx(wsData.Cells(i, "E").Value)
        If mins > 0 And workNo <> "" And cat <> "" Then
            combinedKey = cat & "|" & workNo   ' 【#1 明示宣言済】
            If agg.Exists(combinedKey) Then agg(combinedKey) = agg(combinedKey) + mins _
                                     Else agg.Add combinedKey, mins
        End If
    Next
    If agg.Count = 0 Then MsgBox "有効な時間データがありません。", vbInformation: GoTo Clean

    ' プレビュー
    Dim preview As String, k As Variant
    preview = "以下の内容で転記します。よろしいですか？" & vbCrLf & vbCrLf & _
              "日付: " & Format$(targetDate, "yyyy/mm/dd") & vbCrLf & _
              "---------------------------------------------" & vbCrLf & _
              "作番 |   区分    |   時間" & vbCrLf & _
              "---------------------------------------------" & vbCrLf
    For Each k In agg.Keys
        Dim sp() As String
        sp = Split(CStr(k), "|") ' 0:区分 1:作番
        preview = preview & sp(1) & vbTab & " | " & sp(0) & vbTab & " | " & MinutesToHHMMString(agg(k)) & vbCrLf
    Next
    If MsgBox(preview, vbYesNo + vbQuestion, "転記プレビュー") = vbNo Then GoTo Clean
    If DRY_RUN Then MsgBox "ドライランのため書き込みは実施しません。", vbInformation: GoTo Clean

    ' クリップボード（行明細をコピー）【#7 改善】
    Dim clip As String: clip = ""
    For i = DATA_START_ROW To lastDataRow
        If Not (IsEmpty(wsData.Cells(i, "C").Value) And IsEmpty(wsData.Cells(i, "D").Value) And IsEmpty(wsData.Cells(i, "E").Value)) Then
            clip = clip & CStr(wsData.Cells(i, "C").Value) & vbTab & _
                         CStr(wsData.Cells(i, "D").Value) & vbTab & _
                         wsData.Cells(i, "E").text & vbCrLf ' 見た目重視
        End If
    Next i
    If LenB(clip) > 0 Then CopyTextToClipboardAuto clip

    ' メッセージ列ヘッダ
    EnsureMessageHeader wsMonthly

    ' 書き込み（一括） + 重複上書き検知
    Dim dupCount As Long: dupCount = 0
For Each k In agg.Keys
    Dim col As Long
    Dim sp2() As String
    combinedKey = CStr(k)
    sp2 = Split(combinedKey, "|") ' 0:区分 1:作番
    If UBound(sp2) <> 1 Then sp2 = Array("", "")

    If dicMap.Exists(combinedKey) Then
        col = dicMap(combinedKey)
    Else
        Select Case AUTO_ADD_POLICY
            Case AddPolicy_Reject
                GoTo NextKey
            Case AddPolicy_Auto
                col = EnsureCategoryWorkNoColumn(sp2(0), sp2(1), wsMonthly, dicMap, False)
            Case AddPolicy_Prompt
                col = EnsureCategoryWorkNoColumn(sp2(0), sp2(1), wsMonthly, dicMap, True)
        End Select
        If col = 0 Then GoTo NextKey
    End If

    ' 既存値の検知
    Dim existingSerial As Double
    existingSerial = NzD(wsMonthly.Cells(targetRow, col).Value, 0#)
    If existingSerial <> 0# Then
        dupCount = dupCount + 1
        With wsMonthly.Cells(targetRow, col).Interior
            .Pattern = xlSolid
            .Color = DUP_HIGHLIGHT_COLOR
        End With
        AppendMessage wsMonthly, targetRow, "既存値あり: [" & sp2(1) & "|" & sp2(0) & "] 旧=" & _
                       SerialToHHMMString(existingSerial) & _
                       IIf(ACCUMULATE_MODE, " → 加算", " → 上書")
    End If

    ' 書き込み（加算/上書き）
    Dim addSerial As Double: addSerial = MinutesToSerial(agg(combinedKey))
    With wsMonthly.Cells(targetRow, col)
        If ACCUMULATE_MODE Then
            .Value = existingSerial + addSerial
        Else
            .Value = addSerial
        End If
        .NumberFormatLocal = "[h]:mm"
    End With

NextKey:
Next

MsgBox agg.Count & " 件のキーを処理しました。" & IIf(dupCount > 0, vbCrLf & "※ 重複検知: " & dupCount & " 件（黄色ハイライト＆メモ出力）", ""), vbInformation, "処理完了"

Clean:
    ' 再保護
    If wasProt Then ReprotectIfNeeded wsMonthly, usedPw
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvt
    Application.ScreenUpdating = prevUpd
    Exit Sub

EH:
    Dim emsg As String
    emsg = FriendlyErrorMessage(Err.Number, Err.Description)
    MsgBox emsg, vbCritical, "転記中のエラー"
    Resume Clean
End Sub

'=========================
' データ消去（参考: 既存機能維持）
'=========================
Public Sub ClearInputData()
    Dim wsAcq As Worksheet, wsData As Worksheet
    Dim wasProtA As Boolean, wasProtD As Boolean, usedPwA As String, usedPwD As String

    On Error GoTo EH

    If MsgBox("『" & ACQUISITION_SHEET_NAME & "』『" & DATA_SHEET_NAME & "』の入力をクリアします。よろしいですか？", _
              vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    Set wsAcq = ThisWorkbook.Sheets(ACQUISITION_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    wasProtA = wsAcq.ProtectContents: If wasProtA Then If Not UnprotectSheetEx(wsAcq, usedPwA) Then GoTo Clean
    wasProtD = wsData.ProtectContents: If wasProtD Then If Not UnprotectSheetEx(wsData, usedPwD) Then GoTo Clean

    wsAcq.Range("C8:E22").ClearContents
    wsData.Range("D4").ClearContents
    wsData.Range("F8:F17").ClearContents

    MsgBox "クリア完了", vbInformation

Clean:
    If wasProtA Then ReprotectIfNeeded wsAcq, usedPwA
    If wasProtD Then ReprotectIfNeeded wsData, usedPwD
    Exit Sub

EH:
    MsgBox "クリア処理でエラー: " & Err.Description, vbCritical
    Resume Clean
End Sub

'=========================
' ヘルパー群
'=========================
Private Function FriendlyErrorMessage(ByVal errNum As Long, ByVal errDesc As String) As String
    Select Case errNum
        Case 9 ' Subscript out of range
            FriendlyErrorMessage = _
                "エラー #9（インデックスが有効範囲にありません）" & vbCrLf & _
                "考えられる原因と対処:" & vbCrLf & _
                "・シート名の誤り：『" & DATA_SHEET_NAME & "』『" & MONTHLY_SHEET_NAME & "』が存在するか" & vbCrLf & _
                "・キー解析エラー：区分|作番の形式が不正（区切り '|' が無い等）" & vbCrLf & _
                "・列番号未確定：列追加を拒否して 'col=0' のまま（今回の処理では防御済み）" & vbCrLf & _
                "・配列の未初期化：Split結果が空のまま参照（今回の処理では防御済み）" & vbCrLf & _
                vbCrLf & "詳細: " & errDesc
        Case Else
            FriendlyErrorMessage = "エラー #" & errNum & vbCrLf & errDesc
    End Select
End Function

'=========================
' ヘルパー群
'=========================
Private Function ConvertToMinutesEx(ByVal timeValue As Variant) As Double
    Dim s As String
    ConvertToMinutesEx = 0
    If IsEmpty(timeValue) Then Exit Function

    If IsDate(timeValue) Then
        ConvertToMinutesEx = CDbl(CDate(timeValue)) * 1440#
        Exit Function
    End If

    If IsNumeric(timeValue) Then
        If InStr(1, CStr(timeValue), ".") > 0 Then
            ConvertToMinutesEx = CDbl(timeValue) * 1440#
        Else
            Dim hhmmStr As String, h As Long, m As Long
            hhmmStr = CStr(CLng(timeValue))
            Select Case Len(hhmmStr)
                Case 1, 2: h = 0: m = CLng(hhmmStr)
                Case 3, 4: h = CLng(Left$(hhmmStr, Len(hhmmStr) - 2)): m = CLng(Right$(hhmmStr, 2))
                Case Else: Exit Function
            End Select
            If m < 60 Then ConvertToMinutesEx = (h * 60#) + m
        End If
        Exit Function
    End If

    s = Trim$(CStr(timeValue))
    If InStr(s, ":") > 0 Then
        Dim parts() As String
        parts = Split(s, ":")
        If UBound(parts) = 1 Then
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                Dim h2 As Long, m2 As Long
                h2 = CLng(parts(0)): m2 = CLng(parts(1))
                If m2 >= 0 And m2 < 60 Then ConvertToMinutesEx = (h2 * 60#) + m2
            End If
        End If
    End If
End Function

Private Function MinutesToHHMMString(ByVal totalMinutes As Double) As String
    Dim h As Long, m As Long
    If totalMinutes <= 0 Then MinutesToHHMMString = "0000": Exit Function
    h = Int(totalMinutes / 60#)
    m = Round(totalMinutes - (h * 60#), 0)
    If m = 60 Then h = h + 1: m = 0
    MinutesToHHMMString = Format$(h, "00") & Format$(m, "00")
End Function

Private Function SerialToHHMMString(ByVal serial As Double) As String
    SerialToHHMMString = MinutesToHHMMString(serial * 1440#)
End Function

Private Function MinutesToSerial(ByVal m As Double) As Double
    MinutesToSerial = m / 1440#
End Function

Private Function FindMatchingDateRowDyn(ws As Worksheet, targetDate As Date) As Long
    Dim lastRow As Long, i As Long, d As Date
    lastRow = ws.Cells(ws.rows.Count, DATE_COL).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then Exit Function
    For i = MONTHLY_DATA_START_ROW To lastRow
        If IsDate(ws.Cells(i, DATE_COL).Value) Then
            d = CDate(ws.Cells(i, DATE_COL).Value)
            If Int(d) = Int(targetDate) Then FindMatchingDateRowDyn = i: Exit Function
        End If
    Next i
End Function

Private Sub EnsureDataFormat(ws As Worksheet, ByVal col As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, DATE_COL).End(xlUp).Row
    If lastRow < MONTHLY_DATA_START_ROW Then lastRow = MONTHLY_DATA_START_ROW + 31
    With ws.Range(ws.Cells(MONTHLY_DATA_START_ROW, col), ws.Cells(lastRow, col))
        .NumberFormatLocal = "[h]:mm"
    End With
End Sub

Private Sub EnsureMessageHeader(ws As Worksheet)
    With ws.Cells(MONTHLY_HEADER_ROW, MESSAGE_COL)
        If Trim$(CStr(.Value)) = "" Then .Value = "メッセージ"
        .Font.Bold = True
    End With
End Sub

Private Sub AppendMessage(ws As Worksheet, ByVal rowNum As Long, ByVal note As String)
    With ws.Cells(rowNum, MESSAGE_COL)
        If LenB(.Value) = 0 Then
            .Value = note
        Else
            .Value = CStr(.Value) & vbLf & note
        End If
    End With
End Sub

' 列追加（ポリシーにより自動/都度確認）
Private Function EnsureCategoryWorkNoColumn(ByVal category As String, _
                                           ByVal workNo As String, _
                                           ByRef wsMonthly As Worksheet, _
                                           ByRef dicMap As Object, _
                                           Optional ByVal promptUser As Boolean = True) As Long
    Dim combinedKey As String
    combinedKey = category & "|" & workNo

    If dicMap.Exists(combinedKey) Then
        EnsureCategoryWorkNoColumn = dicMap(combinedKey)
        Exit Function
    End If

    If promptUser Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("区分『" & category & "』＋作番『" & workNo & "』の列がありません。" & vbCrLf & _
                      "月次データに列を追加しますか？", vbYesNo + vbQuestion, "列の追加")
        If resp = vbNo Then EnsureCategoryWorkNoColumn = 0: Exit Function
    End If

    ' 末尾へ追加
    Dim lastCol As Long, newCol As Long, prevCol As Long
    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    prevCol = IIf(newCol > 1, newCol - 1, newCol)

    wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol).Value = category
    wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol).Value = workNo

    On Error Resume Next
    wsMonthly.Columns(newCol).ColumnWidth = wsMonthly.Columns(prevCol).ColumnWidth
    With wsMonthly.Cells(MONTHLY_HEADER_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_HEADER_ROW, prevCol).WrapText
    End With
    With wsMonthly.Cells(MONTHLY_WORKNO_ROW, newCol)
        .HorizontalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).HorizontalAlignment
        .VerticalAlignment = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).VerticalAlignment
        .Interior.Color = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).Interior.Color
        .Font.Bold = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).Font.Bold
        .WrapText = wsMonthly.Cells(MONTHLY_WORKNO_ROW, prevCol).WrapText
    End With
    On Error GoTo 0

    EnsureDataFormat wsMonthly, newCol ' データ行の書式を適用

    dicMap.Add combinedKey, newCol
    EnsureCategoryWorkNoColumn = newCol
End Function

'=== クリップボード: DataObject優先＋WinAPIフォールバック（#7） ===
Private Sub CopyTextToClipboardAuto(ByVal sText As String)
    On Error GoTo API_Fallback
    Dim dobj As Object
    Set dobj = CreateObject("Forms.DataObject") ' 参照設定不要（Late Binding）
    dobj.SetText sText
    dobj.PutInClipboard
    Exit Sub
API_Fallback:
    CopyTextToClipboard_WinAPI sText
End Sub

Private Sub CopyTextToClipboard_WinAPI(ByVal sText As String)
#If VBA7 Then
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
#Else
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
#End If
    Dim cb As Long, i As Long
    If LenB(sText) = 0 Then Exit Sub

    cb = (Len(sText) + 1) * 2
    hGlobalMemory = GlobalAlloc(GMEM_MOVEABLE, cb)
    If hGlobalMemory = 0 Then Exit Sub

    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory <> 0 Then
        lstrcpyW lpGlobalMemory, StrPtr(sText)
        Call GlobalUnlock(hGlobalMemory)

        For i = 1 To 5
            If OpenClipboard(0) <> 0 Then Exit For
            DoEvents
        Next i
        If i <= 5 Then
            Call EmptyClipboard
            If SetClipboardData(CF_UNICODETEXT, hGlobalMemory) = 0 Then
                ' 失敗時はリーク防止
                GlobalFree hGlobalMemory
            End If
            Call CloseClipboard
        Else
            ' OpenClipboard が取れなかった場合も解放
            GlobalFree hGlobalMemory
        End If
    Else
        GlobalFree hGlobalMemory
    End If
End Sub

'=== 保護解除/再保護（UserInterfaceOnly 指定） ===
Private Function UnprotectSheetEx(ws As Worksheet, Optional ByRef usedPw As String = "") As Boolean
    On Error Resume Next
    ws.Unprotect ""
    If Err.Number = 0 Then UnprotectSheetEx = True: Exit Function
    Err.Clear
    usedPw = InputBox("シート『" & ws.Name & "』のパスワードを入力してください。", "保護の解除")
    If usedPw = "" Then UnprotectSheetEx = False: Exit Function
    ws.Unprotect usedPw
    If Err.Number = 0 Then UnprotectSheetEx = True Else UnprotectSheetEx = False
    On Error GoTo 0
End Function

Private Sub ReprotectIfNeeded(ws As Worksheet, Optional ByVal usedPw As String = "")
    On Error Resume Next
    If usedPw = "" Then
        ws.Protect UserInterfaceOnly:=True
    Else
        ws.Protect password:=usedPw, UserInterfaceOnly:=True
    End If
    On Error GoTo 0
End Sub

'======== NzD (Null/Empty/Error を既定値に置き換える) ========
Private Function NzD(ByVal v As Variant, Optional ByVal def As Double = 0#) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then
        NzD = def
    Else
        NzD = v
    End If
End Function



