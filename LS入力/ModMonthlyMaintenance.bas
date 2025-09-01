Option Explicit
'
'===============================================================================
' 繝｢繧ｸ繝･繝ｼ繝ｫ蜷・ ModMonthlyMaintenance
'
' 讎りｦ・
'   譛域ｬ｡繧ｷ繝ｼ繝医・霆｢險倥ョ繝ｼ繧ｿ繧剃ｸ諡ｬ繧ｯ繝ｪ繧｢・亥､縺ｨ蝪励ｊ縺､縺ｶ縺励ｒ隗｣髯､・峨＠縲・'   蟇ｾ雎｡譛医・繧ｫ繝ｬ繝ｳ繝繝ｼ・・蛻励・譌･莉假ｼ峨ｒ譖ｴ譁ｰ縺励∪縺吶・'   繧ｫ繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ縺ｯ Yes/No 縺ｮ遒ｺ隱阪ム繧､繧｢繝ｭ繧ｰ縺ｧ螳溯｡悟庄蜷ｦ繧堤｢ｺ隱阪＠縺ｾ縺吶・'
' 蟇ｾ雎｡迺ｰ蠅・ Excel 2016+ / Windows
'===============================================================================

' --- 繧ｷ繝ｼ繝亥錐繝ｻ菴咲ｽｮ縺ｪ縺ｩ・域悽繝｢繧ｸ繝･繝ｼ繝ｫ蜀・〒菴ｿ逕ｨ縺吶ｋ螳壽焚・・---
Private Const DATA_SHEET_NAME    As String = "繝・・繧ｿ逋ｻ骭ｲ"   ' 蟇ｾ雎｡譌･繧貞叙蠕励☆繧九す繝ｼ繝・Private Const MONTHLY_SHEET_NAME As String = "譛域ｬ｡繝・・繧ｿ"   ' 繧ｯ繝ｪ繧｢繝ｻ譖ｴ譁ｰ縺ｮ蟇ｾ雎｡繧ｷ繝ｼ繝・
Private Const COL_DATE               As Long = 2   ' B蛻・ 譌･莉伜・
Private Const MONTHLY_MIN_COL        As Long = 3   ' C蛻嶺ｻ･髯阪′菴懈･ｭ蛻・Private Const MONTHLY_HEADER_ROW     As Long = 11  ' 隕句・縺暦ｼ井ｽ懈･ｭ繧ｳ繝ｼ繝峨↑縺ｩ・峨・陦・Private Const MONTHLY_DATA_START_ROW As Long = 12  ' 繝・・繧ｿ髢句ｧ玖｡・ ' 繧ｨ繝ｩ繝ｼ陦ｨ遉ｺ繧ｻ繝ｫ縺ｯ蜈ｱ騾壼ｮ壽焚 ERR_CELL_ADDR 繧剃ｽｿ逕ｨ・・odAppConfig.bas・・
'===============================================================================
' 讖溯・蜷・ 譛域ｬ｡繝・・繧ｿ縺ｮ蜈ｨ繧ｯ繝ｪ繧｢・九き繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ
' 讎りｦ・ : 譛域ｬ｡繧ｷ繝ｼ繝医・霆｢險倥ョ繝ｼ繧ｿ・亥､/譎る俣・峨ｒ蜈ｨ豸亥悉縺励∝｡励ｊ縺､縺ｶ縺励ｂ隗｣髯､縺励∪縺吶・'         縺昴・蠕後√後ョ繝ｼ繧ｿ逋ｻ骭ｲ縲阪す繝ｼ繝医・蟇ｾ雎｡譌･・・4蜆ｪ蜈遺・D3・峨→蜷後§譛医〒縲・'         B蛻励↓繧ｫ繝ｬ繝ｳ繝繝ｼ・域律莉假ｼ峨ｒ蜀堺ｽ懈・縺励∪縺吶ょｮ溯｡悟燕縺ｫ遒ｺ隱阪ム繧､繧｢繝ｭ繧ｰ繧定｡ｨ遉ｺ縺励∪縺吶・'===============================================================================
Public Sub ClearMonthlyDataAndRefreshCalendar(Optional ByVal AskConfirm As Boolean = True)
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation

    Dim wsMonthly As Worksheet
    Dim wsData As Worksheet
    Dim wasProtected As Boolean
    Dim pwd As String
    Dim targetDate As Date
    Dim ret As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' --- 繧｢繝励Μ迥ｶ諷九・騾驕ｿ縺ｨ雋闕ｷ霆ｽ貂・---
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' --- 繧ｷ繝ｼ繝亥叙蠕・---
    Set wsMonthly = ThisWorkbook.Sheets(MONTHLY_SHEET_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)

    ' --- 螳溯｡悟燕縺ｫ繧ｨ繝ｩ繝ｼ陦ｨ遉ｺ繧ｻ繝ｫ繧偵け繝ｪ繧｢・亥ｭ伜惠縺吶ｌ縺ｰ・・---
    On Error Resume Next
    wsMonthly.Range(ERR_CELL_ADDR).ClearContents
    wsMonthly.Range(ERR_CELL_ADDR).WrapText = True
    On Error GoTo ErrorHandler

    ' --- 菫晁ｭｷ縺ｮ荳譎りｧ｣髯､ ---
    wasProtected = wsMonthly.ProtectContents
    If wasProtected Then
        On Error Resume Next
        wsMonthly.Unprotect ""
        If Err.Number <> 0 Then
            Err.Clear
            pwd = InputBox("繧ｷ繝ｼ繝医′菫晁ｭｷ縺輔ｌ縺ｦ縺・∪縺吶ゅヱ繧ｹ繝ｯ繝ｼ繝峨ｒ蜈･蜉帙＠縺ｦ縺上□縺輔＞:", _
                           "繧ｷ繝ｼ繝井ｿ晁ｭｷ隗｣髯､")
            If Len(pwd) = 0 Then GoTo CleanUp
            wsMonthly.Unprotect Password:=pwd
        End If
        On Error GoTo ErrorHandler
    End If

    ' --- 霆｢險倥ョ繝ｼ繧ｿ鬆伜沺縺ｮ荳諡ｬ繧ｯ繝ｪ繧｢・亥､/蝪励ｊ縺､縺ｶ縺暦ｼ・---
    ' ・亥ｯｾ雎｡譛育｢ｺ螳壼ｾ後↓繧ｯ繝ｪ繧｢繧貞ｮ滓命・・
    ' --- 蟇ｾ雎｡譌･・・4蜆ｪ蜈遺・D3・牙叙蠕・---
    If Not DetermineTargetDateLocal(wsData, targetDate) Then
    ReportErrorToMonthlySheetLocal wsMonthly, _
        "蟇ｾ雎｡譌･莉倥′蜿門ｾ励〒縺阪∪縺帙ｓ・・4 縺ｾ縺溘・ D3 繧定ｨｭ螳壹＠縺ｦ縺上□縺輔＞・・, True
        GoTo CleanUp
    End If

    ' --- 霆｢險倥ョ繝ｼ繧ｿ鬆伜沺縺ｮ荳諡ｬ繧ｯ繝ｪ繧｢・亥､/蝪励ｊ縺､縺ｶ縺暦ｼ・--
    '     蟇ｾ雎｡譛医・譛ｫ譌･陦後∪縺ｧ繧貞ｯｾ雎｡縺ｨ縺励√◎繧御ｻ･髯搾ｼ亥粋險医・繝｡繝｢陦鯉ｼ峨・蜑企勁縺励↑縺・    Dim daysInMonth As Long, lastDayRow As Long
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))
    lastDayRow = MONTHLY_DATA_START_ROW + daysInMonth - 1
    ClearAllMonthlyTransferArea wsMonthly, lastDayRow

    ' --- 繧ｫ繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ縺ｮ遒ｺ隱・---
    If AskConfirm Then
    ret = MsgBox( _
        "蟇ｾ雎｡譛医・繧ｫ繝ｬ繝ｳ繝繝ｼ・域律莉伜・・峨ｒ譖ｴ譁ｰ縺励∪縺吶・ & vbCrLf & _
        "蟇ｾ雎｡譛・ " & Format$(targetDate, "m/dd(aaa)") & vbCrLf & vbCrLf & _
        "繧医ｍ縺励＞縺ｧ縺吶°・・, _
        vbYesNo + vbQuestion, "繧ｫ繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ縺ｮ遒ｺ隱・)
    If ret <> vbYes Then GoTo CleanUp
    End If

    ' --- 繧ｫ繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ ---
    RefreshMonthlyCalendar wsMonthly, targetDate

CleanUp:
    ' --- 菫晁ｭｷ縺ｮ蠕ｩ蜈・---
    If wasProtected Then
        On Error Resume Next
        If Len(pwd) > 0 Then
            wsMonthly.Protect Password:=pwd, UserInterfaceOnly:=True
        Else
            wsMonthly.Protect UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    End If

    ' --- 繧｢繝励Μ迥ｶ諷九・蠕ｩ蜈・---
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrorHandler:
    ' --- 邁｡譏薙お繝ｩ繝ｼ蝣ｱ蜻奇ｼ・3縺ｫ霑ｽ險假ｼ・---
    On Error Resume Next
    ReportErrorToMonthlySheetLocal wsMonthly, _
        "譛域ｬ｡繧ｯ繝ｪ繧｢/繧ｫ繝ｬ繝ｳ繝繝ｼ譖ｴ譁ｰ繧ｨ繝ｩ繝ｼ: " & Err.Description, True
    On Error GoTo 0
    Resume CleanUp
End Sub

'-------------------------------------------------------------------------------
' 讖溯・蜷・ 霆｢險倥ョ繝ｼ繧ｿ鬆伜沺縺ｮ蜈ｨ繧ｯ繝ｪ繧｢・亥､縺ｨ蝪励ｊ縺､縺ｶ縺暦ｼ・' 蠑墓焚  : wsMonthly・域怦谺｡繧ｷ繝ｼ繝茨ｼ・'-------------------------------------------------------------------------------
Private Sub ClearAllMonthlyTransferArea(ByRef wsMonthly As Worksheet, ByVal lastDayRow As Long)
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range

    ' 繧ｯ繝ｪ繧｢蟇ｾ雎｡縺ｮ譛邨り｡後・縲∝ｯｾ雎｡譛医・譛ｫ譌･陦後∪縺ｧ
    lastRow = lastDayRow

    lastCol = wsMonthly.Cells(MONTHLY_HEADER_ROW, wsMonthly.Columns.Count).End(xlToLeft).Column
    If lastCol < MONTHLY_MIN_COL Then lastCol = MONTHLY_MIN_COL

    ' 蛟､繧ｯ繝ｪ繧｢・句｡励ｊ縺､縺ｶ縺苓ｧ｣髯､
    On Error Resume Next
    Set rng = wsMonthly.Range(wsMonthly.Cells(MONTHLY_DATA_START_ROW, MONTHLY_MIN_COL), _
                              wsMonthly.Cells(lastRow, lastCol))
    If Not rng Is Nothing Then
        rng.ClearContents
        rng.Interior.Pattern = xlNone
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' 讖溯・蜷・ 繧ｫ繝ｬ繝ｳ繝繝ｼ・域律莉伜・B・峨・譖ｴ譁ｰ
' 讎りｦ・ : 蟇ｾ雎｡譛医・1譌･縲懈忰譌･繧達蛻励↓騾｣邯壹〒險ｭ螳壹＠縲∽ｽ吝臆陦後′縺ゅｌ縺ｰ繧ｯ繝ｪ繧｢
' 蠑墓焚  : wsMonthly・域怦谺｡繧ｷ繝ｼ繝茨ｼ峨》argetDate・亥ｯｾ雎｡譌･・・'-------------------------------------------------------------------------------
Private Sub RefreshMonthlyCalendar(ByRef wsMonthly As Worksheet, ByVal targetDate As Date)
    Dim firstDate As Date
    Dim daysInMonth As Long
    Dim r As Long, rowStart As Long
    Dim lastRow As Long

    rowStart = MONTHLY_DATA_START_ROW
    firstDate = DateSerial(Year(targetDate), Month(targetDate), 1)
    daysInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))

    ' 蠢・ｦ∬｡梧焚蛻・∵律莉倥ｒ險ｭ螳・    For r = 0 To daysInMonth - 1
        With wsMonthly.Cells(rowStart + r, COL_DATE)
            .Value = firstDate + r
            .NumberFormatLocal = "mm/dd(aaa)"
            .Interior.Pattern = xlNone
        End With
    Next r

    ' 譛ｫ譌･莉･髯搾ｼ亥粋險医・繝｡繝｢陦後↑縺ｩ・峨・菫晄戟縺吶ｋ縺溘ａ縲√％縺薙〒縺ｯ菴輔ｂ縺励↑縺・End Sub

'-------------------------------------------------------------------------------
' 讖溯・蜷・ 蟇ｾ雎｡譌･・・4蜆ｪ蜈遺・D3・峨・蜿門ｾ・' 蠑墓焚  : wsData・医ョ繝ｼ繧ｿ逋ｻ骭ｲ繧ｷ繝ｼ繝茨ｼ・' 謌ｻ繧雁､: 蜿門ｾ励〒縺阪◆蝣ｴ蜷・True
'-------------------------------------------------------------------------------
Private Function DetermineTargetDateLocal(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    DetermineTargetDateLocal = False
    If IsDate(wsData.Range("D4").Value) Then
        targetDate = CDate(wsData.Range("D4").Value)
        DetermineTargetDateLocal = True
    ElseIf IsDate(wsData.Range("D3").Value) Then
        targetDate = CDate(wsData.Range("D3").Value)
        DetermineTargetDateLocal = True
    End If
End Function

'-------------------------------------------------------------------------------
' 讖溯・蜷・ 繧ｨ繝ｩ繝ｼ繝｡繝・そ繝ｼ繧ｸ縺ｮ陦ｨ遉ｺ・・3・峨Ｂppend=True 縺ｧ霑ｽ險・' 蠑墓焚  : wsMonthly・域怦谺｡繧ｷ繝ｼ繝茨ｼ峨［essage・郁｡ｨ遉ｺ蜀・ｮｹ・峨∥ppend・郁ｿｽ險倥ヵ繝ｩ繧ｰ・・'-------------------------------------------------------------------------------
Private Sub ReportErrorToMonthlySheetLocal(ByRef wsMonthly As Worksheet, ByVal message As String, Optional ByVal append As Boolean = False)
    On Error Resume Next
    If wsMonthly Is Nothing Then Exit Sub
    With wsMonthly.Range(ERR_CELL_ADDR)
        If append And Len(.Value) > 0 Then
            .Value = CStr(.Value) & vbLf & message
        Else
            .Value = message
        End If
        .WrapText = True
    End With
    On Error GoTo 0
End Sub
