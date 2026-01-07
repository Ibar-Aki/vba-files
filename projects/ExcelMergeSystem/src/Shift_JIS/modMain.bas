'==================================================
' modMain - CW[
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' W[xϐiO[oϐ̕ύXj
'--------------------------------------------------
Private m_ConfigData As Object
Private m_LogCollection As Collection
Private m_IsProcessing As Boolean

'--------------------------------------------------
' ݒf[^ւ̃ANZTiJvZj
'--------------------------------------------------
Public Property Get ConfigData() As Object
    Set ConfigData = m_ConfigData
End Property

Public Property Set ConfigData(ByVal value As Object)
    Set m_ConfigData = value
End Property

Public Property Get LogCollection() As Collection
    Set LogCollection = m_LogCollection
End Property

Public Property Set LogCollection(ByVal value As Collection)
    Set m_LogCollection = value
End Property

'--------------------------------------------------
' CGg[|CgiOĂяoj
'--------------------------------------------------
Public Sub ExecuteMerge(ByVal strFile1 As String, ByVal strFile2 As String)
    
    Dim startTime As Date
    Dim result As Boolean
    Dim appState As Object
    
    ' dsh~
    If m_IsProcessing Then
        MsgBox "ɎsłB", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    m_IsProcessing = True
    result = False
    
    ' 
    startTime = Now
    Set appState = SaveApplicationState()
    
    ' AvP[VݒiptH[}Xj
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' O
    Call InitializeLog
    Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Call LogMessage(APP_TITLE & " v" & APP_VERSION & " Jn", LOG_LEVEL_INFO)
    Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Call LogMessage("Excel1: " & GetFileName(strFile1), LOG_LEVEL_INFO)
    Call LogMessage("Excel2: " & GetFileName(strFile2), LOG_LEVEL_INFO)
    
    ' ݒǍ
    If Not LoadConfiguration() Then
        Call LogMessage("ݒt@C̓ǂݍ݂Ɏs܂", LOG_LEVEL_ERROR)
        GoTo Cleanup
    End If
    
    ' ݒl
    If Not ValidateConfigValues() Then
        Call LogMessage("ݒľ؂Ɏs܂", LOG_LEVEL_ERROR)
        GoTo Cleanup
    End If
    
    ' t@C
    If Not ValidateFiles(strFile1, strFile2) Then
        Call LogMessage("t@C؃G[", LOG_LEVEL_ERROR)
        GoTo Cleanup
    End If
    
    ' f[^s
    result = ProcessMerge(strFile1, strFile2)
    
    If result Then
        Call LogMessage(" : " & _
            Format(Now - startTime, TIMESTAMP_FORMAT_TIME), LOG_LEVEL_INFO)
        Call LogMessage("=" & String(40, "="), LOG_LEVEL_INFO)
    Else
        Call LogMessage("s", LOG_LEVEL_ERROR)
    End If
    
Cleanup:
    ' AvP[Vԕ
    Call RestoreApplicationState(appState)
    
    ' 
    Call CleanupResources
    
    ' bZ[W
    If result Then
        MsgBox "܂B" & vbCrLf & _
               "o̓tH_mFĂB", _
               vbInformation, APP_TITLE
    Else
        MsgBox "ɃG[܂B" & vbCrLf & _
               "OmFĂB", _
               vbExclamation, APP_TITLE
    End If
    
    m_IsProcessing = False
    
    ' SȎȏIĩubNɉe^Ȃj
    Call SafeCloseThisWorkbook
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage("VXeG[: " & Err.Description & _
                   " (G[ԍ: " & Err.Number & ")", LOG_LEVEL_ERROR)
    result = False
    Resume Cleanup
    
End Sub

'--------------------------------------------------
' C
'--------------------------------------------------
Private Function ProcessMerge(ByVal file1 As String, ByVal file2 As String) As Boolean
    
    Dim data1 As Object
    Dim data2 As Object
    Dim mergedData As Object
    Dim outputPath As String
    
    On Error GoTo ErrorHandler
    
    ProcessMerge = False
    Set data1 = Nothing
    Set data2 = Nothing
    Set mergedData = Nothing
    
    ' Excel1Ǎ
    Call LogMessage("Excel1ǍJn...", LOG_LEVEL_INFO)
    Set data1 = LoadExcelData(file1, "Excel1")
    If data1 Is Nothing Then
        Call LogMessage("Excel1̃f[^ǍɎs܂", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' Excel2Ǎ
    Call LogMessage("Excel2ǍJn...", LOG_LEVEL_INFO)
    Set data2 = LoadExcelData(file2, "Excel2")
    If data2 Is Nothing Then
        Call LogMessage("Excel2̃f[^ǍɎs܂", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' f[^
    Call LogMessage("f[^Jn...", LOG_LEVEL_INFO)
    Set mergedData = MergeData(data1, data2)
    If mergedData Is Nothing Then
        Call LogMessage("f[^Ɏs܂", LOG_LEVEL_ERROR)
        GoTo CleanupLocal
    End If
    
    ' ʂǉƂĕۑ
    mergedData("File1Name") = GetFileName(file1)
    mergedData("File2Name") = GetFileName(file2)
    
    ' o
    Call LogMessage("t@Co͊Jn...", LOG_LEVEL_INFO)
    outputPath = GenerateOutput(mergedData)
    
    ProcessMerge = (outputPath <> "")
    
    If ProcessMerge Then
        Call LogMessage("o̓t@C: " & outputPath, LOG_LEVEL_INFO)
    End If
    
CleanupLocal:
    ' [JIuWFNg̃
    Set data1 = Nothing
    Set data2 = Nothing
    Set mergedData = Nothing
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ProcessMerge Error: " & Err.Description, LOG_LEVEL_ERROR)
    ProcessMerge = False
    Resume CleanupLocal
    
End Function

'--------------------------------------------------
' AvP[VԂ̕ۑ
'--------------------------------------------------
Private Function SaveApplicationState() As Object
    Dim state As Object
    Set state = CreateObject("Scripting.Dictionary")
    
    With Application
        state("ScreenUpdating") = .ScreenUpdating
        state("DisplayAlerts") = .DisplayAlerts
        state("Calculation") = .Calculation
        state("EnableEvents") = .EnableEvents
    End With
    
    Set SaveApplicationState = state
End Function

'--------------------------------------------------
' AvP[VԂ̕
'--------------------------------------------------
Private Sub RestoreApplicationState(ByVal state As Object)
    If state Is Nothing Then
        ' ftHgԂɕ
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End With
    Else
        With Application
            .ScreenUpdating = state("ScreenUpdating")
            .DisplayAlerts = state("DisplayAlerts")
            .Calculation = state("Calculation")
            .EnableEvents = state("EnableEvents")
        End With
    End If
End Sub

'--------------------------------------------------
' \[XN[Abv
'--------------------------------------------------
Private Sub CleanupResources()
    On Error Resume Next
    
    ' W[xϐ̃NA
    Set m_ConfigData = Nothing
    Set m_LogCollection = Nothing
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' O[o\[X̃N[AbviThisWorkbookĂяoj
'--------------------------------------------------
Public Sub CleanupGlobalResources()
    Call CleanupResources
End Sub

'--------------------------------------------------
' SȎȏI
'--------------------------------------------------
Private Sub SafeCloseThisWorkbook()
    On Error Resume Next
    
    Dim wb As Workbook
    
    ' ̃ubNJĂ邩mF
    If Workbooks.Count > 1 Then
        ' ̃ubNꍇ͎
        ThisWorkbook.Close SaveChanges:=False
    Else
        ' ̏ꍇ͉ȂiExcelIĂ܂߁j
        ' [U[蓮ŕ
    End If
    
    On Error GoTo 0
End Sub

'--------------------------------------------------
' t@C擾
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
