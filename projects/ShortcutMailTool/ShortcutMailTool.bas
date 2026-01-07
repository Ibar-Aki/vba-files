' Fix: stabilize clipboard handling and remove success dialog
Attribute VB_Name = "ShortcutMailTool"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = GMEM_MOVEABLE Or GMEM_ZEROINIT

' メニューを表示して選択した項目をクリップボードにコピーする
Public Sub ShowShortcutMailMenu()
    Dim items As Variant
    items = LoadMailItems(ThisWorkbook.Path & "\sample_data.csv")
    If IsEmpty(items) Then
        MsgBox "データが読み込めませんでした。", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    Dim menuText As String
    For i = LBound(items) To UBound(items)
        menuText = menuText & i & ": " & items(i)(0) & vbCrLf
    Next i

    Dim selectionValue As Variant
    selectionValue = Application.InputBox("コピーする項目番号を入力してください:" & vbCrLf & menuText, "ショートカットメール", Type:=1)
    If VarType(selectionValue) = vbBoolean Then Exit Sub
    If Not IsNumeric(selectionValue) Then
        MsgBox "番号は整数で入力してください。", vbExclamation
        Exit Sub
    End If

    Dim selectedIndex As Long
    selectedIndex = CLng(selectionValue)
    If selectionValue <> selectedIndex Then
        MsgBox "番号は整数で入力してください。", vbExclamation
        Exit Sub
    End If

    If selectedIndex < LBound(items) Or selectedIndex > UBound(items) Then
        MsgBox "無効な番号です。", vbExclamation
        Exit Sub
    End If

    If Not CopyTextToClipboard(items(selectedIndex)(1)) Then
        MsgBox "クリップボードへのコピーに失敗しました。", vbExclamation
        Exit Sub
    End If

End Sub

' CSV ファイルを読み込み項目を配列で返す
Private Function LoadMailItems(csvPath As String) As Variant
    Dim rawText As String
    rawText = ReadCsvText(csvPath)
    If Len(rawText) = 0 Then
        LoadMailItems = Empty
        Exit Function
    End If

    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)

    Dim lines As Variant
    lines = Split(rawText, vbLf)
    If UBound(lines) < 1 Then
        LoadMailItems = Empty
        Exit Function
    End If

    Dim data() As Variant
    ReDim data(1 To UBound(lines))

    Dim count As Long
    Dim i As Long
    For i = 1 To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))
        If Len(line) > 0 Then
            Dim commaPos As Long
            commaPos = InStr(1, line, ",")
            If commaPos > 0 Then
                Dim label As String
                Dim content As String
                label = Trim$(Left$(line, commaPos - 1))
                content = Trim$(Mid$(line, commaPos + 1))
                count = count + 1
                data(count) = Array(label, content)
            End If
        End If
    Next i

    If count = 0 Then
        LoadMailItems = Empty
        Exit Function
    End If

    ReDim Preserve data(1 To count)
    LoadMailItems = data
End Function

' CSV ファイルを文字列として読み込む
Private Function ReadCsvText(csvPath As String) As String
    If Len(Dir$(csvPath)) = 0 Then
        MsgBox "CSV ファイルが見つかりません: " & csvPath, vbExclamation
        Exit Function
    End If

    Dim fileNum As Integer
    Dim fileLength As Long
    fileNum = FreeFile

    On Error GoTo ErrHandler
    Open csvPath For Binary As #fileNum
    fileLength = LOF(fileNum)

    If fileLength = 0 Then
        Close #fileNum
        Exit Function
    End If

    Dim buffer() As Byte
    ReDim buffer(0 To fileLength - 1) As Byte
    Get #fileNum, , buffer
    Close #fileNum

    ReadCsvText = DecodeCsvBuffer(buffer)
    Exit Function

ErrHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "CSV の読み込み中にエラーが発生しました: " & Err.Description, vbExclamation
End Function

' バイト配列を文字列にデコードする
Private Function DecodeCsvBuffer(ByRef buffer() As Byte) As String
    Dim text As String
    text = ReadWithCharset(buffer, "utf-8")
    If Len(text) = 0 Then
        text = ReadWithCharset(buffer, "shift_jis")
    End If
    If Len(text) = 0 Then
        text = ReadWithCharset(buffer, "Windows-31J")
    End If
    If Len(text) = 0 Then
        text = StrConv(buffer, vbUnicode)
    End If

    If Len(text) > 0 Then
        text = Replace(text, ChrW$(&HFEFF), "")
    End If
    DecodeCsvBuffer = text
End Function

' 指定した文字コードで読み込む
Private Function ReadWithCharset(ByRef buffer() As Byte, ByVal charset As String) As String
    On Error GoTo ErrHandler
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 1
        .Open
        .Write buffer
        .Position = 0
        .Type = 2
        .Charset = charset
        ReadWithCharset = .ReadText
        .Close
    End With

    Set stream = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    ReadWithCharset = ""
End Function

' 文字列をクリップボードへコピー（成功時 True）
Private Function CopyTextToClipboard(ByVal txt As String) As Boolean
    Dim result As Boolean
    Dim clipboardText As String
#If VBA7 Then
    Dim byteSize As LongPtr
    Dim hMem As LongPtr
    Dim pMem As LongPtr
#Else
    Dim byteSize As Long
    Dim hMem As Long
    Dim pMem As Long
#End If
    Dim locked As Boolean

    clipboardText = txt & vbNullChar
    byteSize = LenB(clipboardText)
    If byteSize = 0 Then Exit Function

#If VBA7 Then
    Dim hwndTarget As LongPtr
    hwndTarget = Application.hwnd
#Else
    Dim hwndTarget As Long
    hwndTarget = Application.hwnd
#End If

    If OpenClipboard(hwndTarget) = 0 Then Exit Function

    On Error GoTo CleanFail

    EmptyClipboard

    hMem = GlobalAlloc(GHND, byteSize)
    If hMem = 0 Then GoTo CleanFail

    pMem = GlobalLock(hMem)
    If pMem = 0 Then GoTo CleanFail
    locked = True

    CopyMemory ByVal pMem, ByVal StrPtr(clipboardText), byteSize

    GlobalUnlock hMem
    locked = False

    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then GoTo CleanFail

    result = True
    hMem = 0 ' 所有権はシステムへ移動

CleanExit:
    If locked Then
        GlobalUnlock hMem
    End If
    If hMem <> 0 Then
        GlobalFree hMem
    End If
    CloseClipboard
    CopyTextToClipboard = result
    Exit Function

CleanFail:
    result = False
    GoTo CleanExit
End Function

