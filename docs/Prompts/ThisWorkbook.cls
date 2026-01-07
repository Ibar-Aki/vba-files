VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ExitHandler

    If Target.Column <> 4 Then Exit Sub

    Dim clipValue As String
    clipValue = CStr(Target.Offset(0, 1).Value)

    Cancel = True

    If Len(clipValue) = 0 Then
        Application.StatusBar = "E?????????"
        Application.OnTime Now + TimeSerial(0, 0, 3), "ClearStatusBar"
        Exit Sub
    End If

    Dim dataObj As Object
    Set dataObj = CreateObject("MSForms.DataObject")
    dataObj.SetText clipValue
    dataObj.PutInClipboard

    Application.StatusBar = "E???????????????????: " & clipValue
    Application.OnTime Now + TimeSerial(0, 0, 3), "ClearStatusBar"

ExitHandler:
    Set dataObj = Nothing
End Sub
