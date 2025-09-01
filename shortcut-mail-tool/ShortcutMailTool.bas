Attribute VB_Name = "ShortcutMailTool"
Option Explicit

' メニューを表示して選択した項目をクリップボードにコピーする
Public Sub ShowShortcutMailMenu()
    Dim items As Variant
    items = LoadMailItems(ThisWorkbook.Path & "\sample_data.csv")
    If IsEmpty(items) Then
        MsgBox "データが読み込めませんでした", vbExclamation
        Exit Sub
    End If
    Dim i As Long, menuText As String
    For i = LBound(items) To UBound(items)
        menuText = menuText & i & ": " & items(i)(0) & vbCrLf
    Next i
    Dim sel As Variant
    sel = Application.InputBox("コピーする項目番号を入力してください:" & vbCrLf & menuText, "ショートカットメール", Type:=1)
    If sel = False Then Exit Sub
    If sel < LBound(items) Or sel > UBound(items) Then
        MsgBox "無効な番号です", vbExclamation
        Exit Sub
    End If
    CopyToClipboard items(sel)(1)
    MsgBox "コピーしました: " & items(sel)(0)
End Sub

' CSV ファイルを読み込み項目を配列で返す
Private Function LoadMailItems(csvPath As String) As Variant
    Dim f As Integer, text As String, lines As Variant
    f = FreeFile
    On Error GoTo ErrHandler
    Open csvPath For Input As #f
    text = Input$(LOF(f), f)
    Close #f
    text = Replace(text, vbCrLf, vbLf)
    lines = Split(text, vbLf)
    If UBound(lines) <= 0 Then
        LoadMailItems = Empty
        Exit Function
    End If
    Dim data() As Variant, i As Long, fields As Variant, count As Long
    ReDim data(1 To UBound(lines))
    For i = 2 To UBound(lines) + 1 ' 1 行目はヘッダー
        If i - 1 <= UBound(lines) Then
            If Trim(lines(i - 1)) <> "" Then
                fields = Split(lines(i - 1), ",")
                count = count + 1
                data(count) = Array(fields(0), fields(1))
            End If
        End If
    Next i
    If count = 0 Then
        LoadMailItems = Empty
        Exit Function
    End If
    ReDim Preserve data(1 To count)
    LoadMailItems = data
    Exit Function
ErrHandler:
    LoadMailItems = Empty
    MsgBox "データの読み込み中にエラーが発生しました: " & Err.Description, vbExclamation
    On Error Resume Next
    If f > 0 Then Close #f
End Function

' 文字列をクリップボードへコピー
Private Sub CopyToClipboard(ByVal txt As String)
    Dim obj As Object
    Set obj = CreateObject("MSForms.DataObject")
    obj.SetText txt
    obj.PutInClipboard
End Sub
