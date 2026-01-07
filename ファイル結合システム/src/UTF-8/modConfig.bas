'==================================================
' modConfig - 設定管理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' 設定ファイル読込
'--------------------------------------------------
Public Function LoadConfiguration() As Boolean
    
    Dim configPath As String
    Dim wbConfig As Workbook
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim configDict As Object
    
    On Error GoTo ErrorHandler
    
    LoadConfiguration = False
    Set wbConfig = Nothing
    
    ' 設定ファイルパス
    configPath = ThisWorkbook.Path & "\" & CONFIG_FOLDER_NAME & "\" & CONFIG_FILE_NAME
    
    ' 設定ファイル存在確認
    If Not FileExists(configPath) Then
        Call LogMessage("設定ファイルが見つかりません: " & configPath, LOG_LEVEL_WARNING)
        ' デフォルト設定使用
        Call SetDefaultConfig
        LoadConfiguration = True
        Exit Function
    End If
    
    ' 設定読込
    Set configDict = CreateObject("Scripting.Dictionary")
    Set wbConfig = Workbooks.Open(configPath, ReadOnly:=True, UpdateLinks:=0)
    
    ' シート存在確認
    On Error Resume Next
    Set wsConfig = wbConfig.Worksheets(CONFIG_SHEET_NAME)
    If wsConfig Is Nothing Then
        Call LogMessage("設定シート'" & CONFIG_SHEET_NAME & "'が見つかりません", LOG_LEVEL_WARNING)
        wbConfig.Close False
        Call SetDefaultConfig
        LoadConfiguration = True
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' ヘッダー行をスキップ
        If wsConfig.Cells(i, 1).Value <> "" Then
            configDict(CStr(wsConfig.Cells(i, 1).Value)) = _
                CStr(wsConfig.Cells(i, 2).Value)
        End If
    Next i
    
    wbConfig.Close False
    Set wbConfig = Nothing
    
    ' 設定をモジュールに保存
    Set ConfigData = configDict
    
    ' 必須項目チェック
    If Not ValidateConfig() Then
        Call LogMessage("必須設定項目が不足しているため、デフォルト値を使用します", LOG_LEVEL_WARNING)
        Call SetDefaultConfig
    End If
    
    Call LogMessage("設定ファイル読込完了", LOG_LEVEL_INFO)
    LoadConfiguration = True
    
    Exit Function
    
ErrorHandler:
    If Not wbConfig Is Nothing Then
        wbConfig.Close False
        Set wbConfig = Nothing
    End If
    Call LogMessage("設定ファイル読込エラー: " & Err.Description, LOG_LEVEL_ERROR)
    Call SetDefaultConfig
    LoadConfiguration = True
    
End Function

'--------------------------------------------------
' デフォルト設定
'--------------------------------------------------
Private Sub SetDefaultConfig()
    
    Dim configDict As Object
    Set configDict = CreateObject("Scripting.Dictionary")
    
    ' Excel1設定
    configDict(CFG_EXCEL1_HEADER_ROWS) = CStr(DEFAULT_EXCEL1_HEADER_ROWS)
    configDict(CFG_EXCEL1_DATA_START_ROW) = CStr(DEFAULT_EXCEL1_DATA_START_ROW)
    configDict(CFG_EXCEL1_ID_COLUMN) = DEFAULT_EXCEL1_ID_COLUMN
    
    ' Excel2設定
    configDict(CFG_EXCEL2_HEADER_ROWS) = CStr(DEFAULT_EXCEL2_HEADER_ROWS)
    configDict(CFG_EXCEL2_DATA_START_ROW) = CStr(DEFAULT_EXCEL2_DATA_START_ROW)
    configDict(CFG_EXCEL2_ID_COLUMN) = DEFAULT_EXCEL2_ID_COLUMN
    
    ' 出力設定
    configDict(CFG_OUTPUT_FILENAME_FORMAT) = DEFAULT_OUTPUT_FILENAME_FORMAT
    configDict(CFG_INCLUDE_LOG_SHEET) = IIf(DEFAULT_INCLUDE_LOG_SHEET, "TRUE", "FALSE")
    
    Set ConfigData = configDict
    
    Call LogMessage("デフォルト設定を使用します", LOG_LEVEL_INFO)
    
End Sub

'--------------------------------------------------
' 設定検証
'--------------------------------------------------
Private Function ValidateConfig() As Boolean
    
    Dim requiredKeys As Variant
    Dim i As Long
    
    requiredKeys = Array(CFG_EXCEL1_HEADER_ROWS, CFG_EXCEL1_DATA_START_ROW, _
                        CFG_EXCEL1_ID_COLUMN, CFG_EXCEL2_HEADER_ROWS, _
                        CFG_EXCEL2_DATA_START_ROW, CFG_EXCEL2_ID_COLUMN)
    
    ValidateConfig = True
    
    If ConfigData Is Nothing Then
        ValidateConfig = False
        Exit Function
    End If
    
    For i = 0 To UBound(requiredKeys)
        If Not ConfigData.Exists(requiredKeys(i)) Then
            Call LogMessage("必須設定項目が不足: " & requiredKeys(i), LOG_LEVEL_WARNING)
            ValidateConfig = False
            Exit Function
        End If
    Next i
    
End Function

'--------------------------------------------------
' 設定値取得（文字列）
'--------------------------------------------------
Public Function GetConfigValue(ByVal key As String, _
                              Optional ByVal defaultValue As String = "") As String
    On Error Resume Next
    
    If ConfigData Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If
    
    If ConfigData.Exists(key) Then
        GetConfigValue = CStr(ConfigData(key))
    Else
        GetConfigValue = defaultValue
    End If
    
    On Error GoTo 0
End Function

'--------------------------------------------------
' 設定値取得（Long型）
'--------------------------------------------------
Public Function GetConfigValueLong(ByVal key As String, _
                                  Optional ByVal defaultValue As Long = 0) As Long
    Dim strValue As String
    
    On Error Resume Next
    
    strValue = GetConfigValue(key, CStr(defaultValue))
    
    If IsNumeric(strValue) Then
        GetConfigValueLong = CLng(strValue)
    Else
        GetConfigValueLong = defaultValue
    End If
    
    On Error GoTo 0
End Function

'--------------------------------------------------
' 設定値取得（Boolean型）
'--------------------------------------------------
Public Function GetConfigValueBool(ByVal key As String, _
                                  Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim strValue As String
    
    On Error Resume Next
    
    strValue = UCase(GetConfigValue(key, IIf(defaultValue, "TRUE", "FALSE")))
    
    GetConfigValueBool = (strValue = "TRUE" Or strValue = "1" Or strValue = "YES")
    
    On Error GoTo 0
End Function

'--------------------------------------------------
' 出力パス生成
'--------------------------------------------------
Public Function GetOutputPath() As String
    
    Dim outputDir As String
    Dim fileName As String
    Dim dateStr As String
    
    ' 出力ディレクトリ
    outputDir = ThisWorkbook.Path & "\" & OUTPUT_FOLDER_NAME
    
    ' ディレクトリ作成
    If Not FolderExists(outputDir) Then
        On Error Resume Next
        MkDir outputDir
        On Error GoTo 0
    End If
    
    ' ファイル名生成
    dateStr = Format(Now, TIMESTAMP_FORMAT_FILE)
    fileName = GetConfigValue(CFG_OUTPUT_FILENAME_FORMAT, DEFAULT_OUTPUT_FILENAME_FORMAT)
    fileName = Replace(fileName, "[DATE]", dateStr)
    
    GetOutputPath = outputDir & "\" & fileName
    
End Function

'--------------------------------------------------
' ファイル存在確認
'--------------------------------------------------
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

'--------------------------------------------------
' フォルダ存在確認
'--------------------------------------------------
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

'--------------------------------------------------
' 対応拡張子か確認
'--------------------------------------------------
Public Function IsSupportedExtension(ByVal filePath As String) As Boolean
    Dim ext As String
    Dim supported As Variant
    Dim i As Long
    
    ext = LCase(GetFileExtension(filePath))
    supported = Split(SUPPORTED_EXTENSIONS, ",")
    
    IsSupportedExtension = False
    
    For i = 0 To UBound(supported)
        If ext = LCase(supported(i)) Then
            IsSupportedExtension = True
            Exit Function
        End If
    Next i
End Function

'--------------------------------------------------
' ファイル拡張子取得
'--------------------------------------------------
Public Function GetFileExtension(ByVal filePath As String) As String
    Dim pos As Long
    pos = InStrRev(filePath, ".")
    If pos > 0 Then
        GetFileExtension = Mid(filePath, pos)
    Else
        GetFileExtension = ""
    End If
End Function
