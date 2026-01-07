'==================================================
' modValidator - 検証処理モジュール
' Version: 2.0
' Date: 2026/01/07
'==================================================
Option Explicit

'--------------------------------------------------
' ファイル検証
'--------------------------------------------------
Public Function ValidateFiles(ByVal file1 As String, _
                            ByVal file2 As String) As Boolean
    
    Dim ext1 As String
    Dim ext2 As String
    
    ValidateFiles = False
    
    ' ファイル存在確認
    If Not FileExists(file1) Then
        Call LogMessage(ERR_FILE_NOT_FOUND & ": Excel1が見つかりません - " & file1, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    If Not FileExists(file2) Then
        Call LogMessage(ERR_FILE_NOT_FOUND & ": Excel2が見つかりません - " & file2, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' 拡張子確認（複数形式対応）
    If Not IsSupportedExtension(file1) Then
        Call LogMessage(ERR_INVALID_FORMAT & ": Excel1は対応形式ではありません (" & _
                       GetFileExtension(file1) & ")", LOG_LEVEL_ERROR)
        Call LogMessage("対応形式: " & SUPPORTED_EXTENSIONS, LOG_LEVEL_INFO)
        Exit Function
    End If
    
    If Not IsSupportedExtension(file2) Then
        Call LogMessage(ERR_INVALID_FORMAT & ": Excel2は対応形式ではありません (" & _
                       GetFileExtension(file2) & ")", LOG_LEVEL_ERROR)
        Call LogMessage("対応形式: " & SUPPORTED_EXTENSIONS, LOG_LEVEL_INFO)
        Exit Function
    End If
    
    ' 同一ファイルチェック
    If LCase(file1) = LCase(file2) Then
        Call LogMessage(ERR_SAME_FILE & ": 同じファイルが指定されています", LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' ファイルサイズチェック（オプション）
    If Not ValidateFileSize(file1) Then
        Exit Function
    End If
    
    If Not ValidateFileSize(file2) Then
        Exit Function
    End If
    
    ' ファイルアクセス確認
    If Not CanReadFile(file1) Then
        Call LogMessage(ERR_FILE_ACCESS & ": Excel1を読み取れません（使用中の可能性）", LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    If Not CanReadFile(file2) Then
        Call LogMessage(ERR_FILE_ACCESS & ": Excel2を読み取れません（使用中の可能性）", LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ValidateFiles = True
    Call LogMessage("ファイル検証OK", LOG_LEVEL_INFO)
    
End Function

'--------------------------------------------------
' ファイルサイズ検証
'--------------------------------------------------
Private Function ValidateFileSize(ByVal filePath As String) As Boolean
    Dim fso As Object
    Dim fileSize As Double
    Const MAX_FILE_SIZE As Double = 104857600 ' 100MB
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        fileSize = fso.GetFile(filePath).Size
        
        If fileSize > MAX_FILE_SIZE Then
            Call LogMessage("警告: ファイルサイズが大きいです (" & _
                           Format(fileSize / 1048576, "0.0") & " MB): " & _
                           GetFileName(filePath), LOG_LEVEL_WARNING)
        End If
        
        If fileSize = 0 Then
            Call LogMessage(ERR_INVALID_FORMAT & ": ファイルが空です - " & _
                           GetFileName(filePath), LOG_LEVEL_ERROR)
            Set fso = Nothing
            ValidateFileSize = False
            Exit Function
        End If
    End If
    
    Set fso = Nothing
    ValidateFileSize = True
    Exit Function
    
ErrorHandler:
    Set fso = Nothing
    ValidateFileSize = True ' サイズ取得に失敗しても処理は続行
End Function

'--------------------------------------------------
' ファイル読み取り可能確認
'--------------------------------------------------
Private Function CanReadFile(ByVal filePath As String) As Boolean
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Binary Access Read Lock Read As #fileNum
    Close #fileNum
    
    CanReadFile = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    CanReadFile = False
End Function

'--------------------------------------------------
' 設定値検証
'--------------------------------------------------
Public Function ValidateConfigValues() As Boolean
    Dim headerRows1 As Long
    Dim headerRows2 As Long
    Dim dataStart1 As Long
    Dim dataStart2 As Long
    
    ValidateConfigValues = False
    
    ' ヘッダー行数の妥当性
    headerRows1 = GetConfigValueLong(CFG_EXCEL1_HEADER_ROWS, DEFAULT_EXCEL1_HEADER_ROWS)
    headerRows2 = GetConfigValueLong(CFG_EXCEL2_HEADER_ROWS, DEFAULT_EXCEL2_HEADER_ROWS)
    
    If headerRows1 < 1 Or headerRows1 > 100 Then
        Call LogMessage("Excel1のヘッダー行数が不正です: " & headerRows1, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    If headerRows2 < 1 Or headerRows2 > 100 Then
        Call LogMessage("Excel2のヘッダー行数が不正です: " & headerRows2, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' データ開始行の妥当性
    dataStart1 = GetConfigValueLong(CFG_EXCEL1_DATA_START_ROW, DEFAULT_EXCEL1_DATA_START_ROW)
    dataStart2 = GetConfigValueLong(CFG_EXCEL2_DATA_START_ROW, DEFAULT_EXCEL2_DATA_START_ROW)
    
    If dataStart1 < 1 Then
        Call LogMessage("Excel1のデータ開始行が不正です: " & dataStart1, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    If dataStart2 < 1 Then
        Call LogMessage("Excel2のデータ開始行が不正です: " & dataStart2, LOG_LEVEL_ERROR)
        Exit Function
    End If
    
    ' データ開始行 > ヘッダー行数の確認
    If dataStart1 <= headerRows1 Then
        Call LogMessage("警告: Excel1のデータ開始行がヘッダー行数以下です", LOG_LEVEL_WARNING)
    End If
    
    If dataStart2 <= headerRows2 Then
        Call LogMessage("警告: Excel2のデータ開始行がヘッダー行数以下です", LOG_LEVEL_WARNING)
    End If
    
    ValidateConfigValues = True
    
End Function

'--------------------------------------------------
' 出力パス検証
'--------------------------------------------------
Public Function ValidateOutputPath(ByVal outputPath As String) As Boolean
    Dim outputDir As String
    Dim testFile As String
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    ValidateOutputPath = False
    
    ' ディレクトリ部分を取得
    outputDir = Left(outputPath, InStrRev(outputPath, "\") - 1)
    
    ' ディレクトリ存在確認
    If Not FolderExists(outputDir) Then
        ' 作成を試みる
        On Error Resume Next
        MkDir outputDir
        If Err.Number <> 0 Then
            Call LogMessage(ERR_OUTPUT_ACCESS & ": 出力ディレクトリを作成できません - " & outputDir, LOG_LEVEL_ERROR)
            Exit Function
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' 書き込み権限テスト
    testFile = outputDir & "\~test_" & Format(Now, "yyyymmddhhmmss") & ".tmp"
    
    fileNum = FreeFile
    Open testFile For Output As #fileNum
    Print #fileNum, "test"
    Close #fileNum
    
    ' テストファイル削除
    Kill testFile
    
    ValidateOutputPath = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    Kill testFile
    On Error GoTo 0
    
    Call LogMessage(ERR_OUTPUT_ACCESS & ": 出力先に書き込めません - " & outputDir, LOG_LEVEL_ERROR)
    ValidateOutputPath = False
End Function
