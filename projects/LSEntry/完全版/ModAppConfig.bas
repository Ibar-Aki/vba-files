Option Explicit
'
'===============================================================================
' モジュール名: ModAppConfig
' 共有定数・設定値を集約するモジュール
'===============================================================================

' エラー表示セル（全モジュールで参照）
Public Const ERR_CELL_ADDR As String = "J3"
Public Const DATA_ENTRY_DATE_CELL As String = "D4"

' --- シート名 Enum ---
Public Enum SheetName
    Sheet_DataEntry = 1
    Sheet_Monthly = 2
    Sheet_DataAcquire = 3
End Enum

' --- データ登録シートの列 Enum ---
Public Enum DataSheetColumn
    DataCol_WorkNo = 3
    DataCol_Category = 4
    DataCol_Time = 5
End Enum

' --- 月次データシートの列 Enum ---
Public Enum MonthlySheetColumn
    MonthlyCol_Date = 2
    MonthlyCol_Min = 3
End Enum

' --- 月次データシートの行 Enum ---
Public Enum MonthlySheetRow
    MonthlyRow_WorkNo = 10
    MonthlyRow_Header = 11
    MonthlyRow_DataStart = 12
End Enum

' --- シート名取得 ---
Public Function GetSheetName(ByVal sn As SheetName) As String
    Select Case sn
        Case Sheet_DataEntry: GetSheetName = "データ登録"
        Case Sheet_Monthly:   GetSheetName = "月次データ"
        Case Sheet_DataAcquire: GetSheetName = "データ取得"
    End Select
End Function

' --- ワークシート取得 ---
Public Function GetSheet(ByVal sn As SheetName) As Worksheet
    Set GetSheet = ThisWorkbook.Sheets(GetSheetName(sn))
End Function
