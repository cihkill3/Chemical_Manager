$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$xlsmPath = "C:\Users\Jeonghun Lee\.gemini\antigravity\scratch\chemical manager\시약관리.xlsm"
$wb = $excel.Workbooks.Open($xlsmPath)

$wsConfig = $wb.Sheets.Item("설정")
$wsConfig.Cells.Item(7, 2).Value2 = "C:\Users\Jeonghun Lee\.gemini\antigravity\scratch\chemical manager\order book test.xlsx"
$wsConfig.Cells.Item(8, 2).Value2 = "Sheet1"

$excel.Run("SyncOrdersMacro")

$wb.Save()
$wb.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
