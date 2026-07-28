# Test-Sync.ps1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$managerPath = (Get-ChildItem $scriptDir -Filter "*.xlsm" | Select-Object -First 1).FullName
$orderBookPath = (Get-ChildItem $scriptDir -Filter "order book.xlsx" | Select-Object -First 1).FullName

try {
    Write-Host "Opening 시약관리.xlsm..."
    $wb = $excel.Workbooks.Open($managerPath, 0, $false)
    
    $wsConfig = $wb.Worksheets.Item(1)
    
    Write-Host "Configuring settings..."
    # 원본파일 (B4)
    $wsConfig.Cells(4, 2).Value2 = $orderBookPath
    # 원본시트 (B5)
    $wsConfig.Cells(5, 2).Value2 = "2026"
    # 헤더행 (B6)
    $wsConfig.Cells(6, 2).Value2 = 7
    
    Write-Host "Running Sync Orders Macro..."
    try {
        $excel.Run("BtnSyncOrders")
        Write-Host "Macro finished."
    } catch {
        Write-Host "Macro error: $_"
    }
    
    Write-Host "Status after run: " $wsConfig.Cells(8, 2).Value2
    
    Write-Host "Log entries:"
    $wsLog = $wb.Worksheets.Item(3)
    $usedRows = $wsLog.UsedRange.Rows.Count
    for ($r = 2; $r -le $usedRows; $r++) {
        $logStr = ""
        for ($c = 1; $c -le 8; $c++) {
            $logStr += "$($wsLog.Cells($r, $c).Value2) | "
        }
        Write-Host "Row ${r}: $logStr"
    }

    $wb.Save()
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
