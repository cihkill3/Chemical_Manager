$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$filePath = Join-Path $scriptDir "order book.xlsx"

$wb = $excel.Workbooks.Open($filePath, 0, $true)
$ws = $wb.Worksheets.Item("2026")

Write-Host "=== 2026 시트 진단 ==="

$hdrRow = 7
$colOrderNum = 8  # 주문번호
$colReagent = 18  # 시약
$colReceived = 19 # 인수

$lastRow = $ws.UsedRange.Rows.Count
$validCount = 0

for ($r = 8; $r -le $lastRow; $r++) {
    $orderNum = "$($ws.Cells($r, $colOrderNum).Value2)".Trim()
    $reagent = "$($ws.Cells($r, $colReagent).Value2)".Trim()
    $received = "$($ws.Cells($r, $colReceived).Value2)".Trim()
    
    if ($reagent -ne "" -or $received -ne "") {
        Write-Host "행 $r : 주문번호='$orderNum', 시약='$reagent', 인수='$received'"
        
        if (($reagent -eq "O" -or $reagent -eq "o") -and ($received -eq "O" -or $received -eq "o")) {
            Write-Host "  -> 조건 만족!"
            $validCount++
        }
    }
}

Write-Host "총 조건 만족 행 수: $validCount"

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
