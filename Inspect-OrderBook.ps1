# Inspect-OrderBook3.ps1 - 2025 시트 헤더 위치 확인 + 시약/인수 값 샘플 확인
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$filePath = "c:\Users\Jeonghun Lee\.gemini\antigravity\scratch\chemical manager\order book.xlsx"

try {
    $wb = $excel.Workbooks.Open($filePath, 0, $true)

    foreach ($wsIdx in @(1, 2)) {
        $ws = $wb.Worksheets.Item($wsIdx)
        Write-Host ""
        Write-Host "==============================" -ForegroundColor Cyan
        Write-Host "시트 [$($ws.Name)] 헤더 탐색" -ForegroundColor Cyan
        Write-Host "==============================" -ForegroundColor Cyan

        # 헤더행 탐색 (최대 20행)
        $hdrRow = 0
        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 25; $c++) {
                $v = $ws.Cells($r, $c).Value2
                if ("$v" -eq "주문번호") {
                    $hdrRow = $r
                    break
                }
            }
            if ($hdrRow -gt 0) { break }
        }

        if ($hdrRow -eq 0) {
            Write-Host "  [경고] 주문번호 열을 찾지 못했습니다!" -ForegroundColor Red
            continue
        }

        Write-Host "  헤더 행: $hdrRow 행"

        # 헤더 열 목록
        Write-Host "  헤더 목록:"
        for ($c = 1; $c -le 25; $c++) {
            $v = $ws.Cells($hdrRow, $c).Value2
            if ($v -ne $null -and "$v" -ne "") {
                Write-Host "    [열$c] $v"
            }
        }

        # 시약/인수 열 번호
        $colReagent  = 0
        $colReceived = 0
        $colOrderNum = 0
        for ($c = 1; $c -le 25; $c++) {
            $v = "$($ws.Cells($hdrRow, $c).Value2)"
            if ($v -eq "시약")    { $colReagent  = $c }
            if ($v -eq "인수")    { $colReceived = $c }
            if ($v -eq "주문번호") { $colOrderNum = $c }
        }
        Write-Host "  시약열: $colReagent  /  인수열: $colReceived  /  주문번호열: $colOrderNum"

        # 데이터 샘플 (헤더 다음 5행)
        Write-Host "  데이터 샘플 (헤더 다음 5행):"
        for ($r = $hdrRow+1; $r -le $hdrRow+5; $r++) {
            $reagentVal  = $ws.Cells($r, $colReagent).Value2
            $receivedVal = $ws.Cells($r, $colReceived).Value2
            $orderNum    = $ws.Cells($r, $colOrderNum).Value2
            $itemName    = $ws.Cells($r, 6).Value2
            Write-Host "    행$r : 주문번호=[$orderNum]  품목명=[$itemName]  시약=[$reagentVal]  인수=[$receivedVal]"
        }

        # O 값이 실제로 있는지 확인 (전체 스캔)
        $oCount = 0
        $lastRow = $ws.UsedRange.Rows.Count
        if ($colReagent -gt 0) {
            for ($r = $hdrRow+1; $r -le $lastRow; $r++) {
                $v = "$($ws.Cells($r, $colReagent).Value2)"
                if ($v -eq "O" -or $v -eq "o" -or $v.Length -gt 0) {
                    if ($oCount -lt 5) {
                        Write-Host "    [시약값 발견] 행$r : 시약=[$v]  (Len=$($v.Length)  Byte0=$([int][char]$v[0]))"
                    }
                    $oCount++
                }
            }
        }
        Write-Host "  시약 열에 값 있는 행 수: $oCount"
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
