# Excel Trust Center: VBA 프로젝트 모델 접근 허용 설정
# 설치된 Excel 버전 감지 후 AccessVBOM 레지스트리 설정

$baseKey = "HKCU:\Software\Microsoft\Office"
$versionKeys = Get-ChildItem $baseKey -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
    Sort-Object { [double]$_.PSChildName } -Descending

Write-Host "발견된 Office 버전:" -ForegroundColor Cyan
foreach ($key in $versionKeys) {
    $ver = $key.PSChildName
    Write-Host "  - $ver"
    $regPath = "$baseKey\$ver\Excel\Security"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
    $check = (Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
    if ($check -eq 1) {
        Write-Host "    AccessVBOM=1 설정 완료" -ForegroundColor Green
    } else {
        Write-Host "    설정 실패" -ForegroundColor Red
    }
}
