#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bootstrap script: Writes Build-ChemManager.ps1 with proper UTF-8 BOM encoding
so PowerShell can read Korean characters correctly.
Run: python write_builder.py
"""
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
output_ps1 = os.path.join(script_dir, "Build-ChemManager.ps1")

# The PowerShell script content with Korean strings
ps1_content = r"""# =============================================================================
# Build-ChemManager.ps1  (UTF-8 BOM)
# 시약관리.xlsm - Excel COM 자동화 빌더
# =============================================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path $scriptDir "시약관리.xlsm"
$vbaDir     = Join-Path $scriptDir "vba"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " 시약관리.xlsm 빌드 시작" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    Write-Host "[1/7] 새 워크북 생성..." -ForegroundColor Yellow
    $wb = $excel.Workbooks.Add()

    Write-Host "[2/7] 시트 구성..." -ForegroundColor Yellow
    $wsConfig = $wb.Worksheets.Item(1)
    $wsConfig.Name = "설정"
    $wsLog = $wb.Worksheets.Add([System.Type]::Missing, $wsConfig)
    $wsLog.Name = "로그"

    # =========================================================================
    # [설정] 시트
    # =========================================================================
    Write-Host "[3/7] 설정 시트 구성..." -ForegroundColor Yellow
    $wsConfig.Activate()

    $wsConfig.Range("A1").Value2 = "연구실 시약 주문 관리 시스템"
    $wsConfig.Range("A1:F1").Merge() | Out-Null
    $wsConfig.Range("A1").Font.Size = 16
    $wsConfig.Range("A1").Font.Bold = $true
    $wsConfig.Range("A1").Font.Color = 3093272
    $wsConfig.Range("A1").Interior.Color = 15138834
    $wsConfig.Range("A1").HorizontalAlignment = -4108
    $wsConfig.Rows(1).RowHeight = 36

    $wsConfig.Cells(3,1).Value2 = "항목"
    $wsConfig.Cells(3,2).Value2 = "값"
    $wsConfig.Range("A3:B3").Font.Bold = $true
    $wsConfig.Range("A3:B3").Interior.Color = 5025616
    $wsConfig.Range("A3:B3").Font.Color = 16777215
    $wsConfig.Range("A3:B3").HorizontalAlignment = -4108
    $wsConfig.Rows(3).RowHeight = 24

    $labels   = @("원본파일","원본시트","헤더행","마지막동기화","상태")
    $defaults = @("","","1","","파일을 선택하여 시작하세요")
    for ($i=0; $i -lt $labels.Length; $i++) {
        $row = $i + 4
        $wsConfig.Cells($row,1).Value2 = $labels[$i]
        $wsConfig.Cells($row,2).Value2 = $defaults[$i]
        $wsConfig.Cells($row,1).Font.Bold = $true
        $wsConfig.Cells($row,1).Interior.Color = 15921906
        $wsConfig.Rows($row).RowHeight = 24
    }
    # 원본파일/원본시트 강조 (노란 배경)
    $wsConfig.Cells(4,2).Interior.Color = 16777168
    $wsConfig.Cells(5,2).Interior.Color = 16777168
    # 헤더행: 연한 하늘색 배경, 중앙정렬, 숫자 유효성 검사
    $wsConfig.Cells(6,2).Interior.Color = 13433599   # 연한 하늘색
    $wsConfig.Cells(6,2).HorizontalAlignment = -4108  # xlCenter
    $wsConfig.Cells(6,2).Font.Bold = $true
    # 데이터 유효성 검사: 1~20 정수
    $valCell = $wsConfig.Cells(6,2)
    $valCell.Validation.Delete()
    $valCell.Validation.Add(1, 1, 1, 1, 20)  # xlValidateWholeNumber, between 1..20
    $valCell.Validation.InputTitle = "헤더 행 번호"
    $valCell.Validation.InputMessage = "원본 오더북에서 헤더가 있는 행 번호를 입력하세요. (1~20)"

    $wsConfig.Columns("A").ColumnWidth = 18
    $wsConfig.Columns("B").ColumnWidth = 68
    $wsConfig.Columns("C").ColumnWidth = 3
    $wsConfig.Columns("D").ColumnWidth = 20

    $rngBorder = $wsConfig.Range("A3:B8")
    $rngBorder.BorderAround(1,2) | Out-Null
    $rngBorder.Borders.Item(11).LineStyle = 1
    $rngBorder.Borders.Item(12).LineStyle = 1

    $wsConfig.Cells(9,1).Value2 = "[ 버튼 안내 ]"
    $wsConfig.Cells(9,1).Font.Bold = $true
    $wsConfig.Cells(9,1).Font.Color = 10066329

    $guide1 = "1  파일 선택      >>  OneDrive 동기화된 오더북(.xlsx)을 선택합니다."
    $guide2 = "2  시트 목록 읽기 >>  파일의 시트 목록을 불러와 동기화할 시트를 지정합니다."
    $guide3 = "3  주문 동기화    >>  ChemicalList.xlsx 로 데이터를 자동 동기화합니다."
    $wsConfig.Cells(10,1).Value2 = $guide1
    $wsConfig.Cells(11,1).Value2 = $guide2
    $wsConfig.Cells(12,1).Value2 = $guide3

    $wsConfig.Cells(14,1).Value2 = "[ 색상 범례 ]"
    $wsConfig.Cells(14,1).Font.Bold = $true
    $wsConfig.Cells(14,1).Font.Color = 10066329

    $wsConfig.Cells(15,1).Value2 = "연두색"
    $wsConfig.Cells(15,2).Value2 = "원본 오더북 - 복사 완료 행"
    $wsConfig.Cells(15,1).Interior.Color = 13434828
    $wsConfig.Cells(15,1).Font.Bold = $true

    $wsConfig.Cells(16,1).Value2 = "노란색"
    $wsConfig.Cells(16,2).Value2 = "ChemicalList - CAS 번호 없음"
    $wsConfig.Cells(16,1).Interior.Color = 16776960
    $wsConfig.Cells(16,1).Font.Bold = $true

    $wsConfig.Cells(17,1).Value2 = "주황색"
    $wsConfig.Cells(17,2).Value2 = "ChemicalList - 품번 없음"
    $wsConfig.Cells(17,1).Interior.Color = 10961922
    $wsConfig.Cells(17,1).Font.Bold = $true

    $wsConfig.Cells(18,1).Value2 = "빨간색"
    $wsConfig.Cells(18,2).Value2 = "ChemicalList - CAS 번호 + 품번 둘 다 없음"
    $wsConfig.Cells(18,1).Interior.Color = 16711680
    $wsConfig.Cells(18,1).Font.Bold = $true

    # =========================================================================
    # 상세 가이드 및 참고사항
    # =========================================================================
    $wsConfig.Cells(21,1).Value2 = "[ 매크로 사용 가이드 및 참고 사항 ]"
    $wsConfig.Cells(21,1).Font.Bold = $true
    $wsConfig.Cells(21,1).Font.Color = 10066329
    
    $notes = @(
        "1. 동기화 대상 (ChemicalList.xlsx)",
        "   - 동기화를 실행하면 오더북(원본 파일)이 위치한 폴더에 'ChemicalList.xlsx' 파일이 자동 생성되거나 업데이트됩니다.",
        "   - 기존 파일이 있을 경우 'ChemicalList_old_날짜_시간.xlsx' 이름으로 자동 백업되므로 데이터 유실 걱정이 없습니다.",
        "   - ChemicalList.xlsx의 기존 텍스트 서식, 배경색, 메모 등은 그대로 유지되며, 새로운 데이터는 맨 아래에 이어서 추가(Append)됩니다.",
        "",
        "2. 자동화 처리 내역",
        "   - ChemicalList에 신규 추가되는 항목은 가로 가운데 정렬이 자동으로 적용됩니다.",
        "   - ChemicalList의 모든 'Status' 열에는 Quantity(수량)가 Disposed(폐기)보다 크면 'O', 아니면 'X'인 수식이 자동 입력됩니다.",
        "   - 원본 오더북에서 동기화가 무사히 완료된 행은 연두색으로 칠해져 진행 상태를 쉽게 파악할 수 있습니다.",
        "",
        "3. 주의 및 참고 사항",
        "   - 원본 오더북 파일은 열어둔 상태로 동기화를 진행해도 안전하게 처리됩니다.",
        "   - 단, 대상 파일인 ChemicalList.xlsx 파일이 열려있는 상태라면, 매크로가 덮어쓰지 못하므로 닫아주신 후 진행해 주세요.",
        "   - CAS 번호나 품번이 누락된 항목은 ChemicalList에서 경고 색상으로 표시됩니다. 위 색상 범례를 참고하여 데이터를 보완해 주세요."
    )

    for ($i=0; $i -lt $notes.Length; $i++) {
        $r = 22 + $i
        $wsConfig.Cells($r, 1).Value2 = $notes[$i]
        $wsConfig.Range($wsConfig.Cells($r, 1), $wsConfig.Cells($r, 4)).Merge() | Out-Null
    }

    # 버튼 추가 (헤더행이 추가되어 행 번호 한 칸씩 이동)
    $btnLeft   = $wsConfig.Range("D4").Left
    $btnWidth  = 130
    $btnHeight = 26

    $btn1 = $wsConfig.Buttons().Add($btnLeft, $wsConfig.Range("D4").Top, $btnWidth, $btnHeight)
    $btn1.Caption = "파일 선택"
    $btn1.Name = "btnSelectFile"
    $btn1.OnAction = "BtnSelectFile"
    $btn1.Font.Size = 10
    $btn1.Font.Bold = $true

    $btn2 = $wsConfig.Buttons().Add($btnLeft, $wsConfig.Range("D5").Top, $btnWidth, $btnHeight)
    $btn2.Caption = "시트 목록 읽기"
    $btn2.Name = "btnReadSheets"
    $btn2.OnAction = "BtnReadSheets"
    $btn2.Font.Size = 10
    $btn2.Font.Bold = $true

    $btn3 = $wsConfig.Buttons().Add($btnLeft, $wsConfig.Range("D7").Top, $btnWidth, $btnHeight)
    $btn3.Caption = "주문 동기화"
    $btn3.Name = "btnSyncOrders"
    $btn3.OnAction = "BtnSyncOrders"
    $btn3.Font.Size = 10
    $btn3.Font.Bold = $true

    $wsConfig.Tab.Color = 5025616



    # =========================================================================
    # [로그] 시트
    # =========================================================================
    Write-Host "[5/7] 로그 시트 구성..." -ForegroundColor Yellow
    $wsLog.Activate()
    $logHdr = @("일시","파일명","시트명","신규","업데이트","중복","CAS누락","품번누락","처리시간")
    for ($c=0; $c -lt $logHdr.Length; $c++) { $wsLog.Cells(1,$c+1).Value2 = $logHdr[$c] }
    $wsLog.Range("A1:I1").Font.Bold = $true
    $wsLog.Range("A1:I1").Font.Color = 16777215
    $wsLog.Range("A1:I1").Interior.Color = 4209337
    $wsLog.Range("A1:I1").HorizontalAlignment = -4108
    $wsLog.Rows(1).RowHeight = 26
    $logW = @(22,25,18,10,12,10,12,12,12)
    for ($c=0; $c -lt $logW.Length; $c++) { $wsLog.Columns($c+1).ColumnWidth = $logW[$c] }
    $wsLog.Tab.Color = 6316128
    $wsLog.Range("A1").AutoFilter() | Out-Null
    $wsLog.Application.ActiveWindow.FreezePanes = $false
    $wsLog.Range("A2").Select() | Out-Null
    $wsLog.Application.ActiveWindow.FreezePanes = $true

    # =========================================================================
    # VBA 코드 삽입
    # =========================================================================
    Write-Host "[6/7] VBA 코드 삽입..." -ForegroundColor Yellow
    $vbaProject = $wb.VBProject
    try { $vbaProject.VBComponents.Remove($vbaProject.VBComponents.Item("Module1")) } catch {}

    $vbext_ct_StdModule = 1
    $vbext_ct_MSForm    = 3

    # 일반 모듈 삽입
    $modules = @("modUtility","modColor","modConfig","modImport","modMain")
    foreach ($modName in $modules) {
        $basFile = Join-Path $vbaDir "$modName.bas"
        if (Test-Path $basFile) {
            Write-Host "  -> $modName 삽입..." -ForegroundColor Gray
            $code = [System.IO.File]::ReadAllText($basFile, [System.Text.Encoding]::UTF8)
            $code = ($code -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute VB_Name' }) -join "`r`n"
            $mod = $vbaProject.VBComponents.Add($vbext_ct_StdModule)
            $mod.Name = $modName
            $mod.CodeModule.AddFromString($code)
        } else {
            Write-Host "  [경고] $basFile 없음" -ForegroundColor DarkYellow
        }
    }

    # frmSelectSheet
    Write-Host "  -> frmSelectSheet 추가..." -ForegroundColor Gray
    $frmSS = $vbaProject.VBComponents.Add($vbext_ct_MSForm)
    $frmSS.Name = "frmSelectSheet"
    $frmSS.Properties.Item("Caption").Value = "시트 선택"
    $frmSS.Properties.Item("Width").Value   = 290
    $frmSS.Properties.Item("Height").Value  = 270
    $ctlSS = $frmSS.Designer
    $lbl0 = $ctlSS.Controls.Add("Forms.Label.1","lblTitle",$true)
    $lbl0.Caption = "원본 파일의 시트를 선택하세요:"
    $lbl0.Left=6; $lbl0.Top=6; $lbl0.Width=270; $lbl0.Height=18
    $lst = $ctlSS.Controls.Add("Forms.ListBox.1","lstSheet",$true)
    $lst.Left=6; $lst.Top=28; $lst.Width=270; $lst.Height=168
    $okBtn = $ctlSS.Controls.Add("Forms.CommandButton.1","btnOK",$true)
    $okBtn.Caption="확인"; $okBtn.Left=108; $okBtn.Top=204; $okBtn.Width=72; $okBtn.Height=24
    $cxBtn = $ctlSS.Controls.Add("Forms.CommandButton.1","btnCancel",$true)
    $cxBtn.Caption="취소"; $cxBtn.Left=192; $cxBtn.Top=204; $cxBtn.Width=72; $cxBtn.Height=24

    $frmSSFile = Join-Path $vbaDir "frmSelectSheet.frm"
    if (Test-Path $frmSSFile) {
        $raw = [System.IO.File]::ReadAllText($frmSSFile, [System.Text.Encoding]::UTF8)
        $lines = $raw -split "`r?`n"
        $start = 0
        for ($li=0; $li -lt $lines.Length; $li++) { if ($lines[$li] -match '^Attribute VB_') { $start = $li+1 } }
        $frmSS.CodeModule.AddFromString(($lines[$start..($lines.Length-1)]) -join "`r`n")
    }

    # frmProgress
    Write-Host "  -> frmProgress 추가..." -ForegroundColor Gray
    $frmProg = $vbaProject.VBComponents.Add($vbext_ct_MSForm)
    $frmProg.Name = "frmProgress"
    $frmProg.Properties.Item("Caption").Value = "동기화 진행 중..."
    $frmProg.Properties.Item("Width").Value   = 360
    $frmProg.Properties.Item("Height").Value  = 145
    $ctlProg = $frmProg.Designer
    $lblPrg = $ctlProg.Controls.Add("Forms.Label.1","lblProgress",$true)
    $lblPrg.Caption = "[                    ]   0%"
    $lblPrg.Left=10; $lblPrg.Top=18; $lblPrg.Width=330; $lblPrg.Height=28
    $lblStt = $ctlProg.Controls.Add("Forms.Label.1","lblStatus",$true)
    $lblStt.Caption = "준비 중..."; $lblStt.Left=10; $lblStt.Top=54; $lblStt.Width=330; $lblStt.Height=20

    $frmProgFile = Join-Path $vbaDir "frmProgress.frm"
    if (Test-Path $frmProgFile) {
        $raw = [System.IO.File]::ReadAllText($frmProgFile, [System.Text.Encoding]::UTF8)
        $lines = $raw -split "`r?`n"
        $start = 0
        for ($li=0; $li -lt $lines.Length; $li++) { if ($lines[$li] -match '^Attribute VB_') { $start = $li+1 } }
        $frmProg.CodeModule.AddFromString(($lines[$start..($lines.Length-1)]) -join "`r`n")
    }

    # =========================================================================
    # 저장
    # =========================================================================
    Write-Host "[7/7] 저장 중..." -ForegroundColor Yellow
    $wsConfig.Activate()
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
    $wb.SaveAs($outputPath, 52)

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " 완료! 시약관리.xlsm 생성 성공" -ForegroundColor Green
    Write-Host " 경로: $outputPath" -ForegroundColor White
    Write-Host "==========================================" -ForegroundColor Cyan

} catch {
    Write-Host ""
    Write-Host "[오류] $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
"""

# Write with UTF-8 BOM (PowerShell requires BOM to detect UTF-8)
output_path = os.path.join(script_dir, "Build-ChemManager.ps1")
with open(output_path, "w", encoding="utf-8-sig") as f:
    f.write(ps1_content)

print(f"Written: {output_path}")
print("Encoding: UTF-8 with BOM")
