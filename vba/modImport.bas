Attribute VB_Name = "modImport"
Option Explicit

'=============================================================================
' modImport - 동기화 핵심 로직
' Workbook 열기, Array 읽기, Dictionary 생성, 동기화, 색상, 로그
'=============================================================================

' 시트명 상수
Private Const SHEET_IMPORT  As String = "시약리스트"
Private Const SHEET_LOG     As String = "로그"

' 입고목록 헤더명 상수
Private Const HDR_ORDER_NUM  As String = "Order No."
Private Const HDR_DATE       As String = "Order Date"
Private Const HDR_ORDERER    As String = "Ordered By"
Private Const HDR_ITEM_NAME  As String = "Product Name"
Private Const HDR_COMPANY    As String = "Manufacturer"
Private Const HDR_VOLUME     As String = "Package Size"
Private Const HDR_CAS        As String = "CAS No."
Private Const HDR_ITEM_NUM   As String = "Catalog No."
Private Const HDR_ROOM       As String = "Room"
Private Const HDR_TEMP       As String = "Storage Temp."
Private Const HDR_CABINET    As String = "Cabinet"
Private Const HDR_QTY        As String = "Quantity"
Private Const HDR_DISPOSED   As String = "Disposed"
Private Const HDR_STATUS     As String = "Status"
Private Const HDR_REMARK     As String = "Remarks"

' 원본(오더북) 헤더명 상수
Private Const SRC_ORDER_NUM  As String = "번호"
Private Const SRC_DATE       As String = "날짜"
Private Const SRC_ORDERER    As String = "주문자"
Private Const SRC_COMPANY    As String = "회사"
Private Const SRC_ITEM_NAME  As String = "품목명"
Private Const SRC_RECEIVE_DATE As String = "수령날짜"
Private Const SRC_CAS        As String = "CAS 번호"
Private Const SRC_ITEM_NUM   As String = "품번"
Private Const SRC_VOLUME     As String = "용량"
Private Const SRC_QTY        As String = "수량"
Private Const SRC_TEMP       As String = "보관온도"

'=============================================================================
' 메인 동기화 함수
' 원본 파일을 열어 Array로 읽고 Dictionary로 중복 관리 후 입고목록 업데이트
'=============================================================================
Public Sub SyncOrders(ByVal srcFilePath As String, ByVal srcSheetName As String)

    Dim tStart As Double
    tStart = Timer

    ' ── 진행 표시 폼 ────────────────────────────────────────────────────────
    Dim frm As frmProgress
    Set frm = New frmProgress
    frm.Show vbModeless
    frm.UpdateProgress 0, "동기화 준비 중..."
    DoEvents

    ' ── 성능 최적화 ON ──────────────────────────────────────────────────────
    modUtility.SetPerformanceMode True

    ' ── 결과 카운터 초기화 ──────────────────────────────────────────────────
    Dim cntNew       As Long  ' 신규 추가
    Dim cntUpdated   As Long  ' 업데이트
    Dim cntDuplicate As Long  ' 중복 (동일 데이터)
    Dim cntCASMiss   As Long  ' CAS 누락
    Dim cntNumMiss   As Long  ' 품번 누락
    Dim cntSkipped   As Long  ' 필터 제외 (시약/인수 미해당)

    Dim errMsg As String
    Dim success As Boolean
    success = True

    ' 동기화 시작 전 색상표 동적 로드
    modColor.LoadColors

    On Error GoTo ErrHandler

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 1: 원본 파일 열기
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 5, "원본 파일 열기 중..."
    DoEvents

    Dim srcWB As Workbook
    Dim alreadyOpen As Boolean
    alreadyOpen = modUtility.IsWorkbookOpen(srcFilePath)

    If Not alreadyOpen Then
        Set srcWB = Workbooks.Open(Filename:=srcFilePath, ReadOnly:=True, _
                                   UpdateLinks:=False, Notify:=False)
    Else
        Set srcWB = modUtility.GetOpenWorkbook(srcFilePath)
    End If

    If srcWB Is Nothing Then
        errMsg = "원본 파일을 열 수 없습니다." & vbCrLf & srcFilePath
        GoTo ErrCleanup
    End If

    ' ── 원본 시트 확인 ──────────────────────────────────────────────────────
    Dim srcWS As Worksheet
    Set srcWS = modUtility.GetWorksheet(srcWB, srcSheetName)

    If srcWS Is Nothing Then
        errMsg = "원본 시트를 찾을 수 없습니다." & vbCrLf & "시트명: " & srcSheetName
        GoTo ErrCleanup
    End If

    ' ─ config 설정값으로 먼저 시도, 실패 시 자동탐색(최대 20행) ─
    Dim srcHeaderRow As Long
    Dim configHdrRow As Long
    configHdrRow = modConfig.GetHeaderRow()  ' 설정 시트의 헤더행 값

    ' 설정값으로 검증
    Dim chkCol As Long
    chkCol = modUtility.FindColumn(srcWS, SRC_ORDER_NUM, configHdrRow)

    If chkCol > 0 Then
        ' 설정값이 정확함
        srcHeaderRow = configHdrRow
    Else
        ' 설정값 불일치 → 자동탐색 (1~20행)
        srcHeaderRow = 0
        Dim scanRow As Long
        For scanRow = 1 To 20
            If modUtility.FindColumn(srcWS, SRC_ORDER_NUM, scanRow) > 0 Then
                srcHeaderRow = scanRow
                Exit For
            End If
        Next scanRow

        If srcHeaderRow = 0 Then
            errMsg = "원본 시트에서 '번호' 헤더를 찾을 수 없습니다." & vbCrLf & _
                     "설정 시트의 [헤더행] 값을 확인하거나," & vbCrLf & _
                     "오더북에 '번호' 열이 있는지 확인해주세요."
            GoTo ErrCleanup
        End If

        ' 자동탐색 성공 → config 자동 업데이트
        modConfig.SetConfigValue "헤더행", CStr(srcHeaderRow)
        MsgBox "[헤더행] 자동 탐색 결과: " & srcHeaderRow & "행" & vbCrLf & _
               "설정 시트의 [헤더행] 값이 " & srcHeaderRow & "으로 자동 업데이트되었습니다.", _
               vbInformation, "헤더 행 자동 탐색"
    End If

    ' ── 원본 열 번호 조회 ───────────────────────────────────────────────────
    Dim srcColOrderNum  As Long
    Dim srcColDate      As Long
    Dim srcColOrderer   As Long
    Dim srcColCompany   As Long
    Dim srcColItemName  As Long
    Dim srcColReceiveDate As Long
    Dim srcColCAS       As Long
    Dim srcColItemNum   As Long
    Dim srcColVolume    As Long
    Dim srcColQty       As Long
    Dim srcColTemp      As Long

    srcColOrderNum = modUtility.FindColumn(srcWS, SRC_ORDER_NUM, srcHeaderRow)
    srcColDate     = modUtility.FindColumn(srcWS, SRC_DATE,      srcHeaderRow)
    srcColOrderer  = modUtility.FindColumn(srcWS, SRC_ORDERER,   srcHeaderRow)
    srcColCompany  = modUtility.FindColumn(srcWS, SRC_COMPANY,   srcHeaderRow)
    srcColItemName = modUtility.FindColumn(srcWS, SRC_ITEM_NAME, srcHeaderRow)
    srcColReceiveDate = modUtility.FindColumn(srcWS, SRC_RECEIVE_DATE, srcHeaderRow)
    srcColCAS      = modUtility.FindColumn(srcWS, SRC_CAS,       srcHeaderRow)
    srcColItemNum  = modUtility.FindColumn(srcWS, SRC_ITEM_NUM,  srcHeaderRow)
    srcColVolume   = modUtility.FindColumn(srcWS, SRC_VOLUME,    srcHeaderRow)
    srcColQty      = modUtility.FindColumn(srcWS, SRC_QTY,       srcHeaderRow)
    srcColTemp     = modUtility.FindColumn(srcWS, SRC_TEMP,      srcHeaderRow)

    ' 필수 열 존재 확인
    If srcColOrderNum = 0 Then
        errMsg = "원본 시트에 '번호' 열이 없습니다."
        GoTo ErrCleanup
    End If

    ' 수령 열이 없는 경우 경고 (선택적 열)
    Dim hasReceiveDateCol As Boolean
    hasReceiveDateCol = (srcColReceiveDate > 0)

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 3: 원본 전체 데이터를 Array로 읽기 (성능 최적화 핵심)
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 20, "원본 데이터 읽기 중..."
    DoEvents

    Dim srcLastRow As Long
    srcLastRow = modUtility.LastRow(srcWS, srcColOrderNum)

    If srcLastRow <= srcHeaderRow Then
        errMsg = "원본 시트에 데이터가 없습니다."
        GoTo ErrCleanup
    End If

    ' 최대 열 번호 계산 (열 번호 중 가장 큰 값 이상으로 설정)
    Dim srcMaxCol As Long
    srcMaxCol = modUtility.LastCol(srcWS, srcHeaderRow)
    ' 안전 마진: 시약/인수 열이 LastCol보다 클 수 있으므로 보정
    Dim colCheckMax As Long
    colCheckMax = srcColOrderNum
    If srcColDate     > colCheckMax Then colCheckMax = srcColDate
    If srcColOrderer  > colCheckMax Then colCheckMax = srcColOrderer
    If srcColCompany  > colCheckMax Then colCheckMax = srcColCompany
    If srcColItemName > colCheckMax Then colCheckMax = srcColItemName
    If srcColReceiveDate > colCheckMax Then colCheckMax = srcColReceiveDate
    If srcColCAS      > colCheckMax Then colCheckMax = srcColCAS
    If srcColItemNum  > colCheckMax Then colCheckMax = srcColItemNum
    If srcColVolume   > colCheckMax Then colCheckMax = srcColVolume
    If srcColQty      > colCheckMax Then colCheckMax = srcColQty
    If srcColTemp     > colCheckMax Then colCheckMax = srcColTemp
    If colCheckMax > srcMaxCol Then srcMaxCol = colCheckMax

    ' ★ 핵심: 전체 데이터를 한 번에 Variant Array로 읽기
    Dim srcData As Variant
    srcData = srcWS.Range( _
                srcWS.Cells(srcHeaderRow + 1, 1), _
                srcWS.Cells(srcLastRow, srcMaxCol) _
              ).Value

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 4: ChemicalList.xlsx 백업 및 열기 (또는 생성)
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 35, "ChemicalList 준비 중..."
    DoEvents

    Dim srcFolder As String
    srcFolder = Left(srcFilePath, InStrRev(srcFilePath, "\") - 1)
    
    Dim targetFilePath As String
    targetFilePath = srcFolder & "\ChemicalList.xlsx"
    
    Dim wbTarget As Workbook
    Dim wsImport As Worksheet
    
    If modUtility.FileExists(targetFilePath) Then
        ' 백업 파일 생성
        Dim backupTime As String
        backupTime = Format(Now, "yyyyMMdd_HHmmss")
        Dim backupPath As String
        backupPath = srcFolder & "\ChemicalList_old_" & backupTime & ".xlsx"
        
        On Error Resume Next
        FileCopy targetFilePath, backupPath
        On Error GoTo ErrHandler
        
        ' 기존 파일 열기
        Set wbTarget = Workbooks.Open(targetFilePath)
        Set wsImport = wbTarget.Sheets(1)
    Else
        ' 파일이 없으면 새로 생성
        Set wbTarget = Workbooks.Add
        Set wsImport = wbTarget.Sheets(1)
        wsImport.Name = "ChemicalList"
        
        ' 헤더 쓰기
        Dim hArr As Variant
        hArr = Array(HDR_ORDER_NUM, HDR_DATE, HDR_ORDERER, HDR_ITEM_NAME, HDR_COMPANY, _
                     HDR_VOLUME, HDR_CAS, HDR_ITEM_NUM, HDR_ROOM, HDR_TEMP, _
                     HDR_CABINET, HDR_QTY, HDR_DISPOSED, HDR_STATUS, HDR_REMARK)
        Dim hc As Long
        For hc = LBound(hArr) To UBound(hArr)
            wsImport.Cells(1, hc + 1).Value = hArr(hc)
        Next hc
        
        ' 기본 서식
        With wsImport.Range(wsImport.Cells(1, 1), wsImport.Cells(1, UBound(hArr) + 1))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(70, 70, 70)
            .HorizontalAlignment = -4108 ' xlCenter
        End With
    End If

    If wsImport Is Nothing Then
        errMsg = "ChemicalList 시트를 초기화할 수 없습니다."
        GoTo ErrCleanup
    End If

    ' 입고목록 헤더 열 번호 조회
    Dim impColItemName  As Long
    Dim impColCompany   As Long
    Dim impColVolume    As Long
    Dim impColCAS       As Long
    Dim impColItemNum   As Long
    Dim impColOrderNum  As Long
    Dim impColDate      As Long
    Dim impColOrderer   As Long
    Dim impColRoom      As Long
    Dim impColCabinet   As Long
    Dim impColTemp      As Long
    Dim impColQty       As Long
    Dim impColDisposed  As Long
    Dim impColStatus    As Long
    Dim impColRemark    As Long

    impColOrderNum = modUtility.FindColumn(wsImport, HDR_ORDER_NUM)
    impColDate     = modUtility.FindColumn(wsImport, HDR_DATE)
    impColOrderer  = modUtility.FindColumn(wsImport, HDR_ORDERER)
    impColItemName = modUtility.FindColumn(wsImport, HDR_ITEM_NAME)
    impColCompany  = modUtility.FindColumn(wsImport, HDR_COMPANY)
    impColVolume   = modUtility.FindColumn(wsImport, HDR_VOLUME)
    impColCAS      = modUtility.FindColumn(wsImport, HDR_CAS)
    impColItemNum  = modUtility.FindColumn(wsImport, HDR_ITEM_NUM)
    impColRoom     = modUtility.FindColumn(wsImport, HDR_ROOM)
    impColTemp     = modUtility.FindColumn(wsImport, HDR_TEMP)
    impColCabinet  = modUtility.FindColumn(wsImport, HDR_CABINET)
    impColQty      = modUtility.FindColumn(wsImport, HDR_QTY)
    impColDisposed = modUtility.FindColumn(wsImport, HDR_DISPOSED)
    impColStatus   = modUtility.FindColumn(wsImport, HDR_STATUS)
    impColRemark   = modUtility.FindColumn(wsImport, HDR_REMARK)

    If impColOrderNum = 0 Then
        errMsg = "ChemicalList 파일에 '" & HDR_ORDER_NUM & "' 헤더가 없습니다."
        GoTo ErrCleanup
    End If

    ' ★ 핵심: Scripting.Dictionary로 기존 입고목록 인덱싱
    ' Key: 번호(String), Value: 행 번호(Long)
    Dim dictImport As Object
    Set dictImport = CreateObject("Scripting.Dictionary")
    dictImport.CompareMode = vbTextCompare  ' 대소문자 무시

    Dim impLastRow As Long
    impLastRow = modUtility.LastRow(wsImport, impColOrderNum)

    Dim r As Long
    For r = 2 To impLastRow  ' 1행은 헤더
        Dim orderKey As String
        orderKey = modUtility.SafeStr(wsImport.Cells(r, impColOrderNum).Value)
        If orderKey <> "" Then
            If Not dictImport.Exists(orderKey) Then
                dictImport.Add orderKey, r
            End If
        End If
    Next r

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 5: 원본 Array 처리 → 동기화
    ' 조건: 시약=O AND 인수=O (열이 없으면 조건 무시)
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 50, "데이터 동기화 중..."
    DoEvents

    ' 원본 시트에 쓰기용 추적 (복사완료 표시)
    Dim srcDoneRows() As Long
    Dim srcDoneCount As Long
    ReDim srcDoneRows(1 To UBound(srcData, 1))

    Dim totalRows As Long
    totalRows = UBound(srcData, 1)

    Dim i As Long
    For i = 1 To totalRows

        ' 진행률 업데이트 (10행마다)
        If i Mod 100 = 0 Then
            Dim pct As Long
            pct = 50 + CLng((i / totalRows) * 40)
            frm.UpdateProgress pct, "동기화 중... (" & i & " / " & totalRows & "행)"
            DoEvents
        End If

        ' ── 필터 ─────────────────────────────────────────────────


        ' 수령 필터: 수령날짜 열이 비어있지 않은지 검사
        If hasReceiveDateCol Then
            Dim receiveDateVal As String
            receiveDateVal = Trim(modUtility.SafeStr(srcData(i, srcColReceiveDate)))
            If receiveDateVal = "" Then
                cntSkipped = cntSkipped + 1
                GoTo NextRow
            End If
        End If

        ' ── 번호 추출 ──────────────────────────────────────────────────
        Dim curOrderNum As String
        If srcColOrderNum > 0 And srcColOrderNum <= UBound(srcData, 2) Then
            curOrderNum = modUtility.SafeStr(srcData(i, srcColOrderNum))
        End If

        If curOrderNum = "" Then GoTo NextRow  ' 번호 없으면 건너뜀

        ' ── 각 필드 추출 ───────────────────────────────────────────────────
        Dim fItemName  As String
        Dim fCompany   As String
        Dim fVolume    As String
        Dim fCAS       As String
        Dim fItemNum   As String
        Dim fDate      As String
        Dim fOrderer   As String
        Dim fQty       As String
        Dim fTemp      As String

        fItemName = SafeGetArr(srcData, i, srcColItemName)
        fCompany  = SafeGetArr(srcData, i, srcColCompany)
        fVolume   = SafeGetArr(srcData, i, srcColVolume)
        fCAS      = SafeGetArr(srcData, i, srcColCAS)
        fItemNum  = SafeGetArr(srcData, i, srcColItemNum)
        fOrderer  = SafeGetArr(srcData, i, srcColOrderer)
        fQty      = SafeGetArr(srcData, i, srcColQty)
        fTemp     = SafeGetArr(srcData, i, srcColTemp)

        ' 날짜 형식 처리
        If srcColDate > 0 And srcColDate <= UBound(srcData, 2) Then
            Dim rawDate As Variant
            rawDate = srcData(i, srcColDate)
            If IsDate(rawDate) Then
                fDate = Format(CDate(rawDate), "yyyy-mm-dd")
            ElseIf IsNumeric(rawDate) And CDbl(rawDate) > 0 Then
                fDate = Format(CDate(rawDate), "yyyy-mm-dd")
            Else
                fDate = modUtility.SafeStr(rawDate)
            End If
        End If

        ' ── CAS / 품번 존재 여부 ───────────────────────────────────────────
        Dim bHasCAS As Boolean
        Dim bHasNum As Boolean
        bHasCAS = (fCAS <> "")
        bHasNum = (fItemNum <> "")

        If Not bHasCAS Then cntCASMiss = cntCASMiss + 1
        If Not bHasNum Then cntNumMiss = cntNumMiss + 1

        ' ── Dictionary 기반 신규/업데이트 판단 ────────────────────────────
        If dictImport.Exists(curOrderNum) Then
            ' 기존 행 → 업데이트
            Dim existRow As Long
            existRow = dictImport(curOrderNum)

            Dim changed As Boolean
            changed = False

            changed = UpdateCellIfChanged(wsImport, existRow, impColItemName, fItemName) Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColCompany,  fCompany)  Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColVolume,   fVolume)   Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColCAS,      fCAS)      Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColItemNum,  fItemNum)  Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColDate,     fDate)     Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColOrderer,  fOrderer)  Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColQty,      fQty)      Or changed
            changed = UpdateCellIfChanged(wsImport, existRow, impColTemp,     fTemp)     Or changed

            If changed Then
                cntUpdated = cntUpdated + 1
                ' 색상 재적용
                modColor.ApplyRowColor wsImport, existRow, bHasCAS, bHasNum
            Else
                cntDuplicate = cntDuplicate + 1
            End If

        Else
            ' 신규 행 → 추가
            Dim newRow As Long
            newRow = modUtility.AbsoluteLastRow(wsImport) + 1

            ' 신규 행 데이터 쓰기
            If impColItemName > 0 Then wsImport.Cells(newRow, impColItemName).Value = fItemName
            If impColCompany  > 0 Then wsImport.Cells(newRow, impColCompany).Value  = fCompany
            If impColVolume   > 0 Then wsImport.Cells(newRow, impColVolume).Value   = fVolume
            If impColCAS      > 0 Then wsImport.Cells(newRow, impColCAS).Value      = fCAS
            If impColItemNum  > 0 Then wsImport.Cells(newRow, impColItemNum).Value  = fItemNum
            If impColOrderNum > 0 Then wsImport.Cells(newRow, impColOrderNum).Value = curOrderNum
            If impColDate     > 0 Then wsImport.Cells(newRow, impColDate).Value     = fDate
            If impColOrderer  > 0 Then wsImport.Cells(newRow, impColOrderer).Value  = fOrderer
            If impColQty      > 0 Then wsImport.Cells(newRow, impColQty).Value      = fQty
            If impColTemp     > 0 Then wsImport.Cells(newRow, impColTemp).Value     = fTemp

            ' Dictionary에 추가
            dictImport.Add curOrderNum, newRow

            cntNew = cntNew + 1

            ' 색상 적용 및 가운데 정렬
            modColor.ApplyRowColor wsImport, newRow, bHasCAS, bHasNum
            wsImport.Range(wsImport.Cells(newRow, 1), wsImport.Cells(newRow, 15)).HorizontalAlignment = -4108 ' xlCenter
        End If

        ' 원본 행 복사완료 표시 대상으로 기록
        srcDoneCount = srcDoneCount + 1
        srcDoneRows(srcDoneCount) = srcHeaderRow + i  ' 실제 시트 행 번호

        NextRow:
    Next i

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 5.5: ChemicalList 색상 새로고침 및 Status 수식 적용
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 90, "ChemicalList 마무리 중..."
    DoEvents
    modColor.RefreshAllColors wsImport, impColCAS, impColItemNum, 2
    
    If impColStatus > 0 And impColQty > 0 And impColDisposed > 0 Then
        Dim finalLastRow As Long
        finalLastRow = modUtility.LastRow(wsImport, impColOrderNum)
        If finalLastRow >= 2 Then
            wsImport.Range(wsImport.Cells(2, impColStatus), wsImport.Cells(finalLastRow, impColStatus)).FormulaR1C1 = _
                "=IF(RC" & impColQty & ">RC" & impColDisposed & ",""O"",""X"")"
        End If
    End If

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 6: 원본 시트에 복사완료 색상 표시 (연두색)
    '         원본이 쓰기 가능한 경우에만 적용
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 92, "원본 완료 표시 중..."
    DoEvents

    If Not srcWS.Parent.ReadOnly And srcDoneCount > 0 Then
        Dim d As Long
        For d = 1 To srcDoneCount
            modColor.MarkSourceRowDone srcWS, srcDoneRows(d)
        Next d

        ' 원본 파일 저장 시도
        On Error Resume Next
        srcWB.Save
        On Error GoTo ErrHandler
    End If

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 7: 원본 파일 닫기 (원래 닫혀 있었으면)
    ' ════════════════════════════════════════════════════════════════════════
    If Not alreadyOpen And Not srcWB Is Nothing Then
        Application.DisplayAlerts = False
        srcWB.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Set srcWB = Nothing
    End If
    
    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 7.5: ChemicalList 저장 및 닫기
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 94, "ChemicalList 저장 중..."
    DoEvents
    
    If cntNew > 0 Or cntUpdated > 0 Then
        Application.DisplayAlerts = False
        wbTarget.SaveAs targetFilePath
        Application.DisplayAlerts = True
    End If
    
    wbTarget.Close SaveChanges:=False

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 8: 로그 기록
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 95, "로그 기록 중..."
    DoEvents

    Dim elapsed As Double
    elapsed = Timer - tStart

    WriteLog srcFilePath, srcSheetName, cntNew, cntUpdated, cntDuplicate, _
             cntCASMiss, cntNumMiss, elapsed

    ' ── 설정 업데이트 ───────────────────────────────────────────────────────
    modConfig.UpdateLastSync "동기화 완료 (신규:" & cntNew & " 업데이트:" & cntUpdated & ")"

    ' ════════════════════════════════════════════════════════════════════════
    ' STEP 9: 완료
    ' ════════════════════════════════════════════════════════════════════════
    frm.UpdateProgress 100, "완료!"
    DoEvents

    modUtility.SetPerformanceMode False

    Unload frm
    Set frm = Nothing

    ' 완료 메시지
    Dim summary As String
    summary = "═══════════════════════════════" & vbCrLf & _
              "  주문 동기화 완료" & vbCrLf & _
              "═══════════════════════════════" & vbCrLf & vbCrLf & _
              "  신규 추가   : " & cntNew      & "건" & vbCrLf & _
              "  업데이트    : " & cntUpdated  & "건" & vbCrLf & _
              "  중복(변경없음): " & cntDuplicate & "건" & vbCrLf & vbCrLf & _
              "  CAS 번호 누락: " & cntCASMiss & "건" & vbCrLf & _
              "  품번 누락   : " & cntNumMiss & "건" & vbCrLf & vbCrLf & _
              "  처리 시간   : " & Format(elapsed, "0.00") & "초" & vbCrLf & _
              "═══════════════════════════════"

    MsgBox summary, vbInformation, "동기화 완료"

    Exit Sub

ErrCleanup:
    modUtility.SetPerformanceMode False
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    If Not srcWB Is Nothing And Not alreadyOpen Then
        Application.DisplayAlerts = False
        srcWB.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    If Not wbTarget Is Nothing Then
        On Error Resume Next
        wbTarget.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox errMsg, vbCritical, "동기화 실패"
    Exit Sub

ErrHandler:
    success = False
    errMsg = Err.Description
    modUtility.SetPerformanceMode False
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
    If Not srcWB Is Nothing And Not alreadyOpen Then
        On Error Resume Next
        Application.DisplayAlerts = False
        srcWB.Close SaveChanges:=False
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If
    If Not wbTarget Is Nothing Then
        On Error Resume Next
        wbTarget.Close SaveChanges:=False
        On Error GoTo 0
    End If
    modUtility.ShowError "동기화 중 오류 발생", Err.Number, errMsg
End Sub

'=============================================================================
' 헤더 행 자동 탐색 (최대 maxRows행까지)
' 반환: 헤더가 있는 행 번호. 없으면 0.
'=============================================================================
Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal keyHeader As String, _
                               ByVal maxRows As Long) As Long
    Dim r As Long
    For r = 1 To maxRows
        Dim c As Long
        Dim lastC As Long
        lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastC
            If Trim(CStr(ws.Cells(r, c).Value)) = keyHeader Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    FindHeaderRow = 0
End Function

'=============================================================================
' Array에서 안전하게 값 읽기 (열 번호가 0이거나 범위 초과 시 빈 문자열)
'=============================================================================
Private Function SafeGetArr(ByRef arr As Variant, ByVal rowIdx As Long, _
                            ByVal colIdx As Long) As String
    If colIdx <= 0 Or colIdx > UBound(arr, 2) Then
        SafeGetArr = ""
    Else
        SafeGetArr = modUtility.SafeStr(arr(rowIdx, colIdx))
    End If
End Function

'=============================================================================
' 셀 값이 다를 때만 업데이트 (불필요한 쓰기 방지)
' 반환: True = 변경됨, False = 동일
'=============================================================================
Private Function UpdateCellIfChanged(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                     ByVal colNum As Long, ByVal newVal As String) As Boolean
    If colNum <= 0 Then
        UpdateCellIfChanged = False
        Exit Function
    End If

    Dim oldVal As String
    oldVal = modUtility.SafeStr(ws.Cells(rowNum, colNum).Value)

    If oldVal <> newVal Then
        ws.Cells(rowNum, colNum).Value = newVal
        UpdateCellIfChanged = True
    Else
        UpdateCellIfChanged = False
    End If
End Function

'=============================================================================
' 로그 시트에 동기화 결과 기록
'=============================================================================
Private Sub WriteLog(ByVal srcFile As String, ByVal srcSheet As String, _
                     ByVal cntNew As Long, ByVal cntUpdated As Long, _
                     ByVal cntDuplicate As Long, ByVal cntCASMiss As Long, _
                     ByVal cntNumMiss As Long, ByVal elapsed As Double)
    On Error GoTo ErrHandler

    Dim wsLog As Worksheet
    Set wsLog = modUtility.GetWorksheet(ThisWorkbook, SHEET_LOG)
    If wsLog Is Nothing Then Exit Sub

    ' 로그 열 번호 조회 (헤더 기반)
    Dim colTime     As Long
    Dim colFile     As Long
    Dim colSheet    As Long
    Dim colNew      As Long
    Dim colUpdated  As Long
    Dim colDup      As Long
    Dim colCASMiss  As Long
    Dim colNumMiss  As Long
    Dim colElapsed  As Long

    colTime    = modUtility.FindColumn(wsLog, "일시")
    colFile    = modUtility.FindColumn(wsLog, "파일명")
    colSheet   = modUtility.FindColumn(wsLog, "시트명")
    colNew     = modUtility.FindColumn(wsLog, "신규")
    colUpdated = modUtility.FindColumn(wsLog, "업데이트")
    colDup     = modUtility.FindColumn(wsLog, "중복")
    colCASMiss = modUtility.FindColumn(wsLog, "CAS누락")
    colNumMiss = modUtility.FindColumn(wsLog, "품번누락")
    colElapsed = modUtility.FindColumn(wsLog, "처리시간")

    ' 새 로그 행 추가
    Dim newRow As Long
    newRow = modUtility.LastRow(wsLog, 1) + 1

    If colTime    > 0 Then wsLog.Cells(newRow, colTime).Value    = Now
    If colFile    > 0 Then wsLog.Cells(newRow, colFile).Value    = modUtility.FileBaseName(srcFile)
    If colSheet   > 0 Then wsLog.Cells(newRow, colSheet).Value   = srcSheet
    If colNew     > 0 Then wsLog.Cells(newRow, colNew).Value     = cntNew
    If colUpdated > 0 Then wsLog.Cells(newRow, colUpdated).Value = cntUpdated
    If colDup     > 0 Then wsLog.Cells(newRow, colDup).Value     = cntDuplicate
    If colCASMiss > 0 Then wsLog.Cells(newRow, colCASMiss).Value = cntCASMiss
    If colNumMiss > 0 Then wsLog.Cells(newRow, colNumMiss).Value = cntNumMiss
    If colElapsed > 0 Then wsLog.Cells(newRow, colElapsed).Value = Format(elapsed, "0.00") & "초"

    ' 날짜 열 형식 적용
    If colTime > 0 Then
        wsLog.Cells(newRow, colTime).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End If

    Exit Sub
ErrHandler:
    Debug.Print "WriteLog Error: " & Err.Description
End Sub
