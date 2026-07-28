Attribute VB_Name = "modColor"
Option Explicit

'=============================================================================
' modColor - 색상 상수 및 적용 함수
' 입고목록의 CAS/품번 누락 표시, 원본 복사완료 표시
'=============================================================================

' ─── 동적 색상 변수 ──────────────────────────────────────────────────────────────
Private m_CASMissingBg   As Long
Private m_CASMissingFont As Long

Private m_NumMissingBg   As Long
Private m_NumMissingFont As Long

Private m_BothMissingBg   As Long
Private m_BothMissingFont As Long

Private m_SourceDoneBg   As Long
Private m_SourceDoneFont As Long

' ─────────────────────────────────────────────────────────────────────────────
' 설정 시트의 범례에서 색상을 읽어와 캐싱한다. (동기화 시작 시 1회 호출)
' ─────────────────────────────────────────────────────────────────────────────
Public Sub LoadColors()
    Dim wsConfig As Worksheet
    Set wsConfig = modUtility.GetWorksheet(ThisWorkbook, "설정")
    If Not wsConfig Is Nothing Then
        ' 원본 복사 완료 (15행)
        m_SourceDoneBg = wsConfig.Cells(15, 1).Interior.Color
        m_SourceDoneFont = wsConfig.Cells(15, 1).Font.Color
        ' CAS 없음 (16행)
        m_CASMissingBg = wsConfig.Cells(16, 1).Interior.Color
        m_CASMissingFont = wsConfig.Cells(16, 1).Font.Color
        ' 품번 없음 (17행)
        m_NumMissingBg = wsConfig.Cells(17, 1).Interior.Color
        m_NumMissingFont = wsConfig.Cells(17, 1).Font.Color
        ' 둘 다 없음 (18행)
        m_BothMissingBg = wsConfig.Cells(18, 1).Interior.Color
        m_BothMissingFont = wsConfig.Cells(18, 1).Font.Color
    End If
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 입고목록 행에 CAS/품번 상태에 따른 색상 적용
' ─────────────────────────────────────────────────────────────────────────────
Public Sub ApplyRowColor(ByVal ws As Worksheet, ByVal rowIndex As Long, _
                         ByVal hasCAS As Boolean, ByVal hasNum As Boolean)
    On Error GoTo ErrHandler

    Dim lastCol As Long
    lastCol = modUtility.LastCol(ws, 1)
    If lastCol < 1 Then lastCol = 8  ' 헤더 열 수 기본값

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))

    If Not hasCAS And Not hasNum Then
        ' 둘 다 없음
        rng.Interior.Color = m_BothMissingBg
        rng.Font.Color = m_BothMissingFont
    ElseIf Not hasCAS Then
        ' CAS 없음
        rng.Interior.Color = m_CASMissingBg
        rng.Font.Color = m_CASMissingFont
    ElseIf Not hasNum Then
        ' 품번 없음
        rng.Interior.Color = m_NumMissingBg
        rng.Font.Color = m_NumMissingFont
    Else
        ' 정상 (변경 없음)
        rng.Interior.ColorIndex = xlNone
        rng.Font.ColorIndex = xlAutomatic
    End If

    Exit Sub
ErrHandler:
    Debug.Print "ApplyRowColor Error at row " & rowIndex & ": " & Err.Description
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 원본 오더북에서 성공적으로 복사된 행에 연두색 표시
'
' Parameters:
'   ws       - 원본 Worksheet
'   rowIndex - 색상을 적용할 행 번호
' ─────────────────────────────────────────────────────────────────────────────
Public Sub MarkSourceRowDone(ByVal ws As Worksheet, ByVal rowIndex As Long)
    On Error GoTo ErrHandler

    Dim lastCol As Long
    lastCol = modUtility.LastCol(ws, 1)
    If lastCol < 1 Then lastCol = 20

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
    
    rng.Interior.Color = m_SourceDoneBg
    rng.Font.Color = m_SourceDoneFont

    Exit Sub
ErrHandler:
    Debug.Print "MarkSourceRowDone Error at row " & rowIndex & ": " & Err.Description
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 입고목록 전체 색상 재계산 (전체 새로고침 시 사용)
'
' Parameters:
'   ws          - 입고목록 Worksheet
'   colCAS      - CAS 번호 열 번호
'   colItemNum  - 품번 열 번호
'   dataStartRow - 데이터 시작 행 번호 (헤더 제외)
' ─────────────────────────────────────────────────────────────────────────────
Public Sub RefreshAllColors(ByVal ws As Worksheet, ByVal colCAS As Long, _
                            ByVal colItemNum As Long, ByVal dataStartRow As Long)
    On Error GoTo ErrHandler

    modColor.LoadColors

    Dim lastDataRow As Long
    lastDataRow = modUtility.LastRow(ws, 1)
    If lastDataRow < dataStartRow Then Exit Sub

    Dim r As Long
    For r = dataStartRow To lastDataRow
        Dim hasCAS As Boolean
        Dim hasNum As Boolean
        hasCAS = (Trim(CStr(ws.Cells(r, colCAS).Value)) <> "")
        hasNum = (Trim(CStr(ws.Cells(r, colItemNum).Value)) <> "")
        ApplyRowColor ws, r, hasCAS, hasNum
    Next r

    Exit Sub
ErrHandler:
    Debug.Print "RefreshAllColors Error: " & Err.Description
End Sub
