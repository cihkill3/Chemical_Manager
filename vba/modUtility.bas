Attribute VB_Name = "modUtility"
Option Explicit

'=============================================================================
' modUtility - 공통 유틸리티 함수
' FindColumn, LastRow, 파일/폴더 존재, 워크북 관련, 문자열 처리
'=============================================================================

' ─────────────────────────────────────────────────────────────────────────────
' 헤더명으로 열 번호를 반환한다.
' 반드시 열 번호 하드코딩 없이 이 함수를 사용해야 한다.
'
' Parameters:
'   ws         - 검색 대상 Worksheet
'   headerName - 찾을 헤더명
'   headerRow  - 헤더가 있는 행 번호 (기본값 1)
'
' Returns: 열 번호 (Long). 찾지 못하면 0 반환.
' ─────────────────────────────────────────────────────────────────────────────
Public Function FindColumn(ByVal ws As Worksheet, ByVal headerName As String, _
                           Optional ByVal headerRow As Long = 1) As Long
    On Error GoTo ErrHandler

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(headerRow, c).Value)) = Trim(headerName) Then
            FindColumn = c
            Exit Function
        End If
    Next c

    FindColumn = 0  ' 찾지 못함
    Exit Function
ErrHandler:
    FindColumn = 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 배열에서 헤더명으로 열 인덱스를 반환한다.
' (Array 기반 처리 시 사용)
'
' Parameters:
'   headerArr  - 헤더 1차원 배열 (1-based 또는 0-based)
'   headerName - 찾을 헤더명
'   baseIndex  - 배열 시작 인덱스 (기본값 1)
'
' Returns: 배열 인덱스. 찾지 못하면 -1 반환.
' ─────────────────────────────────────────────────────────────────────────────
Public Function FindColumnInArray(ByRef headerArr As Variant, ByVal headerName As String, _
                                  Optional ByVal baseIndex As Long = 1) As Long
    Dim i As Long
    For i = LBound(headerArr) To UBound(headerArr)
        If Trim(CStr(headerArr(i))) = Trim(headerName) Then
            FindColumnInArray = i
            Exit Function
        End If
    Next i
    FindColumnInArray = -1
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 지정 열(colNum)의 마지막 데이터 행 번호 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function LastRow(ByVal ws As Worksheet, Optional ByVal colNum As Long = 1) As Long
    On Error GoTo ErrHandler
    LastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    Exit Function
ErrHandler:
    LastRow = 1
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 시트 전체에서 내용이 있는 마지막 행 번호 반환 (어떤 열이든)
' ─────────────────────────────────────────────────────────────────────────────
Public Function AbsoluteLastRow(ByVal ws As Worksheet) As Long
    On Error GoTo ErrHandler
    Dim rng As Range
    Set rng = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rng Is Nothing Then
        AbsoluteLastRow = 1
    Else
        AbsoluteLastRow = rng.Row
    End If
    Exit Function
ErrHandler:
    AbsoluteLastRow = 1
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 지정 행(rowNum)의 마지막 데이터 열 번호 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function LastCol(ByVal ws As Worksheet, Optional ByVal rowNum As Long = 1) As Long
    On Error GoTo ErrHandler
    LastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    Exit Function
ErrHandler:
    LastCol = 1
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 파일 존재 여부 확인
' ─────────────────────────────────────────────────────────────────────────────
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 폴더 존재 여부 확인
' ─────────────────────────────────────────────────────────────────────────────
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 특정 경로의 워크북이 현재 열려 있는지 확인
' ─────────────────────────────────────────────────────────────────────────────
Public Function IsWorkbookOpen(ByVal filePath As String) As Boolean
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
    IsWorkbookOpen = False
    On Error GoTo 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 열려 있는 워크북 반환 (파일 경로 기준)
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetOpenWorkbook(ByVal filePath As String) As Workbook
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOpenWorkbook = Nothing
    On Error GoTo 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' Workbook에서 시트명으로 Worksheet 반환
' 시트가 없으면 Nothing 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    If Err.Number <> 0 Then
        Set GetWorksheet = Nothing
    Else
        Set GetWorksheet = ws
    End If
    On Error GoTo 0
End Function

' ─────────────────────────────────────────────────────────────────────────────
' Null / Empty 값을 빈 문자열로 변환하는 안전한 문자열 변환
' ─────────────────────────────────────────────────────────────────────────────
Public Function SafeStr(ByVal val As Variant) As String
    If IsNull(val) Or IsEmpty(val) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(val))
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 오류 메시지 표시 (표준 형식)
' ─────────────────────────────────────────────────────────────────────────────
Public Sub ShowError(ByVal msg As String, ByVal errNum As Long, ByVal errDesc As String)
    Dim fullMsg As String
    fullMsg = msg & vbCrLf & vbCrLf & _
              "오류 번호: " & errNum & vbCrLf & _
              "오류 내용: " & errDesc
    MsgBox fullMsg, vbCritical, "오류 발생"
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 파일명에서 확장자를 제외한 이름 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function FileBaseName(ByVal filePath As String) As String
    Dim fileName As String
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        FileBaseName = Left(fileName, dotPos - 1)
    Else
        FileBaseName = fileName
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' Application 성능 최적화 설정 ON/OFF
' ─────────────────────────────────────────────────────────────────────────────
Public Sub SetPerformanceMode(ByVal enable As Boolean)
    If enable Then
        ' 성능 최적화 ON
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = True
    Else
        ' 성능 최적화 OFF (원래대로 복원)
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
    End If
End Sub
