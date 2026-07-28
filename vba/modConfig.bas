Attribute VB_Name = "modConfig"
Option Explicit

'=============================================================================
' modConfig - 설정 관리
' 파일 선택, 시트 선택, 설정 읽기/저장
'=============================================================================

' 설정 시트 이름 상수
Private Const SHEET_CONFIG   As String = "설정"

' 설정 시트 레이블 상수 (A열에 있는 항목명)
Private Const LABEL_SRC_FILE   As String = "원본파일"
Private Const LABEL_SRC_SHEET  As String = "원본시트"
Private Const LABEL_HEADER_ROW As String = "헤더행"        ' 원본 헤더 행 번호 (기본값 1)
Private Const LABEL_LAST_SYNC  As String = "마지막동기화"
Private Const LABEL_STATUS     As String = "상태"

' ─────────────────────────────────────────────────────────────────────────────
' 원본 파일 경로 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetSourceFile() As String
    GetSourceFile = GetConfigValue(LABEL_SRC_FILE)
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 원본 시트명 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetSourceSheet() As String
    GetSourceSheet = GetConfigValue(LABEL_SRC_SHEET)
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 원본 헤더 행 번호 반환
' 설정값이 없거나 0이면 기본값 1 반환
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetHeaderRow() As Long
    Dim val As String
    val = GetConfigValue(LABEL_HEADER_ROW)
    If val = "" Or Not IsNumeric(val) Then
        GetHeaderRow = 1   ' 기본값
    Else
        Dim n As Long
        n = CLng(val)
        GetHeaderRow = IIf(n < 1, 1, n)  ' 최소 1
    End If
End Function

' ─────────────────────────────────────────────────────────────────────────────
' [파일 선택] 버튼 핸들러
' Excel 파일 선택 대화상자 → 경로 저장
' ─────────────────────────────────────────────────────────────────────────────
Public Sub SelectSourceFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "원본 오더북 파일 선택 (OneDrive 동기화 로컬 경로)"
        .Filters.Clear
        .Filters.Add "Excel 파일", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False

        ' 현재 저장된 경로가 있으면 초기 폴더로 설정
        Dim savedPath As String
        savedPath = GetSourceFile()
        If savedPath <> "" Then
            Dim folderPath As String
            folderPath = Left(savedPath, InStrRev(savedPath, "\") - 1)
            If modUtility.FolderExists(folderPath) Then
                .InitialFileName = folderPath & "\"
            End If
        End If

        If .Show = True Then
            Dim selectedPath As String
            selectedPath = .SelectedItems(1)

            ' 웹 URL 경고 (OneDrive 웹 URL 방지)
            If InStr(selectedPath, "https://") > 0 Or InStr(selectedPath, "http://") > 0 Then
                MsgBox "웹 URL은 사용할 수 없습니다." & vbCrLf & _
                       "OneDrive가 PC에 동기화된 로컬 경로를 선택해주세요.", _
                       vbExclamation, "경로 오류"
                Exit Sub
            End If

            ' 경로 저장
            SetConfigValue LABEL_SRC_FILE, selectedPath
            SetConfigValue LABEL_STATUS, "파일 선택 완료"

            ' 시트명 초기화 (새 파일 선택 시)
            SetConfigValue LABEL_SRC_SHEET, ""

            MsgBox "파일이 선택되었습니다." & vbCrLf & vbCrLf & selectedPath & vbCrLf & vbCrLf & _
                   "이제 [시트 목록 읽기] 버튼을 눌러 시트를 선택해주세요.", _
                   vbInformation, "파일 선택 완료"
        End If
    End With
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' [시트 목록 읽기] 버튼 핸들러
' 원본 파일의 시트 목록을 UserForm으로 표시하고 사용자가 선택
' ─────────────────────────────────────────────────────────────────────────────
Public Sub ReadAndSelectSheet()
    Dim srcFile As String
    srcFile = GetSourceFile()

    If srcFile = "" Then
        MsgBox "먼저 [파일 선택] 버튼으로 원본 파일을 선택해주세요.", _
               vbExclamation, "파일 미선택"
        Exit Sub
    End If

    If Not modUtility.FileExists(srcFile) Then
        MsgBox "원본 파일을 찾을 수 없습니다." & vbCrLf & srcFile, _
               vbExclamation, "파일 없음"
        Exit Sub
    End If

    ' 원본 파일 열기 (읽기 전용, 화면 갱신 없이)
    Dim wb As Workbook
    Dim alreadyOpen As Boolean
    alreadyOpen = modUtility.IsWorkbookOpen(srcFile)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If Not alreadyOpen Then
        Set wb = Workbooks.Open(Filename:=srcFile, ReadOnly:=True, UpdateLinks:=False, Notify:=False)
    Else
        Set wb = modUtility.GetOpenWorkbook(srcFile)
    End If

    Application.ScreenUpdating = True

    If wb Is Nothing Then
        MsgBox "파일을 열 수 없습니다." & vbCrLf & srcFile, vbExclamation, "열기 실패"
        Exit Sub
    End If

    ' 시트 목록 수집
    Dim sheetNames() As String
    Dim i As Integer
    ReDim sheetNames(1 To wb.Worksheets.Count)
    For i = 1 To wb.Worksheets.Count
        sheetNames(i) = wb.Worksheets(i).Name
    Next i

    ' 원본 파일이 원래 닫혀 있었으면 다시 닫기
    If Not alreadyOpen Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If

    ' UserForm 표시
    Dim frm As frmSelectSheet
    Set frm = New frmSelectSheet
    frm.LoadSheets sheetNames
    frm.Show

    ' 선택 결과 저장
    If frm.SelectedSheet <> "" Then
        SetConfigValue LABEL_SRC_SHEET, frm.SelectedSheet
        SetConfigValue LABEL_STATUS, "시트 선택 완료 → 동기화 준비됨"
        MsgBox "시트가 선택되었습니다: [" & frm.SelectedSheet & "]" & vbCrLf & vbCrLf & _
               "이제 [주문 동기화] 버튼을 눌러 동기화를 실행해주세요.", _
               vbInformation, "시트 선택 완료"
    End If

    Unload frm
    Set frm = Nothing
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 마지막 동기화 시각 업데이트
' ─────────────────────────────────────────────────────────────────────────────
Public Sub UpdateLastSync(ByVal statusMsg As String)
    SetConfigValue LABEL_LAST_SYNC, Format(Now, "yyyy-mm-dd hh:mm:ss")
    SetConfigValue LABEL_STATUS, statusMsg
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 설정값 읽기 (A열 레이블 기준으로 B열에서 값 반환)
' ─────────────────────────────────────────────────────────────────────────────
Public Function GetConfigValue(ByVal labelName As String) As String
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = modUtility.GetWorksheet(ThisWorkbook, SHEET_CONFIG)
    If ws Is Nothing Then
        GetConfigValue = ""
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = modUtility.LastRow(ws, 1)

    Dim i As Long
    For i = 1 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = labelName Then
            GetConfigValue = Trim(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i

    GetConfigValue = ""
    Exit Function
ErrHandler:
    GetConfigValue = ""
End Function

' ─────────────────────────────────────────────────────────────────────────────
' 설정값 저장 (A열 레이블 기준으로 B열에 값 설정)
' ─────────────────────────────────────────────────────────────────────────────
Public Sub SetConfigValue(ByVal labelName As String, ByVal newValue As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = modUtility.GetWorksheet(ThisWorkbook, SHEET_CONFIG)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modUtility.LastRow(ws, 1)

    Dim i As Long
    For i = 1 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = labelName Then
            ws.Cells(i, 2).Value = newValue
            Exit Sub
        End If
    Next i

    ' 없으면 새 행 추가
    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, 1).Value = labelName
    ws.Cells(newRow, 2).Value = newValue

    Exit Sub
ErrHandler:
    ' 설정 저장 실패는 무시 (로그에 기록되지 않도록)
    Debug.Print "SetConfigValue Error: " & Err.Description
End Sub
