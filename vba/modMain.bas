Attribute VB_Name = "modMain"
Option Explicit

'=============================================================================
' modMain - 진입점 및 버튼 핸들러
' 시약관리 시스템의 메인 진입점 모음
'=============================================================================

' ─────────────────────────────────────────────────────────────────────────────
' 버튼: [파일 선택]
' 설정 시트에서 원본 파일 경로를 사용자가 직접 선택하도록 처리
' ─────────────────────────────────────────────────────────────────────────────
Public Sub BtnSelectFile()
    On Error GoTo ErrHandler
    modConfig.SelectSourceFile
    Exit Sub
ErrHandler:
    modUtility.ShowError "파일 선택 중 오류가 발생했습니다.", Err.Number, Err.Description
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 버튼: [시트 목록 읽기]
' 원본 파일을 열어 시트 목록을 읽고 사용자가 선택하도록 처리
' ─────────────────────────────────────────────────────────────────────────────
Public Sub BtnReadSheets()
    On Error GoTo ErrHandler
    modConfig.ReadAndSelectSheet
    Exit Sub
ErrHandler:
    modUtility.ShowError "시트 목록 읽기 중 오류가 발생했습니다.", Err.Number, Err.Description
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 버튼: [주문 동기화]
' 원본 오더북에서 데이터를 읽어 입고목록을 업데이트
' ─────────────────────────────────────────────────────────────────────────────
Public Sub BtnSyncOrders()
    On Error GoTo ErrHandler

    ' 설정 검증
    Dim srcFile As String
    Dim srcSheet As String
    srcFile = modConfig.GetSourceFile()
    srcSheet = modConfig.GetSourceSheet()

    If srcFile = "" Then
        MsgBox "원본 파일을 먼저 선택해주세요." & vbCrLf & "[파일 선택] 버튼을 눌러주세요.", _
               vbExclamation, "설정 필요"
        Exit Sub
    End If

    If srcSheet = "" Then
        MsgBox "원본 시트를 먼저 선택해주세요." & vbCrLf & "[시트 목록 읽기] 버튼을 눌러주세요.", _
               vbExclamation, "설정 필요"
        Exit Sub
    End If

    If Not modUtility.FileExists(srcFile) Then
        MsgBox "원본 파일을 찾을 수 없습니다." & vbCrLf & srcFile, _
               vbExclamation, "파일 없음"
        Exit Sub
    End If

    ' 동기화 실행
    modImport.SyncOrders srcFile, srcSheet

    Exit Sub
ErrHandler:
    modUtility.ShowError "주문 동기화 중 오류가 발생했습니다.", Err.Number, Err.Description
End Sub
