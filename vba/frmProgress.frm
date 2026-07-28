VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress
   Caption         =   "동기화 진행 중..."
   ClientHeight    =   2400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6480
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=============================================================================
' frmProgress - 진행률 표시 UserForm
' 텍스트 기반 진행 바 (■/□) 표시
'=============================================================================
Option Explicit

Private Const BAR_WIDTH As Integer = 20  ' 진행 바 전체 길이

' ─────────────────────────────────────────────────────────────────────────────
' 진행률 업데이트
' percent: 0~100
' statusText: 상태 메시지
' ─────────────────────────────────────────────────────────────────────────────
Public Sub UpdateProgress(ByVal percent As Integer, ByVal statusText As String)
    ' 범위 제한
    If percent < 0 Then percent = 0
    If percent > 100 Then percent = 100

    ' 진행 바 계산
    Dim filled As Integer
    filled = CLng((percent / 100) * BAR_WIDTH)
    Dim barEmpty As Integer
    barEmpty = BAR_WIDTH - filled

    ' 진행 바 문자열 생성
    Dim barStr As String
    barStr = "[" & String(filled, ChrW(9608)) & String(barEmpty, ChrW(9633)) & "] " & percent & "%"

    ' 라벨 업데이트
    lblProgress.Caption = barStr
    lblStatus.Caption = statusText

    ' 강제 화면 갱신
    Me.Repaint

End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 폼 초기화
' ─────────────────────────────────────────────────────────────────────────────
Private Sub UserForm_Initialize()
    Me.Caption = "주문 동기화 진행 중..."
    UpdateProgress 0, "준비 중..."
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 폼 닫기 버튼 막기 (동기화 중 임의 종료 방지)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True  ' X 버튼으로 닫기 방지
    End If
End Sub
