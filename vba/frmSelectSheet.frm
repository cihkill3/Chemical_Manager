VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSheet
   Caption         =   "시트 선택"
   ClientHeight    =   5400
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5760
   OleObjectBlob   =   "frmSelectSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=============================================================================
' frmSelectSheet - 시트 선택 UserForm
' 원본 파일의 시트 목록을 ListBox로 표시하고 사용자 선택 처리
'=============================================================================
Option Explicit

Private m_selectedSheet As String

' ─────────────────────────────────────────────────────────────────────────────
' 외부에서 접근 가능한 선택된 시트명
' ─────────────────────────────────────────────────────────────────────────────
Public Property Get SelectedSheet() As String
    SelectedSheet = m_selectedSheet
End Property

' ─────────────────────────────────────────────────────────────────────────────
' 시트 목록 로드 (배열을 받아서 ListBox에 채움)
' ─────────────────────────────────────────────────────────────────────────────
Public Sub LoadSheets(ByRef sheetNames() As String)
    Dim i As Integer
    lstSheet.Clear
    For i = 1 To UBound(sheetNames)
        If sheetNames(i) <> "" Then
            lstSheet.AddItem sheetNames(i)
        End If
    Next i

    ' 첫 번째 항목 자동 선택
    If lstSheet.ListCount > 0 Then
        lstSheet.ListIndex = 0
    End If

    ' 현재 저장된 시트가 있으면 해당 항목 선택
    Dim savedSheet As String
    savedSheet = modConfig.GetSourceSheet()
    If savedSheet <> "" Then
        Dim j As Integer
        For j = 0 To lstSheet.ListCount - 1
            If lstSheet.List(j) = savedSheet Then
                lstSheet.ListIndex = j
                Exit For
            End If
        Next j
    End If
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 폼 초기화
' ─────────────────────────────────────────────────────────────────────────────
Private Sub UserForm_Initialize()
    m_selectedSheet = ""
    Me.Caption = "시트 선택"
    lblTitle.Caption = "원본 파일의 시트를 선택하세요:"
    btnOK.Caption = "확인"
    btnCancel.Caption = "취소"
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 확인 버튼
' ─────────────────────────────────────────────────────────────────────────────
Private Sub btnOK_Click()
    If lstSheet.ListIndex < 0 Then
        MsgBox "시트를 선택해주세요.", vbExclamation, "선택 필요"
        Exit Sub
    End If
    m_selectedSheet = lstSheet.List(lstSheet.ListIndex)
    Me.Hide
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 취소 버튼
' ─────────────────────────────────────────────────────────────────────────────
Private Sub btnCancel_Click()
    m_selectedSheet = ""
    Me.Hide
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' ListBox 더블클릭 → 바로 확인
' ─────────────────────────────────────────────────────────────────────────────
Private Sub lstSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstSheet.ListIndex >= 0 Then
        m_selectedSheet = lstSheet.List(lstSheet.ListIndex)
        Me.Hide
    End If
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' 폼 닫기 버튼
' ─────────────────────────────────────────────────────────────────────────────
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        m_selectedSheet = ""
        Cancel = False
    End If
End Sub
