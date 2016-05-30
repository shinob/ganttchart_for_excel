VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInazuma 
   Caption         =   "Status Line Configuration"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   OleObjectBlob   =   "frmInazuma.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmInazuma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Base 1

Private Items As New clsItems
Public flgEdit As Boolean

Private Const CALENDAR = "����I��"

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()
    
    btnSave.Caption = "�ݒ�"
    btnExit.Caption = "���"
    btnDelete.Caption = "�폜"
    
    Me.Caption = "�C�i�Y�}��"
    
End Sub

'�Ǎ�
Private Sub Load()

    Dim Schedules As New clsSchedules
    
    '���ړǍ�
    Call Items.Load
    
    '�H���Ǎ�
    Call Schedules.Load(-1)
    
    '���ڂ֍H����ݒ�
    Dim i As Long
    Dim j As Long
    For i = 1 To Schedules.Count()
        
        j = Schedules.Items(i).Item
        If j < 2 Then
        
        Else
            With Items.Items(j - 1).Schedules
                .Add
                .Items(.Count()) = Schedules.Items(i)
            End With
        End If
        
    Next i
    
End Sub

'�ꗗ�̍X�V
Private Sub UpdateList()
    
    Dim i As Integer
    Dim s As String
    
    lstInazuma.Clear
    
    For i = 1 To COUNT_INAZUMA
        
        s = Format(i, "00") & " : "
        If FIRSTDATE < Items.InazumaDate(i) Then s = s & Format(Items.InazumaDate(i), "yyyy/mm/dd")
        lstInazuma.AddItem s
        
    Next i
    
End Sub

'�C�i�Y�}���̐ݒ�
Private Sub setInazuma(Num As Integer)
    
    Dim frm As New frmCalendar
    Dim wkDate As Date
    
    wkDate = Items.InazumaDate(Num)
    If wkDate < FIRSTDATE Then wkDate = Now()
    
    With frm
        
        .Caption = CALENDAR
        .mode = "SELECT"
        Call .SetDate(wkDate)
        .Show
        
        If .flgSelect Then
            'Items.InazumaDate(num) = .SelectedDate
            Call Items.setInazuma(Num, .SelectedDate)
        End If
        
    End With
    
    Call UpdateList
    
End Sub

'�C�i�Y�}���폜
Private Sub deleteInazuma(Num As Integer)
    
    Dim i As Long
    
    With Items
        
        .InazumaDate(Num) = 0
        
        For i = 1 To .Count()
        
            .Items(i).Inazuma(Num) = 0
            
        Next i
        
    End With
    
    Call UpdateList
    
End Sub

'�폜
Private Sub btnDelete_Click()
    
    If lstInazuma.ListIndex < 0 Then Exit Sub
    Call deleteInazuma(lstInazuma.ListIndex + 1)
    
End Sub

'���
Private Sub btnExit_Click()
    
    Me.Hide
    
End Sub

'�ۑ�
Private Sub btnSave_Click()
    
    Items.SaveInazuma
    flgEdit = True
    Me.Hide
    
End Sub

'�ҏW
Private Sub lstInazuma_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    If lstInazuma.ListIndex < 0 Then Exit Sub
    Call setInazuma(lstInazuma.ListIndex + 1)
    
End Sub

Private Sub UserForm_Initialize()
        
    Call Load
    Call UpdateList
        
    Call InitInterFace
    Call mdlTools.setFontOnForm(Me)
    
End Sub
