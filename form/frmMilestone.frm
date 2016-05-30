VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMilestone 
   Caption         =   "Milestone"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895.001
   OleObjectBlob   =   "frmMilestone.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmMilestone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Milestones As New clsMilestones

Public flgEdit As Boolean

Private Active As Long
Private flgChange As Boolean

Private Const MSG_UPDATE = "���e���ύX����Ă��܂��B" & vbCr & _
    "�ύX�𔽉f���܂���?"
Private Const MSG_MISS = "�ݒ���e���s�\���ł��B"
Private Const ADD_NEW = "[�V�K�}�C���X�g�[��]"

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()
    
    btnSet.Caption = "�ݒ�"
    btnDelete.Caption = "�폜"
    btnSave.Caption = "�ۑ�"
    btnCancel.Caption = "���"
    
    frmEdit.Caption = "�ҏW"
    frmEdit.Visible = False
    
    Me.Caption = "�}�C���X�g�[�� "
    
End Sub

'���X�g�X�V
Private Sub UpdateList()
    
    Call Milestones.setControl(lstMilestone)
    lstMilestone.AddItem ADD_NEW
    
End Sub

'�f�[�^���t�H�[���ɓ]��
Private Sub SetDataOnForm(Num As Long)
    
    Dim wk As clsMilestone
    
    If Num < 1 Then
        Set wk = New clsMilestone
    Else
        Set wk = Milestones.Milestones(Num)
    End If
    
    With wk
        
        lblNo.Caption = .No
        
        Dim D As String
        If FIRSTDATE <= .TargetDate Then D = Format(.TargetDate, "yyyy/mm/dd")
        lblTargetDate.Caption = D
        
        txtName.Text = .Name
        
    End With
    
    Set wk = Nothing
    
    flgChange = False
    
End Sub

'���t�f�[�^���ۂ�
Private Function isDate(wk As String) As Boolean
    
    On Error GoTo ERR
    
    Dim D As Date
    
    D = CDate(wk)
    isDate = True
    Exit Function
    
ERR:
    isDate = False
    
End Function

'�t�H�[������f�[�^���擾
Private Function GetDataFromForm(Num As Long) As Boolean
    
    If Not isDate(lblTargetDate.Caption) Then Exit Function
    
    If CDate(lblTargetDate.Caption) < FIRSTDATE Or _
        txtName.Text = "" Then Exit Function
    
    Dim wk As clsMilestone
    
    If Num < 1 Then
        
        Set wk = New clsMilestone
        
    Else
    
        Set wk = Milestones.Milestones(Num)
        
    End If
    
    With wk
        
        .No = lblNo.Caption
        .TargetDate = CDate(lblTargetDate.Caption)
        .Name = txtName.Text
        
        .flgDelete = False
        
    End With
    
    If Num < 1 Then Call Milestones.Add(wk)
    
    GetDataFromForm = True
    
End Function

'���
Private Sub btnCancel_Click()
    
    Me.Hide
    
End Sub

'�I�����ڂ̎擾
Private Function getSelected() As Long
    
    Dim i As Long
    
    i = lstMilestone.ListIndex + 1
    If i < 1 Or Milestones.Count() < i Then Exit Function
    
    getSelected = i
    
End Function

'�폜
Private Sub btnDelete_Click()
    
    Dim i As Long
    i = getSelected()
    
    If 0 < i Then
        
        Milestones.Milestones(i).flgDelete = True
        Call UpdateList
        
    End If
    
    frmEdit.Visible = False
    
End Sub

'�ۑ�
Private Sub btnSave_Click()
    
    If flgChange Then
    
        If vbYes = MsgBox(MSG_UPDATE, vbYesNo) Then _
            Call GetDataFromForm(Active)
        
    End If
    
    Call Milestones.Save
    
    flgEdit = True
    
    Me.Hide
    
End Sub

'�ݒ�
Private Sub btnSet_Click()
    
    Dim flg As Boolean
    
    flgChange = False
    flg = GetDataFromForm(Active)
    Call UpdateList
    
    If flg Then
    
        frmEdit.Visible = False
        
    Else
    
        MsgBox MSG_MISS
        
    End If
    
End Sub

'���t�I��
Private Sub lblTargetDate_Click()
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = "�}�C���X�g�[��"
        .mode = "SELECT"
        
        If lblTargetDate.Caption <> "" Then
            Call .SetDate(CDate(lblTargetDate.Caption))
        Else
            Call .SetDate(Now())
        End If
        
        .Show
        
        If .flgSelect Then
            
            lblTargetDate.Caption = Format(.SelectedDate, "yyyy/mm/dd")
            flgChange = True
            
        End If
        
    End With
    
End Sub

'�ҏW�f�[�^�I��
Private Sub lstMilestone_Click()
    
    Dim i As Long
    i = getSelected()
    
    If flgChange Then
        
        If vbYes = MsgBox(MSG_UPDATE, vbYesNo) Then
            Call GetDataFromForm(Active)
            Call UpdateList
            lstMilestone.ListIndex = i
        End If
        
    End If
    
    Active = i
    Call SetDataOnForm(Active)
    frmEdit.Visible = True
    
End Sub

Private Sub txtName_Change()
    
    flgChange = True
    
End Sub

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call mdlTools.setFontOnForm(Me)
    
    Call Milestones.Load
    Call UpdateList
    
End Sub
