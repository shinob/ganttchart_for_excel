VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerson 
   Caption         =   "Person"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535.001
   OleObjectBlob   =   "frmPerson.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'�S���ҕҏW�p�t�H�[��
Option Explicit
Option Base 1

Private psn As New clsPersons   '�S���f�[�^
Private flgEdit As Boolean      '�ҏW�t���O

Private Const NEW_ITEM = "�V�K�o�^"
Private Const MSG_NONAME = "���O�����͂���Ă��܂���"
Private Const MSG_ISSETCHANGE = "�ύX���e�ݒ肵�܂���?"

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()
    
    Label1.Caption = "����"
    Label2.Caption = "�d�b1"
    Label3.Caption = "�d�b2"
    Label4.Caption = "FAX"
    Label5.Caption = "E-Mail"
    Label6.Caption = "���l"
    
    btnSet.Caption = "�ݒ�"
    btnSave.Caption = "�ۑ�"
    btnCancel.Caption = "���"
    
    frmDetail.Caption = "�ڍ�"
    
    Me.Caption = "�S����"
    
End Sub

'�ڍד]��
Private Sub SetDetailData(Person As clsPerson)
    
    With Person
        
        txtNo.Text = .No
        txtName.Text = .Name
        txtPhone1.Text = .Phone1
        txtPhone2.Text = .Phone2
        txtFax.Text = .FAX
        txtMail.Text = .EMail
        txtNote.Text = .Note
        
    End With
    
End Sub

'�t�H�[������V�K�S���҃f�[�^�𐶐�
Private Function CREATE() As clsPerson
    
    Dim tmp As New clsPerson
    
    With tmp
        
        .No = val(txtNo.Text)
        .Name = txtName.Text
        .Phone1 = txtPhone1.Text
        .Phone2 = txtPhone2.Text
        .FAX = txtFax.Text
        .EMail = txtMail.Text
        .Note = txtNote.Text
        
    End With
    
    Set CREATE = tmp
    
End Function

'�ꗗ�X�V
Private Sub UpdateList()
    
    With lstPerson
        
        .Clear
        
        .AddItem NEW_ITEM
        
        Dim i As Long
        
        For i = 1 To psn.Count()
            
            .AddItem psn.Items(i).Name
            
        Next i
        
    End With
    
    flgEdit = False
    
End Sub

'�ݒ�
Private Sub btnSet_Click()
    
    If txtName.Text = "" Then
        MsgBox MSG_NONAME
        Exit Sub
    End If
        
    Dim buf As clsPerson
    
    Set buf = CREATE()
    
    If buf.No < 2 Then
    
        buf.No = psn.Count() + 2
        txtNo.Text = buf.No
        
        psn.Add
        psn.Items(psn.Count()) = buf
        
    Else
        
        psn.Items(buf.No - 1) = Nothing
        psn.Items(buf.No - 1) = buf
        
    End If
    
    UpdateList
    
End Sub

'�ۑ�
Private Sub btnSave_Click()
    
    If flgEdit Then
    
        If MsgBox(MSG_ISSETCHANGE, vbYesNo) = vbYes Then
            Call btnSet_Click
        End If
        
    End If
    
    psn.Save
    Me.Hide
    
End Sub

'���
Private Sub btnCancel_Click()
    Me.Hide
End Sub

'�S���ґI��
Private Sub lstPerson_Click()
    
    Dim i As Long
    i = lstPerson.ListIndex
    
    If flgEdit Then
        
        If MsgBox(MSG_ISSETCHANGE, vbYesNo) = vbYes Then
            Call btnSet_Click
            lstPerson.ListIndex = i
        End If
        
    End If
    
    If i < 1 Then
        
        Dim buf As New clsPerson
        SetDetailData buf
        
    Else
        
        SetDetailData psn.Items(i)
        
    End If
    
    flgEdit = False
    
End Sub

'�ҏW�̔���
Private Sub txtName_Change(): flgEdit = True: End Sub
Private Sub txtPhone1_Change(): flgEdit = True: End Sub
Private Sub txtPhone2_Change(): flgEdit = True: End Sub
Private Sub txtFax_Change(): flgEdit = True: End Sub
Private Sub txtMail_Change(): flgEdit = True: End Sub
Private Sub txtNote_Change(): flgEdit = True: End Sub


'������
Private Sub UserForm_Initialize()

    Call InitInterFace
    Call UpdateList
    
    Call mdlTools.setFontOnForm(Me)

End Sub
