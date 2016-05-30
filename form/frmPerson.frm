VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerson 
   Caption         =   "Person"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535.001
   OleObjectBlob   =   "frmPerson.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'担当者編集用フォーム
Option Explicit
Option Base 1

Private psn As New clsPersons   '担当データ
Private flgEdit As Boolean      '編集フラグ

Private Const NEW_ITEM = "新規登録"
Private Const MSG_NONAME = "名前が入力されていません"
Private Const MSG_ISSETCHANGE = "変更内容設定しますか?"

'インターフェース設定
Private Sub InitInterFace()
    
    Label1.Caption = "氏名"
    Label2.Caption = "電話1"
    Label3.Caption = "電話2"
    Label4.Caption = "FAX"
    Label5.Caption = "E-Mail"
    Label6.Caption = "備考"
    
    btnSet.Caption = "設定"
    btnSave.Caption = "保存"
    btnCancel.Caption = "取消"
    
    frmDetail.Caption = "詳細"
    
    Me.Caption = "担当者"
    
End Sub

'詳細転写
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

'フォームから新規担当者データを生成
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

'一覧更新
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

'設定
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

'保存
Private Sub btnSave_Click()
    
    If flgEdit Then
    
        If MsgBox(MSG_ISSETCHANGE, vbYesNo) = vbYes Then
            Call btnSet_Click
        End If
        
    End If
    
    psn.Save
    Me.Hide
    
End Sub

'取消
Private Sub btnCancel_Click()
    Me.Hide
End Sub

'担当者選択
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

'編集の判定
Private Sub txtName_Change(): flgEdit = True: End Sub
Private Sub txtPhone1_Change(): flgEdit = True: End Sub
Private Sub txtPhone2_Change(): flgEdit = True: End Sub
Private Sub txtFax_Change(): flgEdit = True: End Sub
Private Sub txtMail_Change(): flgEdit = True: End Sub
Private Sub txtNote_Change(): flgEdit = True: End Sub


'初期化
Private Sub UserForm_Initialize()

    Call InitInterFace
    Call UpdateList
    
    Call mdlTools.setFontOnForm(Me)

End Sub
