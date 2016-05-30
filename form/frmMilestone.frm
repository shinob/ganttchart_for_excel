VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMilestone 
   Caption         =   "Milestone"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895.001
   OleObjectBlob   =   "frmMilestone.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

Private Const MSG_UPDATE = "内容が変更されています。" & vbCr & _
    "変更を反映しますか?"
Private Const MSG_MISS = "設定内容が不十分です。"
Private Const ADD_NEW = "[新規マイルストーン]"

'インターフェース設定
Private Sub InitInterFace()
    
    btnSet.Caption = "設定"
    btnDelete.Caption = "削除"
    btnSave.Caption = "保存"
    btnCancel.Caption = "取消"
    
    frmEdit.Caption = "編集"
    frmEdit.Visible = False
    
    Me.Caption = "マイルストーン "
    
End Sub

'リスト更新
Private Sub UpdateList()
    
    Call Milestones.setControl(lstMilestone)
    lstMilestone.AddItem ADD_NEW
    
End Sub

'データをフォームに転写
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

'日付データか否か
Private Function isDate(wk As String) As Boolean
    
    On Error GoTo ERR
    
    Dim D As Date
    
    D = CDate(wk)
    isDate = True
    Exit Function
    
ERR:
    isDate = False
    
End Function

'フォームからデータを取得
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

'取消
Private Sub btnCancel_Click()
    
    Me.Hide
    
End Sub

'選択項目の取得
Private Function getSelected() As Long
    
    Dim i As Long
    
    i = lstMilestone.ListIndex + 1
    If i < 1 Or Milestones.Count() < i Then Exit Function
    
    getSelected = i
    
End Function

'削除
Private Sub btnDelete_Click()
    
    Dim i As Long
    i = getSelected()
    
    If 0 < i Then
        
        Milestones.Milestones(i).flgDelete = True
        Call UpdateList
        
    End If
    
    frmEdit.Visible = False
    
End Sub

'保存
Private Sub btnSave_Click()
    
    If flgChange Then
    
        If vbYes = MsgBox(MSG_UPDATE, vbYesNo) Then _
            Call GetDataFromForm(Active)
        
    End If
    
    Call Milestones.Save
    
    flgEdit = True
    
    Me.Hide
    
End Sub

'設定
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

'日付選択
Private Sub lblTargetDate_Click()
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = "マイルストーン"
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

'編集データ選択
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
