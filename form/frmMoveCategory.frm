VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveCategory 
   Caption         =   "UserForm1"
   ClientHeight    =   4460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   OleObjectBlob   =   "frmMoveCategory.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMoveCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Base 1

Private ctg As New clsCategorys
Private itm As New clsItems

Public flgEdit As Boolean

'インターフェース設定
Private Sub InitInterFace()

    btnUp.Caption = "上へ"
    btnDown.Caption = "下へ"
    
    btnSave.Caption = "保存"
    btnExit.Caption = "取消"
    
    btnDelete.Caption = "削除"
    
    Me.Caption = "分類並替"
    
End Sub

'分類へ項目を設定
Private Sub MargeDataByCategory()
    
    Dim i As Long
    Dim j As Long
    
    Call itm.Load
    
    '分類へ項目を設定
    For i = 1 To itm.Count()
        
        j = itm.Items(i).LinkItem
        If j < 2 Then
        Else
            With itm.Items(j - 1).SubItems
                .Add
                .Items(.Count()) = itm.Items(i)
            End With
        End If
        
        j = itm.Items(i).Category
        If j < 2 Then
        
        Else
            With ctg.Items(j - 1).Items
                .Add
                .Items(.Count()) = itm.Items(i)
            End With
        End If
        
    Next i
    
End Sub

'リスト更新
Private Sub UpdateList()
    
    Dim i As Long
    Dim tmp As String
    Dim wk As clsCategory
    
    lstCategory.Clear
    
    For i = 1 To ctg.Count
        
        Set wk = ctg.Items(i)
        'lstCategory.AddItem wk.No & " : " & wk.Name & " : " & wk.Items.Count()
        
        If wk.Delete Then
            tmp = " [削除]"
        Else
            tmp = ""
        End If
        
        lstCategory.AddItem wk.No & " : " & wk.Name & tmp
        'lstCategory.AddItem wk.NO & " : " & wk.Name & tmp & wk.items.Count
        
    Next i
    
End Sub

'上へ移動
Private Sub MoveToUp()
    
    Dim wk As clsCategory
    Dim i As Long
    
    i = lstCategory.ListIndex
    If i < 1 Then Exit Sub
    
    If ctg.Items(i + 1).Delete Then Exit Sub
    
    Call ctg.Exchange(i, i + 1)
    
    Call UpdateList
    lstCategory.ListIndex = i - 1
    
End Sub

'下へ移動
Private Sub MoveToDown()
    
    Dim wk As clsCategory
    Dim i As Long
    
    i = lstCategory.ListIndex + 1
    If i < 1 Or i = ctg.Count() Then Exit Sub
    
    If ctg.Items(i).Delete Or ctg.Items(i + 1).Delete Then Exit Sub
    Call ctg.Exchange(i, i + 1)
    
    Call UpdateList
    lstCategory.ListIndex = i
    
End Sub

Private Sub btnDelete_Click()
    
    Dim wk As clsCategory
    
    Dim i As Long
    Dim j As Long
    
    i = lstCategory.ListIndex + 1
    If i < 1 Then Exit Sub
    
    If ctg.CountNotDeleted = 1 Then
        MsgBox "最低1つの分類が登録されている必要があります"
        Exit Sub
    End If
    
    Set wk = ctg.Items(i)
    
    If wk.Items.Count > 0 Then
        MsgBox "項目の登録された分類は削除できません"
        Exit Sub
'    Else
'        wk.Items.Load_By_Category wk.NO
'        If wk.Items.Count > 0 Then Exit Sub
    End If
    
    If wk.Delete Then
        Exit Sub
    Else
        wk.Delete = True
    End If
    
    For j = i To ctg.Count - 1
        
        Call ctg.Exchange(j, j + 1)
        
    Next j
    
    Call UpdateList
    lstCategory.ListIndex = i - 1
    
End Sub

'下
Private Sub btnDown_Click()
    Call MoveToDown
End Sub

'取消
Private Sub btnExit_Click()
    Me.Hide
End Sub

'保存
Private Sub btnSave_Click()
    
    ctg.Save
    itm.SaveCateogory
    
    flgEdit = True
    Me.Hide
    
End Sub

'上
Private Sub btnUp_Click()
    Call MoveToUp
End Sub

'初期化
Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call mdlTools.setFontOnForm(Me)
    
    Call MargeDataByCategory
    Call UpdateList

End Sub
