VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmItemSelect.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmItemSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Base 1

Private obj As New clsCategorys
Private Item As clsItems
Private Const STR_INDENT = "  "
Private Const MSG_NG = "���ڂ��I������Ă��܂���"

Public mode As Integer
Public SelectedItem As Long
Public ActiveItem As Long

Private indCnt As Long
Private ind() As Long

Private Sub setInterFace()
    
    btnSet.Caption = "�ݒ�"
    btnCancel.Caption = "���"
    btnClear.Caption = "����"
    
    Me.Caption = "���ڑI��"
    
End Sub

'���ށA���ړǍ�
Private Sub Load()
    
    Dim wk As New clsItems
    
    wk.Load
    
    Dim i As Long
    Dim j As Long
    Dim itm As clsItem
    
    For i = 1 To wk.Count()
        
        Set itm = wk.Items(i)
        
        j = itm.LinkItem
        If 1 < j Then
            With wk.Items(j - 1).SubItems
                .Add
                .Items(.Count()) = itm
            End With
        End If
        
        j = itm.Category
        If 1 < j Then
            With obj.Items(j - 1).Items
                .Add
                .Items(.Count()) = itm
            End With
        End If
        
    Next i
    
    Set Item = wk
    indCnt = 1
    ReDim ind(wk.Count() + obj.Count())
    
End Sub

'���X�g�X�V
Private Sub UpdateList(lst As Control)
    
    Dim i As Long
    
    For i = 1 To obj.Count()
        
        lst.AddItem obj.Items(i).Name
        
        ind(indCnt) = 0
        indCnt = indCnt + 1
        
        Call AddItemOnList(obj.Items(i).Items, 1, lst)
        
    Next i
    
End Sub

'���X�g�֍��ڒǉ�
Private Sub AddItemOnList(wkItems As clsItems, Deps As Integer, lst As Control)
    
    Dim i As Long
    
    For i = 1 To wkItems.Count()
        
        lst.AddItem indent(Deps) & wkItems.Items(i).Name
        
        ind(indCnt) = wkItems.Items(i).No
        indCnt = indCnt + 1
        
        Call AddItemOnList(wkItems.Items(i).SubItems, Deps + 1, lst)
        
    Next i
    
End Sub

'�C���f���g�擾
Private Function indent(Deps As Integer) As String
    
    Dim i As Integer
    indent = STR_INDENT
    For i = 1 To Deps
        indent = indent & STR_INDENT
    Next i
    
End Function

'�L�����Z��
Private Sub btnCancel_Click()
    mode = ITEMSELECT_CANCEL
    Me.Hide
End Sub

'�ݒ����
Private Sub btnClear_Click()
    mode = ITEMSELECT_UNLINK
    Me.Hide
End Sub

'�ݒ�
Private Sub btnSet_Click()
    
    If SelectedItem < 2 Then
        MsgBox MSG_NG
        Exit Sub
    ElseIf SelectedItem = ActiveItem Then
        MsgBox "�������g����ʍ��ڂɐݒ�͂ł��܂���"
        Exit Sub
    ElseIf CheckLink() Then
        MsgBox "�����̉��ʍ��ڂ���ʍ��ڂɐݒ�͂ł��܂���B"
        Exit Sub
    End If
    
    mode = ITEMSELECT_LINK
    Me.Hide
End Sub

'�����N���K�����ۂ�
Private Function CheckLink() As Boolean
    
    Dim i As Long
    Dim flg  As Boolean
    
    Dim wk As clsItem
    Set wk = Item.Items(SelectedItem - 1)
    
    'MsgBox wk.No & " " & wk.Name & " " & wk.LinkItem
        
    Do Until wk.LinkItem = 0
        
        If wk.LinkItem = ActiveItem Then
            flg = True
            Exit Do
        Else
            Set wk = Item.Items(wk.LinkItem - 1)
        End If
        
    Loop
    
    CheckLink = flg
    
End Function

'�I��
Private Sub lstItem_Click()
    
    Dim i As Long
    i = lstItem.ListIndex + 1
    If i = 0 Then Exit Sub
    
    SelectedItem = ind(i)
    
End Sub

Private Sub UserForm_Initialize()
    
    Call setInterFace
    Call mdlTools.setFontOnForm(Me)
    
    Call Load
    Call UpdateList(lstItem)
    
End Sub
