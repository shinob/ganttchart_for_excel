VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMemos 
   Caption         =   "Memos"
   ClientHeight    =   3500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   OleObjectBlob   =   "frmMemos.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private wkMemo As clsMemos

Private Sub InitInterFace()
    
    Me.Caption = "備考編集"
    
    btnSet.Caption = "設定"
    btnCancel.Caption = "取消"
    
End Sub

Public Sub Load(ByRef Memo As clsMemos)
    
    Set wkMemo = Memo
    Call SetValues
    
End Sub

Private Sub SetValues()
    
    Dim i As Integer
    
    For i = 1 To COUNT_MEMO
        
        Me.Controls("TextBox" & i).Text = wkMemo.Text(i)
        
    Next i
    
End Sub

Private Sub GetValues()
    
    Dim i As Integer
    
    For i = 1 To COUNT_MEMO
        
        wkMemo.Text(i) = Me.Controls("TextBox" & i).Text
        
    Next i
    
End Sub

Private Sub btnCancel_Click()
    
    Me.Hide
    
End Sub

Private Sub btnSet_Click()
    
    Call GetValues
    Me.Hide
    
End Sub

'初期化
Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call mdlTools.setFontOnForm(Me)

End Sub
