VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataManager 
   Caption         =   "UserForm1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   OleObjectBlob   =   "frmDataManager.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DataList As New clsDataList

Private flgResize As Boolean
Private ResizeX As Single
Private ResizeY As Single

'インターフェース設定
Private Sub InitInterFace()
    
    Me.Caption = "項目並替"
    
    btnCancel.Caption = "キャンセル"
    btnUp.Caption = "上へ"
    btnDown.Caption = "下へ"
    btnUpdate.Caption = "工程表更新"
    '8 - fmMousePointerSizeNWSE
    imgResize.MousePointer = fmMousePointerSizeNWSE
    setResizeControls
    
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub setResizeControls()
    
    Const margin = 5
    Const DefaultWidth = 320
    Const DefaultHeight = 240
    
    If Me.Height < DefaultHeight Then Me.Height = DefaultHeight
    If Me.Width < DefaultWidth Then Me.Width = DefaultWidth
    
    With btnCancel
        .Top = Me.InsideHeight - .Height - margin
    End With
    
    With btnUp
        .Top = Me.InsideHeight - .Height - margin
    End With
    
    With btnDown
        .Top = Me.InsideHeight - .Height - margin
    End With
    
    With btnUpdate
        .Top = Me.InsideHeight - .Height - margin
    End With
    
    With imgResize
        .Top = Me.InsideHeight - .Height
        .Left = Me.InsideWidth - .Width
    End With
    
    With lstDatas
        .Left = margin
        .Top = margin
        .Width = Me.InsideWidth - margin * 2
        .Height = btnCancel.Top - margin * 2
    End With
    
End Sub

Private Sub btnDown_Click()
    
    DataList.MoveToDown lstDatas.ListIndex
    DataList.makeListBox lstDatas
    
End Sub

Private Sub btnUp_Click()
    
    'MsgBox lstDatas.ListIndex
    DataList.MoveToUp lstDatas.ListIndex
    DataList.makeListBox lstDatas
    
End Sub

Private Sub btnUpdate_Click()
    
    mdlMain.UpdateChart
    Unload Me
    
End Sub

Private Sub imgResize_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then
        flgResize = True
        ResizeX = x
        ResizeY = y
    End If
End Sub

Private Sub imgResize_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If flgResize Then
        Me.Width = Me.Width + x - ResizeX
        Me.Height = Me.Height + y - ResizeY
        ResizeX = x
        ResizeY = y
    End If
End Sub

Private Sub imgResize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then
        flgResize = False
        setResizeControls
    End If
End Sub

Private Sub UserForm_Initialize()
    
    InitInterFace
    
    DataList.makeListBox lstDatas
    
End Sub
