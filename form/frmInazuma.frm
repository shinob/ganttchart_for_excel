VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInazuma 
   Caption         =   "Status Line Configuration"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   OleObjectBlob   =   "frmInazuma.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

Private Const CALENDAR = "基準日選択"

'インターフェース設定
Private Sub InitInterFace()
    
    btnSave.Caption = "設定"
    btnExit.Caption = "取消"
    btnDelete.Caption = "削除"
    
    Me.Caption = "イナズマ線"
    
End Sub

'読込
Private Sub Load()

    Dim Schedules As New clsSchedules
    
    '項目読込
    Call Items.Load
    
    '工程読込
    Call Schedules.Load(-1)
    
    '項目へ工程を設定
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

'一覧の更新
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

'イナズマ線の設定
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

'イナズマ線削除
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

'削除
Private Sub btnDelete_Click()
    
    If lstInazuma.ListIndex < 0 Then Exit Sub
    Call deleteInazuma(lstInazuma.ListIndex + 1)
    
End Sub

'取消
Private Sub btnExit_Click()
    
    Me.Hide
    
End Sub

'保存
Private Sub btnSave_Click()
    
    Items.SaveInazuma
    flgEdit = True
    Me.Hide
    
End Sub

'編集
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
