VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
' メニューバークラス
' 2004/10/14
'**

Option Explicit
Option Base 1

Private cbOperation As CommandBar
Const Label = "ガントチャートforExcel"
Private Name As String

'* 既に作成済か否か
Public Function IsCreated() As Boolean
    
    Dim CB As CommandBar
    
    For Each CB In Application.CommandBars
        
        If CB.Name = Name Then
        
            IsCreated = True
            Exit For
            
        End If
        
    Next CB
    
End Function

'* メニュー作成
Public Sub CREATE()
    
    If IsCreated Then
    
        'Debug.Print "MenuBar Already Created."
        Set cbOperation = Application.CommandBars(Name)
        cbOperation.Visible = True
        'Delete
        Exit Sub
        
    End If
    
    Set cbOperation = _
        Application.CommandBars.Add(Name:=Name, MenuBar:=False, Temporary:=True)
    'Set cbOperation = Application.CommandBars("Worksheet Menu Bar")
    
    With cbOperation
        
        .Visible = True
        '.Position = msoBarTop
        .Position = msoBarBottom
        '.Width = 20000
        
    End With
    
    AddButton "新規項目", "mdlMain.ShowNewItemForm"
    AddButton "表示更新", "mdlMain.UpdateChart"
    AddButton "日付変更", "mdlMain.ChangeDate"
    AddButton "休日設定", "mdlMain.EditHoliday"
    AddButton "イナズマ線", "mdlMain.EditInazuma"
    AddButton "マイルストーン", "mdlMain.EditMilestone"
    AddButton "分類並替", "mdlMain.MoveCategory"
    AddButton "項目並替", "mdlMain.showDataManager"
    AddButton "シート出力", "mdlMain.ChartSheetExport"
    AddButton "環境設定", "mdlMain.EditProperty"
    'AddButton "バージョン", ""
    AddButton "終了", "mdlMain.Quit"
    
End Sub

'* メニューの追加
Private Sub AddButton(NewButtonName As String, Action As String)
    
    Dim ctrl As CommandBarButton
    
    Set ctrl = cbOperation.Controls.Add(Type:=msoControlButton)
    
    With ctrl
        
        .Caption = NewButtonName
        .Style = msoButtonCaption
        .OnAction = Action
        
    End With
    
    'Debug.Print "Add menu " & NewButtonName
    
End Sub

Public Sub Hide()
    
    If Not IsCreated Then Exit Sub
        
    If Not IsSet(cbOperation) Then
        Set cbOperation = Application.CommandBars(Name)
    End If
    
    cbOperation.Visible = False
    
End Sub

'* メニュー削除
Public Sub Delete()
    
    If Not IsCreated Then Exit Sub
    
    If Not IsSet(cbOperation) Then
        Set cbOperation = Application.CommandBars(Name)
    End If
    
    cbOperation.Delete
    
    'MsgBox "clsMenuBar.Delete"
    
End Sub

Private Sub Class_Initialize()
    
    Name = Label & " - " & ThisWorkbook.Name
    
End Sub

Private Sub Class_Terminate()
    
    Delete
    Set cbOperation = Nothing
    
End Sub
