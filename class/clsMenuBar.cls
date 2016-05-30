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
' ���j���[�o�[�N���X
' 2004/10/14
'**

Option Explicit
Option Base 1

Private cbOperation As CommandBar
Const Label = "�K���g�`���[�gforExcel"
Private Name As String

'* ���ɍ쐬�ς��ۂ�
Public Function IsCreated() As Boolean
    
    Dim CB As CommandBar
    
    For Each CB In Application.CommandBars
        
        If CB.Name = Name Then
        
            IsCreated = True
            Exit For
            
        End If
        
    Next CB
    
End Function

'* ���j���[�쐬
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
    
    AddButton "�V�K����", "mdlMain.ShowNewItemForm"
    AddButton "�\���X�V", "mdlMain.UpdateChart"
    AddButton "���t�ύX", "mdlMain.ChangeDate"
    AddButton "�x���ݒ�", "mdlMain.EditHoliday"
    AddButton "�C�i�Y�}��", "mdlMain.EditInazuma"
    AddButton "�}�C���X�g�[��", "mdlMain.EditMilestone"
    AddButton "���ޕ���", "mdlMain.MoveCategory"
    AddButton "���ڕ���", "mdlMain.showDataManager"
    AddButton "�V�[�g�o��", "mdlMain.ChartSheetExport"
    AddButton "���ݒ�", "mdlMain.EditProperty"
    'AddButton "�o�[�W����", ""
    AddButton "�I��", "mdlMain.Quit"
    
End Sub

'* ���j���[�̒ǉ�
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

'* ���j���[�폜
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