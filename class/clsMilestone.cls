VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMilestone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�}�C���X�g�[���p�N���X
Option Explicit
Option Base 1

Public No As Long
Public TargetDate As Date
Public Name As String

Public flgDelete As Boolean

Private sht As Worksheet

'�Ǎ�
Public Function Load(Num As Long) As Boolean
    
    If Num < 2 Then Exit Function
    
    With sht
        
        No = .Cells(Num, 1).Value
        TargetDate = .Cells(Num, 2).Value
        Name = .Cells(Num, 3).Value
        
    End With
    
    If No < 2 Then Exit Function
    
    Load = True
    
End Function

'�ۑ�
Public Function Save(Num As Long) As Boolean
    
    If flgDelete Then
        
        Call Delete
        Exit Function
    
    End If
    
    If Num < 2 Then
        No = mdlTools.FindBlankRow(sht, 2, 2)
    Else
        No = Num
    End If
    
    With sht
        
        .Cells(No, 1).Value = No
        .Cells(No, 2).Value = TargetDate
        .Cells(No, 3).Value = Name
        
        .Cells(No + 1, 1).Value = ""
        .Cells(No + 1, 2).Value = ""
        .Cells(No + 1, 3).Value = ""
        
    End With
    
    Save = True
    
End Function

'�폜
Public Function Delete() As Boolean
    
    If No < 2 Then Exit Function
    
    With sht
        
        .Cells(No, 1).Value = No
        .Cells(No, 2).Value = ""
        .Cells(No, 3).Value = ""
        
    End With
    
    Delete = True
    
End Function

'�`��̗v�E�s�v
Public Function isDraw(BeginDate As Date, EndDate As Date)
    
    If BeginDate <= TargetDate And TargetDate < EndDate Then
        isDraw = True
    Else
        isDraw = False
    End If
    
End Function

'���X�g�p����
Public Property Get Text() As String
    
    Dim msg As String
    Dim D As Date
    
    D = TargetDate
    
    If D <= FIRSTDATE Then
        
        msg = "          "
        
    Else
        
        msg = Format(TargetDate, "yyyy/mm/dd")
        
    End If
    
    msg = msg & " " & Name
    
    If flgDelete Then msg = msg & " [�폜]"
    
    Text = msg
    
End Property

Private Sub Class_Initialize()
    
    Set sht = shtMilestone
    
End Sub