VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�x���Ǘ�
Option Explicit
Option Base 1

'Private Const FIRSTDATE = #1/1/2000#
Private sht As Worksheet

'�J�����_�[�J�n���̎擾
Public Property Get getFirstDate() As Date
    
    getFirstDate = FIRSTDATE
    
End Property

'�x���̊m�F
Public Property Get isHoliday(TargetDate As Date) As Boolean
    
    isHoliday = TargetCell(TargetDate).Value
    
End Property

'�x���̐ݒ�
Public Property Let setHoliday(TargetDate As Date, Value As Boolean)
    
    If Value Then
        TargetCell(TargetDate).Value = Value
    Else
        TargetCell(TargetDate).Clear
    End If
    
End Property

'�Ώۓ��̃Z�����擾
Private Function TargetCell(TargetDate As Date) As Range
    
    Dim y As Integer
    Dim m As Integer
    Dim D As Integer
    
    y = Year(TargetDate) - Year(FIRSTDATE)
    m = Month(TargetDate)
    D = Day(TargetDate)
    
    Set TargetCell = sht.Cells(y * 12 + m, Day(TargetDate))
    
End Function

'������
Private Sub Class_Initialize()
    Set sht = shtHoliday
End Sub

'
Private Sub Class_Terminate()
    Set sht = Nothing
End Sub