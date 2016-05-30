VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ストレージ用クラス
Option Explicit
Option Base 1

Private sht As Worksheet
Private Direction As Boolean
    'Row : True / Column : False
Private No As Long
Private Begin As Long

Private Position As Long

Public Sub Initialize(DataSheet As Worksheet, DirectionType As Boolean, _
    Num As Long, Start As Long)
    
    Set sht = DataSheet
    Direction = DirectionType
    No = Num
    
    Begin = Start
    Position = Start
    
End Sub

Public Property Let Column(Num As Integer)
    
    No = Num
    
End Property

Public Property Get getPosition() As Long
    
    getPosition = Position
    
End Property

Public Function Data() As Range
    
    Dim col As Long
    Dim Row As Long
    
    If Direction Then
        col = No
        Row = Position
    Else
        col = Position
        Row = No
    End If
    
    Call Skip
    
    Set Data = sht.Cells(col, Row)
    
End Function

Public Sub Skip()
    
    Position = Position + 1
    
End Sub
