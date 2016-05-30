VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'’S“–ŽÒƒNƒ‰ƒX
Option Explicit
Option Base 1

Public No As Long
Public Name As String
Public Phone1 As String
Public Phone2 As String
Public FAX As String
Public EMail As String
Public Note As String

Private sht As Worksheet

Public Items As New clsItems

'•Û‘¶
Public Function Save() As Boolean
    
    If No < 2 Then
        No = FindBlankRow(sht, 2, 2)
    End If
    
    Dim strg As New clsStrage
    
    With strg
        
        Call .Initialize(sht, True, No, 1)
        
        .Data = No
        .Data = Name
        .Data = Phone1
        .Data = Phone2
        .Data = FAX
        .Data = EMail
        .Data = Note
        
    End With
    
    Save = True
    
End Function

'“Çž
Public Function Load(Row As Long) As Boolean
    
    If sht.Cells(Row, 1) = "" Then
        Load = False
        Exit Function
    End If
    
    Dim strg As New clsStrage
    
    With strg
        
        Call .Initialize(sht, True, Row, 1)
        
        No = .Data
        Name = .Data
        Phone1 = .Data
        Phone2 = .Data
        FAX = .Data
        EMail = .Data
        Note = .Data
        
    End With
    
    Load = True
    
End Function

Private Sub Class_Initialize()
    Set sht = shtPerson
End Sub
