VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'担当者クラス
Option Explicit
Option Base 1

Private itm() As clsPerson

'登録済み担当者数
Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    i = UBound(itm)
    
    Count = i
    
End Property

'担当者取得
Public Property Get Items(Number As Long) As clsPerson
    
    On Error Resume Next
    
    Set Items = itm(Number)
    
End Property

'担当者設定
Public Property Let Items(Number As Long, newData As clsPerson)
    
    On Error Resume Next
    
    Set itm(Number) = newData
    
End Property

'保存
Public Sub Save()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        itm(i).Save
        
    Next i
    
End Sub

'読込み
Public Sub Load()
    
    Dim buf As clsPerson
    Dim i As Long
    i = 2
    
    Do
        
        Set buf = New clsPerson
        
        If buf.Load(i) Then
            
            Add
            Set itm(Count) = buf
            
        Else
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
End Sub

'追加
Public Function Add() As clsPerson
    
    Dim i As Long
    i = Count() + 1
    
    ReDim Preserve itm(i)
    'Set itm(i) = New clsPerson
    
    Set Add = itm(i)
    
End Function

Private Sub Class_Initialize()
    Call Load
End Sub
