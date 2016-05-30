VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInazumas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Line(COUNT_INAZUMA) As Shape
Private Inazuma(COUNT_INAZUMA) As New clsInazuma

'イナズマ線追加
Public Sub Add(wkItem As clsItem, wkRow As Long)
    
    Dim i As Integer
    
    For i = 1 To COUNT_INAZUMA
        
        Call Inazuma(i).Add(wkRow, wkItem.Inazuma(i))
        
    Next i
    
End Sub

'描画
Public Sub Draw(wkProperty As clsProperty)
    
    Dim i As Integer
    
    For i = 1 To COUNT_INAZUMA
        
        Call Inazuma(i).Draw(Line(i), wkProperty)
        
    Next i
    
End Sub

'初期化
Public Sub Initialize(wkItems As clsItems)
    
    Dim i As Integer
    
    For i = 1 To COUNT_INAZUMA
    
        Call getLine(i)
        Inazuma(i).TargetDate = wkItems.InazumaDate(i)
    
    Next i
    
End Sub

'イナズマ線図形取得
Private Sub getLine(Num As Integer)
    
    On Error GoTo CREATE
    
    Dim wk As Shape
    Dim Name As String
    
    Name = "Inazuma" & Format(Num, "00")
    
    Set Line(Num) = ActiveSheet.Shapes(Name)
    Exit Sub
    
CREATE:
    
    Set Line(Num) = CreateInazumaObject(Name)
    
End Sub

'新規イナズマ線図形
Private Property Get CreateInazumaObject(Name As String) As Shape
    
    With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, 0#, 0#)
        .AddNodes msoSegmentLine, msoEditingAuto, 100#, 100#
        .ConvertToShape.Select
    End With
    
    Selection.Name = Name
    
    Set CreateInazumaObject = ActiveSheet.Shapes(Name)
    
End Property
