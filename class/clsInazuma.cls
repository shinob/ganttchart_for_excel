VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInazuma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'イナズマ線用クラス
Option Explicit
Option Base 1

Private Points() As clsInazumaPoint
Public TargetDate As Date

'追加
Public Sub Add(Row As Long, Value As Date)
    
    Dim i As Long
    
    i = Count() + 1
    
    ReDim Preserve Points(i)
    Set Points(i) = New clsInazumaPoint
    
    With Points(i)
        
        .Row = Row
        .Value = Value
        
    End With
    
End Sub

'個数
Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    i = UBound(Points)
    
    Count = i
    
End Property

'描画
Public Sub Draw(Line As Shape, Property As clsProperty)
    
    Dim i As Long
    Dim Nodes As Long
    
    If TargetDate < Property.BeginDate Or _
        Property.EndDate < TargetDate Or _
        Not Property.InazumaDraw Then
        
        Line.Visible = msoFalse
        Exit Sub
        
    Else
        
        Line.Visible = msoTrue
        
    End If
    
    'MsgBox "clsInazuma.Draw"
    
    Nodes = Count() * 3 + 2
    
    Do Until Nodes = Line.Nodes.Count()
        
        If Line.Nodes.Count() < Nodes Then
            
            Line.Nodes.Insert 1, msoSegmentCurve, msoEditingAuto, 0#, 0#
                    
        Else
        
            Line.Nodes.Delete 1
            
        End If
        
    Loop
    
    Dim Position As Double
    Dim Base As Single
    Dim wkY As Single
    Dim wkRow As Range
    
    Position = Property.getColumnForDate(TargetDate)
    Base = getY(Position)
    
    Line.Select
    
    With Selection
        
        Set wkRow = ActiveSheet.Rows(Property.RowForChartTop)
        .ShapeRange.Nodes.SetPosition 1, Base, wkRow.Top
            
        Nodes = 2
        
        For i = 1 To Count()
            
            Set wkRow = ActiveSheet.Rows(Points(i).Row)
            
            .ShapeRange.Nodes.SetPosition Nodes, Base, wkRow.Top + 2
            Nodes = Nodes + 1
            
            If Points(i).Value < FIRSTDATE Then
                wkY = Base
            Else
                Position = Property.getColumnForDate(Points(i).Value)
                wkY = getY(Position)
            End If
            .ShapeRange.Nodes.SetPosition Nodes, _
                wkY, wkRow.Top + wkRow.Height / 2
            Nodes = Nodes + 1
            
            .ShapeRange.Nodes.SetPosition Nodes, _
                Base, ActiveSheet.Rows(Points(i).Row).Top + wkRow.Height - 2
            Nodes = Nodes + 1
            
        Next i
        
        Set wkRow = ActiveSheet.Rows(Property.RowForBottom)
        .ShapeRange.Nodes.SetPosition Nodes, _
            Base, wkRow.Top + wkRow.Height
            
        .ShapeRange.Line.ForeColor.RGB = Property.InazumaColor
        .ShapeRange.Line.Weight = Property.InazumaWidth
        
    End With
    
End Sub

'X座標取得
Private Function getY(Value As Double)
    
    Dim Column As Range
    
    Set Column = ActiveSheet.Columns(CInt(Value))
    
    getY = Column.Left + Column.Width * (Value - CInt(Value))
    
End Function
