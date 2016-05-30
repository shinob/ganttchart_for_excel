VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Rect() As Shape
Private Line() As Shape
Private cntRect As Long
Private cntLine As Long

Private sht As Worksheet

'èâä˙âª
Public Sub Initialize(RectNum As Long, LineNum As Long)
    
    Dim s As Shape
    
    cntRect = 0
    cntLine = 0
    
    For Each s In sht.Shapes
        
        If s.Name Like CHARTBAR & "*" Then
            
            Select Case s.Type
                Case msoAutoShape
                    Call AddRect(s)
                Case msoLine
                    Call AddLine(s)
            End Select
            
        End If
        
    Next s
    
    Dim i As Long
    Dim j As Long
    
    If cntRect < RectNum Then
        
        j = RectNum - cntRect
        For i = 1 To j
            Call AddRect(sht.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0))
        Next i
        
    End If
    
    If cntLine < LineNum Then
        
        j = LineNum - cntLine
        'MsgBox "clsChartBar.Initialize " & j
        For i = 1 To j
            Call AddLine(sht.Shapes.AddLine(0, 0, 100, 100))
        Next i
        
    End If
    
    'MsgBox "Rect = " & cntRect & " Line = " & cntLine
    
    cntRect = 1
    cntLine = 1
    
End Sub

'í∑ï˚å`Çí«â¡
Private Sub AddRect(s As Shape)
    
    cntRect = cntRect + 1
    'MsgBox s.Name & " " & s.Type
    s.Name = CHARTBAR & "-" & cntRect + cntLine
    s.Visible = msoFalse
    s.Line.Weight = 0.75
    s.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    ReDim Preserve Rect(cntRect)
    Set Rect(cntRect) = s
    
End Sub

'íºê¸Çí«â¡
Private Sub AddLine(s As Shape)

    cntLine = cntLine + 1
    'MsgBox s.Name & " " & s.Type
    s.Name = CHARTBAR & "-" & cntRect + cntLine
    s.Visible = msoFalse
    
    ReDim Preserve Line(cntLine)
    Set Line(cntLine) = s
    
End Sub

'í∑ï˚å`ÇéÊìæ
Public Property Get getRect() As Shape
    
    Set getRect = Rect(cntRect)
    cntRect = cntRect + 1
    
End Property

'íºê¸ÇéÊìæ
Public Property Get getLine() As Shape

    Set getLine = Line(cntLine)
    cntLine = cntLine + 1
    
End Property

'å¬êî
Private Function Count(ChartType As Single) As Long
    
    On Error Resume Next
    
    Select Case ChartType
        
        Case gntChartRect
            Count = UBound(Rect)
        Case gntChartLine
            Count = UBound(Line)
            
    End Select
    
End Function

Private Sub Class_Initialize()
    Set sht = ActiveSheet
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    Dim cnt As Long
    
    cnt = Count(gntChartRect)
    For i = 1 To cnt
        Set Rect(i) = Nothing
    Next i
    
    cnt = Count(gntChartLine)
    For i = 1 To cnt
        Set Line(i) = Nothing
    Next i
    
End Sub
