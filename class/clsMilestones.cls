VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMilestones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'マイルストーン用クラス
Option Explicit
Option Base 1

Private Milestone() As clsMilestone
Private Line() As Shape
Private Text() As Shape

Private cntLine As Long
Private cntText As Long

Private Const objName = "MILESTONE"

Public Sub Draw(pr As clsProperty)
    
    Dim i As Long
    Dim j As Long
    Dim cntDraw As Long
    
    Call Load
    
    '必要な描画数を取得
    'cntDraw = CountDraw(pr.BeginDate, pr.EndDate)
    
    '描画するTextとLineを確保
    Call setShapes(Count())
    
    Dim Top As Single
    Dim Height As Single
    
    Top = ActiveSheet.Rows(pr.RowForChartTop).Top
    Height = ActiveSheet.Rows(pr.RowForBottom + 1).Top - Top
    
    For i = 1 To Count()
        
        If Milestone(i).isDraw(pr.BeginDate, pr.EndDate) Then
            Call setObjects(pr, i, Top, Height)
        End If
        
    Next
    
End Sub

Public Sub setObjects(pr As clsProperty, Num As Long, Top As Single, Height As Single)
    
    With Line(Num)
        
        .Visible = msoTrue
        
        .Left = mdlTools.GetXonChart(pr, Milestone(Num).TargetDate)
        .Width = 0
        .Top = Top
        .Height = Height
        
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = pr.MilestoneWidth
        
        .Name = objName & "Line" & Format(Num, "00000")
        
    End With
    
    With Text(Num)
        
        .Visible = msoTrue
        
        .Left = mdlTools.GetXonChart(pr, Milestone(Num).TargetDate) + 1
        .Top = Top + 1
        
        .Select
            Selection.Characters.Text = Milestone(Num).Name
            Selection.ShapeRange.Fill.Transparency = 0.5
            Selection.ShapeRange.Line.Visible = msoFalse
            Selection.AutoSize = True
        
        Call setTextBox
        
        .Name = objName & "Text" & Format(Num, "00000")
        
    End With
    
End Sub

Private Sub setTextBox()

    With Selection.Font
        .Name = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Public Sub setShapes(Num As Long)
    
    Dim s As Shape
    
    cntLine = 0
    cntText = 0
    
    For Each s In ActiveSheet.Shapes
        
        If s.Name Like objName & "*" Then
            
            Call AddObject(s)
            
        End If
        
    Next
    
    Do While cntLine < Num
        
        ActiveSheet.Shapes.AddLine(0#, 0#, 100#, 100#).Select
            Call AddObject(ActiveSheet.Shapes(Selection.Name))
            
    Loop
    
    Do While cntText < Num
    
        ActiveSheet.Shapes.AddTextbox(msoTextOrientationVerticalFarEast, 0#, 0#, 100#, 100#).Select
            Call AddObject(ActiveSheet.Shapes(Selection.Name))
    
    Loop
    
End Sub

Public Sub AddObject(s As Shape)
    
    Dim n As String
    
    n = objName
    
    Select Case s.Type
        
        Case msoLine
            
            cntLine = cntLine + 1
            n = n & "Line-" & cntLine
            ReDim Preserve Line(cntLine)
            Set Line(cntLine) = s
            
        Case msoTextBox
        
            cntText = cntText + 1
            n = n & "Text-" & cntText
            ReDim Preserve Text(cntText)
            Set Text(cntText) = s
        
    End Select
    
    s.Name = n
    s.Visible = msoFalse
    
End Sub

Public Property Get CountObjects(ShapeType As Integer) As Long
    
    On Error Resume Next
    
    Dim i As Long
    
    Select Case ShapeType
        
        Case msoLine
            i = UBound(Line)
            
        Case msoTextBox
            i = UBound(Text)
        
    End Select
    
    CountObjects = i
    
End Property


Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    
    i = UBound(Milestone)
    Count = i
    
End Property

Public Property Get CountDraw(BeginDate As Date, EndDate As Date) As Long
    
    Dim i As Long
    Dim j As Long
    
    j = 0
    For i = 1 To Count()
        
        If Milestone(i).isDraw(BeginDate, EndDate) Then j = j + 1
        
    Next
    
    CountDraw = j
    
End Property

Public Function Load() As Long
    
    Dim wk As clsMilestone
    Dim flg As Boolean
    Dim i As Long
    Dim j As Long
    
    i = 2
    j = 1
    
    Do
        
        Set wk = New clsMilestone
        flg = wk.Load(i)
        
        If flg And FIRSTDATE <= wk.TargetDate Then
            
            ReDim Preserve Milestone(j)
            Set Milestone(j) = wk
            j = j + 1
            
        Else
            
            Set wk = Nothing
            
        End If
        
        i = i + 1
        
    Loop While flg
    
    Load = j
    
End Function

Public Function Save() As Long
    
    Dim i As Long
    
    For i = 1 To Count()
        
        Call Milestone(i).Save(i + 1)
        
    Next
    
End Function

Public Sub Add(wk As clsMilestone)
    
    Dim i As Long
    i = Count() + 1
    
    ReDim Preserve Milestone(i)
    Set Milestone(i) = wk
    
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        Set Milestone(i) = Nothing
        
    Next
    
End Sub

Public Sub setControl(lst As Control)
    
    Dim i As Long
    
    lst.Clear
    
    For i = 1 To Count()
        
        lst.AddItem Milestone(i).Text
        
    Next i
    
End Sub

Public Property Get Milestones(Num As Long) As clsMilestone
    
    Set Milestones = Milestone(Num)
    
End Property

Public Property Let Milestones(Num As Long, wk As clsMilestone)
    
    Set Milestone(Num) = wk
    
End Property
