VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'工程クラス
Option Explicit
Option Base 1

Public No As Long
Public Name As String
Public Item As Long
Public PlanBegin As Date
Public PlanEnd As Date
Public PlanColor As Long
Public ActBegin As Date
Public ActEnd As Date
Public ActColor As Long
Public ChartType As Byte
Public Weight As Long

Public Delete As Boolean    '削除フラグ
Public Draw As Boolean      '描画フラグ

Public WorkDays As Double

Private sht As Worksheet

Public Property Get getBegin(mode As Integer) As Date
    
    Select Case mode
        
        Case MODE_PLAN
            getBegin = PlanBegin
        Case MODE_ACT
            getBegin = ActBegin
            
    End Select
    
End Property

Public Property Get getEnd(mode As Integer) As Date
    
    Dim b As Date
    Dim e As Date
    
    Select Case mode
    
        Case MODE_PLAN
            b = PlanBegin
            e = PlanEnd
        Case MODE_ACT
            b = ActBegin
            e = ActEnd
    End Select
    
    If b < FIRSTDATE Then
        getEnd = 0
    ElseIf e < b Then
        getEnd = b + 1
    Else
        getEnd = e
    End If
    
End Property

Public Property Get Text() As String
    
    Dim s As String
    
    s = No & vbCr & _
        Name & vbCr & _
        Item & vbCr & _
        PlanBegin & vbCr
    
    Text = s
    
End Property

Public Sub Copy(Value As clsSchedule)
    
    With Value
        
        Name = .Name
        Item = .Item
        PlanBegin = .PlanBegin
        PlanEnd = .PlanEnd
        PlanColor = .PlanColor
        ActBegin = .ActBegin
        ActEnd = .ActEnd
        ActColor = .ActColor
        Weight = .Weight
        
    End With
    
End Sub

'保存
Public Function Save() As Boolean
    
    'MsgBox "clsSchedule.NO = " & NO
    
    If No < 2 Then
        
        If Delete Then Exit Function
        
        No = FindBlankRow(sht, 2, 3)
        
    End If
    
    Dim strg As New clsStrage
    
    Call strg.Initialize(sht, True, No, 1)
    
    If Delete Then
        
        Dim i As Integer
        strg.Skip
        For i = 2 To 9: strg.Data = "": Next i
    
    Else
        
        With strg
            
            .Data = No
            .Data = Name
            .Data = Item
            .Data = PlanBegin
            .Data = PlanEnd
            .Data = PlanColor
            .Data = ActBegin
            .Data = ActEnd
            .Data = ActColor
            .Data = ChartType
            .Data = Weight
            
        End With
        
    End If
    
    Save = True
    
End Function

Public Sub SaveItem()
    
    If No < 2 Then Exit Sub
    
    With sht
        .Cells(No, 3).Value = Item
    End With
    
End Sub

'読込
Public Function Load(Row As Long) As Boolean
    
    If sht.Cells(Row, 1) = "" Then
        Load = False
        Exit Function
    End If
    
    Dim strg As New clsStrage
    
    Call strg.Initialize(sht, True, Row, 1)
    
    With strg
        No = .Data
        Name = .Data
        Item = .Data
        PlanBegin = .Data
        PlanEnd = .Data
        PlanColor = .Data
        ActBegin = .Data
        ActEnd = .Data
        ActColor = .Data
        ChartType = .Data
        Weight = .Data
    End With
    
    If Weight < 1 Then Weight = 1
    Load = True
    
End Function

'開始日
Public Property Get BeginDate() As Date
    
    BeginDate = #1/1/2100#
    
    If PlanBegin < FIRSTDATE Then
        
        If ActBegin < FIRSTDATE Then
            
        Else
        
            BeginDate = ActBegin
            
        End If
        
    Else
        
        If ActBegin < FIRSTDATE Then
        
            BeginDate = PlanBegin
            
        Else
            
            If PlanBegin < ActBegin Then
            
                BeginDate = PlanBegin
                
            Else
            
                BeginDate = ActBegin
                
            End If
            
        End If
        
    End If
    
End Property

Public Property Get EndDate() As Date
    
    If PlanEnd < FIRSTDATE Then
        
        If ActEnd < FIRSTDATE Then
            
        Else
        
            EndDate = ActEnd
            
        End If
        
    Else
        
        If ActEnd < FIRSTDATE Then
        
            EndDate = PlanEnd
            
        Else
            
            If ActEnd < PlanEnd Then
            
                EndDate = PlanEnd
                
            Else
            
                EndDate = ActEnd
                
            End If
            
        End If
        
    End If
    
End Property

Public Sub CalcWorkDays(hd As clsHoliday)
    
    Dim i As Long
    Dim wkDate As Date
    Dim msg As String
    
    wkDate = PlanBegin
    WorkDays = 0
    
    If PlanBegin < FIRSTDATE Then Exit Sub
    
    If PlanBegin < PlanEnd Then
    
        Do Until PlanEnd < wkDate
        
            If hd.isHoliday(wkDate) Then
                
            Else
                
                WorkDays = WorkDays + (CLng(wkDate) + 1) - wkDate
                If PlanEnd < CLng(wkDate) + 1 Then
                    WorkDays = WorkDays - (CLng(wkDate) + 1 - PlanEnd)
                End If
                
            End If
            wkDate = CLng(wkDate) + 1
        Loop
        
    Else
        
        If hd.isHoliday(PlanBegin) Then
        
        Else
            WorkDays = 1
        End If
        
    End If
    
    'MsgBox "clsSchedule.CalcWorkDays " & WorkDays
    
End Sub

Public Function getWorkDays() As Double
    
    If PlanBegin < FIRSTDATE Then
        getWorkDays = 0
    ElseIf PlanBegin < PlanEnd Then
        getWorkDays = PlanEnd - PlanBegin
    Else
        getWorkDays = 1
    End If
    
End Function

Private Sub Class_Initialize()
    Set sht = shtSchedule
    Weight = 1
End Sub
