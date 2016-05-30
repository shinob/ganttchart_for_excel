VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'項目クラス
Option Explicit
Option Base 1

Public No As Long
Public Name As String
Public Category As Long
Public Person As Long
Public Note As String
Public Complete As Boolean
Public Status As Integer
Public Hyperlink As String
Public LinkItem As Long

Private Ina(COUNT_INAZUMA) As Date
Public Memo As New clsMemos

Public Schedules As New clsSchedules
Public Visible As Boolean

Public Used As Boolean

Public SubItems As clsItems

Private sht As Worksheet

'保存
Public Function Save() As Boolean
    
    If No < 2 Then
        No = FindBlankRow(sht, 2, 2)
    End If
    
    With sht
    
        .Cells(No, 1) = No
        .Cells(No, 2) = Name
        .Cells(No, 3) = Category
        .Cells(No, 4) = Person
        .Cells(No, 5) = Note
        .Cells(No, 6) = Complete
        .Cells(No, 7) = Status
        .Cells(No, 8) = Hyperlink
        .Cells(No, 9) = LinkItem
        '.Cells(NO, 10) = NotUsing...
        
        Dim i As Integer
        
        For i = 1 To COUNT_INAZUMA
            .Cells(No, 10 + i) = Ina(i)
        Next i
        
        For i = 1 To COUNT_MEMO
            .Cells(No, 30 + i) = Memo.Text(i)
        Next i
        
    End With
    
    Save = True
    
End Function

'イナズマ線保存
Public Function SaveInazuma() As Boolean
    
    If No < 2 Then
        Exit Function
    End If
    
    With sht
        
        Dim i As Integer
        
        For i = 1 To 20
            
            .Cells(No, 10 + i) = Ina(i)
            
        Next i
        
    End With
    
End Function

'分類データのみ保存
Public Function SaveCategory() As Boolean
    
    If No < 2 Then
        Exit Function
    End If
    
    With sht
        
        .Cells(No, 3) = Category
        
    End With
    
End Function

'上位項目のみ保存
Public Sub SaveLinkItem()
    
    If No < 2 Then
        Exit Sub
    End If
    
    With sht
        .Cells(No, 9).Value = LinkItem
    End With
    
End Sub

'下位項目のリンクを現在の番号に合わせる
Public Sub UpdateLinkItemOfSubItems()
    
    Dim i As Long
    
    For i = 1 To SubItems.Count()
        
        With SubItems.getByIndex(i)
            .LinkItem = No
            .SaveLinkItem
        End With
        
    Next i
    
End Sub

Public Sub ClearSubItems()
    Set SubItems = Nothing
    Set SubItems = New clsItems
End Sub

'読込
Public Function Load(Row As Long) As Boolean

    If sht.Cells(Row, 1) = "" Then
        Load = False
        Exit Function
    End If
    
    With sht
    
        No = .Cells(Row, 1)
        Name = .Cells(Row, 2)
        Category = .Cells(Row, 3)
        Person = .Cells(Row, 4)
        Note = .Cells(Row, 5)
        Complete = .Cells(Row, 6)
        Status = .Cells(Row, 7)
        Hyperlink = .Cells(Row, 8)
        LinkItem = .Cells(Row, 9)
        'NotUsing... = .Cells(Row, 10)
        
        Dim i As Integer
        
        For i = 1 To COUNT_INAZUMA
            Ina(i) = .Cells(Row, 10 + i)
        Next i
        
        For i = 1 To COUNT_MEMO
            Memo.Text(i) = .Cells(Row, 30 + i)
        Next i
        
    End With
    
    'Schedules.Load Row
    
    Load = True
    
End Function

'工程読込
Public Sub LoadSchedule()
    
    Call Schedules.Load(No)
    
End Sub

'イナズマ線取得
Public Property Get Inazuma(Num As Integer) As Date
    
    Inazuma = Ina(Num)
    
End Property

'イナズマ線設定
Public Property Let Inazuma(Num As Integer, Value As Date)
    
    Ina(Num) = Value
    
End Property

Public Sub setInazuma(Num As Integer, Value As Date)
    
    If Complete Then
        Ina(Num) = Value
    ElseIf BeginDate < Value Then
        Ina(Num) = Schedules.ActLatest
    Else
        If BeginDate < Schedules.ActLatest Then
            Ina(Num) = Schedules.ActLatest
        Else
            Ina(Num) = Value
        End If
    End If
    
End Sub

'開始日
Public Property Get BeginDate() As Date
    
    BeginDate = Schedules.BeginDate
    
End Property

'終了日
Public Property Get EndDate() As Date
    
    EndDate = Schedules.EndDate
    
End Property

'作業日数取得
Public Property Get getWorkDays() As Double
    
    Dim i As Long
    Dim wk As Double
    Dim sdl As clsSchedule
    
    wk = 0
    For i = 1 To Schedules.Count()
        
        Set sdl = Schedules.Items(i)
        
        If sdl.PlanBegin <= sdl.PlanEnd Then
            wk = wk + (sdl.PlanEnd - sdl.PlanBegin) * sdl.Weight
        Else
            wk = wk + sdl.Weight
        End If
        
    Next i
    
    getWorkDays = wk
    
End Property

'作業日数取得(除休日)
Public Property Get getWorkDaysWithOutHolidays() As Double
    
    Dim i As Long
    Dim wk As Double
    
    wk = 0
    For i = 1 To Schedules.Count()
        
        With Schedules.Items(i)
            wk = wk + .WorkDays * .Weight
        End With
        
    Next i
    
    getWorkDaysWithOutHolidays = wk
    
End Property

'作業進捗状況
Public Property Get getStatus() As Single
    
    Dim i As Long
    Dim wkPlan As Single
    Dim wkAct As Single
    Dim sdl As clsSchedule
    
    wkPlan = 0
    wkAct = 0
    
    For i = 1 To Schedules.Count()
        
        Set sdl = Schedules.Items(i)
        
        If sdl.PlanBegin <= sdl.PlanEnd Then
            wkPlan = wkPlan + (sdl.PlanEnd - sdl.PlanBegin) * sdl.Weight
        End If
        
        If sdl.ActBegin <= sdl.ActEnd Then
            wkAct = wkAct + (sdl.ActEnd - sdl.ActBegin) * sdl.Weight
        End If
        
    Next i
    
    If wkPlan = 0 Then
        getStatus = 0
    Else
        getStatus = wkAct / wkPlan * 100
    End If
    
End Property

'期間内の工程有無
Public Property Get haveDrawSchedule() As Boolean
    
    Dim i As Long
    
    For i = 1 To Schedules.Count()
        If Schedules.Items(i).Draw Then
            haveDrawSchedule = True
            Exit For
        End If
    Next i
    
End Property

Public Property Get DataSheet() As Worksheet
    Set DataSheet = sht
End Property

Private Sub Class_Initialize()
    
    Set sht = shtItem
    Set SubItems = New clsItems
    
End Sub

Private Sub Class_Terminate()
    
    Set Schedules = Nothing
    Set SubItems = Nothing
    
End Sub
