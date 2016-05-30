VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'プロパティ用クラス
Option Explicit
Option Base 1

Public No As Integer

Public RefType As String

Public SheetName As String
Public BeginDate As Date
Public CellType As String
Public DrawColumns As Integer
Public DrawRows As Long

Public Top As Long
Public Left As Integer

Public AutoUpdate As Boolean

Public ChartWidth As Double
Public PlanPosition As Byte
Public ActPosition As Byte
Public ChartBarType As String

Public PlanLineColor As Long
Public PlanFillColor As Long
Public ActLineColor As Long
Public ActFillColor As Long

Public HOLIDAY As Range

Public InazumaDraw As Boolean
Public InazumaWidth As Integer
Public InazumaColor As Long

Public GroupingType As String
Public SortType As String
Public VisibleAll As Boolean
Public VisibleNotComplete As Boolean
Public InsertBlankRow As Boolean
Public PrintLabel As Boolean
Public CalcurateStatus As Boolean
Public WorkDaysExceptHoliday As Boolean

Public MilestoneWidth As Integer

Public UseMemos As Boolean

Public Complete As Range
Public PaintComplete As Boolean

Public ItemNo As Boolean

'読込
Public Sub Load(Num As Integer)
    
    If Num < DEFAULT_PROPERTY Then Exit Sub
    
    Dim strg As New clsStrage
    
    With strg
        
        Call .Initialize(shtProperty, False, CLng(Num), 2)
        
        No = Num
        RefType = .Data
        .Skip
        SheetName = .Data
        BeginDate = .Data
        CellType = .Data
        DrawColumns = .Data
        DrawRows = .Data
        .Skip
        Top = .Data
        Left = .Data
        .Skip
        AutoUpdate = .Data
        .Skip
        
        'If IsDefault Then Call .Initialize(shtProperty, False, DEFAULT_PROPERTY, .getPosition)
        If IsDefault Then .Column = DEFAULT_PROPERTY
        
        ChartWidth = .Data
        PlanPosition = .Data
        ActPosition = .Data
        ChartBarType = .Data
        .Skip
        PlanLineColor = .Data
        PlanFillColor = .Data
        ActLineColor = .Data
        ActFillColor = .Data
        .Skip
        Set HOLIDAY = .Data
        .Skip
        InazumaDraw = .Data
        InazumaWidth = .Data
        InazumaColor = .Data
        .Skip
        GroupingType = .Data
        SortType = .Data
        VisibleAll = .Data
        VisibleNotComplete = .Data
        InsertBlankRow = .Data
        PrintLabel = .Data
        CalcurateStatus = .Data
        WorkDaysExceptHoliday = .Data
        .Skip
        MilestoneWidth = .Data
        .Skip
        UseMemos = .Data
        .Skip
        Set Complete = .Data
        PaintComplete = .Data
        .Skip
        ItemNo = .Data
        
        Select Case CellType
            
            Case CELLTYPE_TIME
                BeginDate = CLng(BeginDate)
                
            Case CELLTYPE_10DAYS, CELLTYPE_MONTH
                BeginDate = DateSerial(Year(BeginDate), Month(BeginDate), 1)
                
        End Select
        
        
    End With
    
End Sub

'保存
Public Sub Save()
    
    If No < DEFAULT_PROPERTY Then Exit Sub
    
    Dim strg As New clsStrage
    
    With strg
        
        Call .Initialize(shtProperty, False, CLng(No), 2)
        
        .Data = RefType
        .Skip
        .Data = SheetName
        .Data = BeginDate
        .Data = CellType
        .Data = DrawColumns
        .Data = DrawRows
        .Skip
        .Data = Top
        .Data = Left
        .Skip
        .Data = AutoUpdate
        .Skip
        
        If IsDefault Then Call .Initialize(shtProperty, False, DEFAULT_PROPERTY, .getPosition)
        
        .Data = ChartWidth
        .Data = PlanPosition
        .Data = ActPosition
        .Data = ChartBarType
        .Skip
        .Data = PlanLineColor
        .Data = PlanFillColor
        .Data = ActLineColor
        .Data = ActFillColor
        .Skip
        .Skip   'Set .Data = Holiday
        .Skip
        .Data = InazumaDraw
        .Data = InazumaWidth
        .Data = InazumaColor
        .Skip
        .Data = GroupingType
        .Data = SortType
        .Data = VisibleAll
        .Data = VisibleNotComplete
        .Data = InsertBlankRow
        .Data = PrintLabel
        .Data = CalcurateStatus
        .Data = WorkDaysExceptHoliday
        .Skip
        .Data = MilestoneWidth
        .Skip
        .Data = UseMemos
        .Skip
        .Skip   'Set .Data = Compelte
        .Data = PaintComplete
        .Skip
        .Data = ItemNo

    End With
    
End Sub

'セル単位用一覧
Public Sub MakeCellTypeList(lst As Control)
    
    With lst
    
        .Clear
        
        .AddItem CELLTYPE_TIME
        .AddItem CELLTYPE_DAY
        .AddItem CELLTYPE_WEEK
        .AddItem CELLTYPE_10DAYS
        .AddItem CELLTYPE_MONTH
        
        '.Text = CellType
        .Value = CellType
        
    End With
    
End Sub

'チャートバー形式用一覧
Public Sub MakeChartBarTypeList(lst As Control)
    
    With lst
        
        .Clear
        
        .AddItem CHARTBARTYPE_RECT
        .AddItem CHARTBARTYPE_LINE
        
        '.Text = ChartBarType
        '.Caption = ChartBarType
        .Value = ChartBarType
        
    End With
    
End Sub

'描画の分類形式用一覧
Public Sub MakeGroupingTypeList(lst As Control)
    
    With lst
        
        .Clear
        
        .AddItem GROUPTYPE_CATEGORY
        .AddItem GROUPTYPE_PERSON
        
        .Value = GroupingType
        
    End With
    
End Sub

'描画の並替形式用一覧
Public Sub MakeSortTypeList(lst As Control)
    
    With lst
        
        .Clear
        
        .AddItem SORTTYPE_DATE
        .AddItem SORTTYPE_DATA
        
        .Value = SortType
        
    End With
    
End Sub

'標準
Public Property Get IsDefault() As Boolean
    
    Select Case RefType
        
        Case PROPTYPE_DEF
            IsDefault = True
        Case PROPTYPE_CST
            IsDefault = False
    End Select
    
End Property

'標準設定
Public Property Let RefTypeDefault(def As Boolean)
    
    If def Then
        RefType = PROPTYPE_DEF
    Else
        RefType = PROPTYPE_CST
    End If
    
End Property

'例番号から日付を取得
Public Property Get getDate(Column As Integer) As Date
        
    Dim wkDate As Date
    
    Dim y As Integer
    Dim m As Integer
    Dim D As Integer
    
    Select Case CellType
        
        Case CELLTYPE_TIME
            wkDate = Column / 24 + BeginDate
        Case CELLTYPE_DAY
            wkDate = Column + BeginDate
        Case CELLTYPE_WEEK
            wkDate = Column * 7 + BeginDate
        Case CELLTYPE_10DAYS
        
            y = Year(BeginDate)
            m = Int(Column / 3) + Month(BeginDate)
            
            D = Column Mod 3
            If D = 0 Then
                D = 1
            Else
                D = D * 10 + 1
            End If
            
            wkDate = DateSerial(y, m, D)
            
        Case CELLTYPE_MONTH
        
            y = Year(BeginDate)
            m = Month(BeginDate) + Column
            D = 1
            wkDate = DateSerial(y, m, D)
            
    End Select
    
    getDate = wkDate
    
End Property

'最終日取得
Public Property Get EndDate() As Date
    
    EndDate = getDate(DrawColumns)
    
End Property

'工程が期間内か否か
Public Property Get isScheduleOnChart(schedule As clsSchedule) As Byte
    
    Dim wk As Byte
    
    wk = 0
    With schedule
        
        If isDateOnChart(.PlanBegin, .PlanEnd) Then wk = wk + 1
        If isDateOnChart(.ActBegin, .ActEnd) Then wk = wk + 1
        
    End With
    
    isScheduleOnChart = wk
    
End Property

'日付が期間内か否か
Public Property Get isDateOnChart(wkBegin As Date, wkEnd As Date) As Boolean
    
    If (BeginDate <= wkBegin And wkBegin < EndDate) Or _
        (BeginDate < wkEnd And wkEnd <= EndDate) Or _
        (wkBegin < BeginDate And EndDate < wkEnd) Then
        isDateOnChart = True
    Else
        isDateOnChart = False
    End If
    
End Property

'日付から列を取得
Public Property Get getColumnForDate(wkDate As Date) As Double
    
    Dim wkCell As Double
    
    Dim y As Integer
    Dim m As Integer
    Dim D As Integer
    
    wkCell = wkDate - BeginDate
    If wkCell < 0 Then
        wkCell = 0
    ElseIf EndDate < wkDate Then
        wkCell = DrawColumns
    Else
    
        Select Case CellType
            
            Case CELLTYPE_TIME
                
                wkCell = wkCell * 24
                
            Case CELLTYPE_DAY
                
                'wkCell = Int(wkCell)
            
            Case CELLTYPE_WEEK
                
                wkCell = wkCell / 7
                
            Case CELLTYPE_10DAYS
                
                m = (Year(wkDate) - Year(BeginDate)) * 12 + (Month(wkDate) - Month(BeginDate))
                wkCell = m * 3
                
                D = Day(wkDate)
                
                If Hour(wkDate) + Minute(wkDate) > 0 Then
                    D = D + 1
                End If
                
                'wkCell = wkCell + Int(D / 10)
                
                'If D < 21 Then
                '    wkCell = wkCell + (D Mod 10) / 10
                'Else
                '    wkCell = wkCell + (D Mod 10) / Day(DateSerial(Year(wkDate), Month(wkDate) + 1, 1) - 1)
                'End If
                'MsgBox wkCell & " " & wkDate
                
                If D < 11 Then
                    wkCell = wkCell + (D - 1) / 10
                ElseIf D < 21 Then
                    wkCell = wkCell + (D - 11) / 10 + 1
                Else
                    wkCell = wkCell + (D - 21) / (Day(DateSerial(Year(wkDate), Month(wkDate) + 1, 1) - 1) - 20) + 2
                End If
                
            Case CELLTYPE_MONTH
                
                m = (Year(wkDate) - Year(BeginDate)) * 12 + (Month(wkDate) - Month(BeginDate))
                D = Day(wkDate)
                wkCell = m + D / Day(DateSerial(Year(wkDate), Month(wkDate) + 1, 1) - 1)
                
        End Select
    
    End If
    
    getColumnForDate = wkCell + ColumnForChart
    
End Property

Public Property Get isSortByDate() As Boolean
    
    'MsgBox "clsProperty.isSortByDate " & SortType
    If SortType = SORTTYPE_DATE Then isSortByDate = True
    
End Property

Public Property Get isGroupingByCategory() As Boolean
    
    If GroupingType = GROUPTYPE_CATEGORY Then isGroupingByCategory = True
    
End Property

'列番号取得
Public Property Get ColumnForNo() As Integer
    ColumnForNo = Left
End Property

Public Property Get ColumnForName() As Integer
    ColumnForName = Left + 1
End Property

Public Property Get ColumnForPerson() As Integer
    ColumnForPerson = Left + 2
End Property

Public Property Get ColumnForWorkDays() As Integer
    ColumnForWorkDays = Left + 3
End Property

Public Property Get ColumnForStatus() As Integer
    ColumnForStatus = Left + 4
End Property

Public Property Get ColumnForNote() As Integer
    ColumnForNote = Left + 6 + DrawColumns
End Property

Public Property Get ColumnForChart() As Integer
    ColumnForChart = Left + 6
End Property

'行番号取得
Public Property Get RowForDate() As Long
    RowForDate = Top + 1
End Property

Public Property Get RowForWeek() As Long
    RowForWeek = Top + 2
End Property

Public Property Get RowForChartTop() As Long
    RowForChartTop = Top + 3
End Property

Public Property Get RowForBottom() As Long
    RowForBottom = Top + DrawRows + 2
End Property

'先頭日付用セル取得
Public Property Get getDateCell() As Range
    
    Set getDateCell = ActiveSheet.Cells(Top + 1, ColumnForChart)
    
End Property

'チャート描画範囲取得
Public Property Get getChartRange() As Range
    
    Dim wkLeft As Integer
    Dim TopLeft As Range
    Dim BottomRight As Range
    
    If UseMemos Then
        wkLeft = Left - 5
    Else
        wkLeft = Left
    End If
    
    With ActiveSheet
        
        Set TopLeft = .Cells(RowForChartTop, wkLeft)
        Set BottomRight = .Cells(RowForBottom, ColumnForNote)
        Set getChartRange = Range(TopLeft, BottomRight)
        
    End With
    
End Property

'日付範囲取得
Public Property Get getDateRange(wkColumn As Integer) As Range
    
    Dim TopCell As Range
    Dim BottomCell As Range
    
    With ActiveSheet
        
        Set TopCell = .Cells(RowForWeek, wkColumn)
        Set BottomCell = .Cells(RowForBottom, wkColumn)
        Set getDateRange = Range(TopCell, BottomCell)
        
    End With
    
End Property

'
Public Property Get getRowRange(wkRow As Long) As Range
    
    Dim wkLeft As Integer
    Dim wkRight As Integer
    
    With ActiveSheet
        
        Set getRowRange = Range( _
            .Cells(wkRow, Left), .Cells(wkRow, ColumnForNote) _
        )
        
    End With
    
End Property

'日付からX座標を取得
Public Function GetXonChart(TargetDate As Date) As Single
    
    Dim Position As Double
    Dim Column As Range
    
    Position = getColumnForDate(TargetDate)
    Set Column = ActiveSheet.Columns(CInt(Position))
    
    GetXonChart = Column.Left + Column.Width * (Position - CInt(Position))
    
End Function
