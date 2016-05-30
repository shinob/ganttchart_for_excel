VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGanttChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'チャート描画用クラス
Option Explicit
Option Base 1

Private prpt As New clsProperty
Private psn As New clsPersons
Private ctg As New clsCategorys
Private itm As New clsItems
Private sdl As New clsSchedules

Private chb As New clsChartBars
Private edb As New clsEditBoxes

Private inz As New clsInazumas

Public sht As Worksheet

Private ActiveRow As Long

Private flgEndDrawing As Boolean
Private flgChangeDrawingRows As Boolean

'プロパティ読込
Private Sub LoadProperty()

    Dim ChartNo As Integer
    ChartNo = mdlTools.ChartNum(sht)
    'MsgBox ChartNo
    
    Call prpt.Load(CLng(ChartNo))
    'MsgBox "clsGanttChart.LoadProperty " & prpt.BeginDate & " = " & prpt.EndDate
    
End Sub

'データ読込
Private Sub LoadData()

    'Debug.Print "Load Data " & Now()
    
    Dim i As Long
    Dim msg As String
    
    '担当読込(不要)
    'Call psn.Load
    
    '分類読込(不要)
    'Call ctg.Load
    
    '項目読込
    Call itm.Load
    'Debug.Print "Item Loaded " & itm.Count()
    
    'EditBox、イナズマ線初期化
    Call edb.Initialize(itm.Count())
    Call inz.Initialize(itm)
    'Debug.Print "Item Loaded " & Now()
    
    '工程読込
    Call sdl.Load(-1)
    
    '工程の表示・非表示を確認
    Dim wk As clsSchedule
    Dim cnt As Long
    Dim cntRect As Long
    Dim cntLine As Long
    Dim hld As New clsHoliday
    
    For i = 1 To sdl.Count()
    
        Set wk = sdl.Items(i)
        
        Call wk.CalcWorkDays(hld)
        
        If wk.PlanEnd < wk.PlanBegin Then wk.PlanEnd = wk.PlanBegin + 1
        If wk.ActEnd < wk.ActBegin Then wk.ActEnd = wk.ActBegin + 1
        
        cnt = prpt.isScheduleOnChart(wk)
        If 0 < cnt Then
            wk.Draw = True
            Select Case wk.ChartType
                Case gntChartRect
                    cntRect = cntRect + cnt
                Case gntChartLine
                    cntLine = cntLine + cnt
            End Select
        Else
            wk.Draw = False
        End If
        
    Next i
    
    'チャートバー初期化
    'MsgBox "clsGanttChart.LoadData Rect = " & cntRect & " Line = " & cntLine
    Call chb.Initialize(cntRect, cntLine)
    'Debug.Print "Schedule Loaded " & Now()
    
    '項目へ工程を設定
    Dim j As Long
    For i = 1 To sdl.Count()
        
        j = sdl.Items(i).Item
        If j < 2 Then
        
        Else
            With itm.Items(j - 1).Schedules
                .Add
                .Items(.Count()) = sdl.Items(i)
            End With
        End If
        
    Next i
    
    'Debug.Print "Schedule Marged to Item " & Now()
    
End Sub

'分類へ項目を設定
Private Sub MargeDataByCategory()
    
    Dim i As Long
    Dim j As Long
    
    '分類へ項目を設定
    For i = 1 To itm.Count()
        
        j = itm.Items(i).LinkItem
        If j < 2 Then
        Else
            With itm.Items(j - 1).SubItems
                .Add
                .Items(.Count()) = itm.Items(i)
            End With
        End If
        
        j = itm.Items(i).Category
        If j < 2 Then
        
        Else
            With ctg.Items(j - 1).Items
                .Add
                .Items(.Count()) = itm.Items(i)
            End With
        End If
        
    Next i
    
    '項目並替
    'MsgBox "clsGanttChart.LoadData " & prpt.isSortByDate
    If prpt.isSortByDate Then
        
        For i = 1 To ctg.Count()
            
            ctg.Items(i).Items.Sort
            
        Next i
        
    End If
    
End Sub

'担当へ項目を設定
Private Sub MargeDataByPerson()

    Dim i As Long
    Dim j As Long
    
    '分類へ項目を設定
    For i = 1 To itm.Count()
        
        j = itm.Items(i).Person
        If j < 2 Then
        
        Else
            With psn.Items(j - 1).Items
                .Add
                .Items(.Count()) = itm.Items(i)
            End With
        End If
        
    Next i
    
    '項目並替
    'MsgBox "clsGanttChart.LoadData " & prpt.isSortByDate
    If prpt.isSortByDate Then
        
        For i = 1 To psn.Count()
            
            psn.Items(i).Items.Sort
            
        Next i
        
    End If
    
End Sub

'描画行の移動
Private Sub NextRow()
    
    If isEndDrawing Then Exit Sub
    
    ActiveRow = ActiveRow + 1
    
    If prpt.RowForBottom < ActiveRow And flgChangeDrawingRows = False Then
        
        If vbYes = MsgBox("描画するチャート行数が設定値を超えました" & _
            vbCr & "描画を続けますか?", vbYesNo) Then
            flgChangeDrawingRows = True
        End If
        
        flgEndDrawing = True
        
    End If
    
    If Not isEndDrawing Then Call CopyRowFormat
    
End Sub

'描画終了か否か
Private Function isEndDrawing() As Boolean
    
    If flgChangeDrawingRows = False And flgEndDrawing = True Then
        isEndDrawing = True
    End If
    
End Function

'行の書式コピー
Private Sub CopyRowFormat()
    
    prpt.getRowRange(prpt.RowForChartTop).Copy
    prpt.getRowRange(ActiveRow).PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub

'チャート描画
Public Sub DrawChart()
    
    prpt.getChartRange.Select
    Selection.ClearContents
    'MsgBox "clsGanttChart.DrawChart Clear!!"
    
    Call UpdateDate
    Call PaintHoliday
    
    ActiveRow = prpt.RowForChartTop
    
    Dim i As Long
    
    If prpt.isGroupingByCategory Then
        For i = 1 To ctg.Count()
        
            If isEndDrawing Then Exit For
            
            If ctg.Items(i).Visible Then
                Call PrintCategory(ctg.Items(i), CStr(i))
                If prpt.InsertBlankRow Then Call NextRow
            End If
            
            SetStatusString "チャート描画中", i / ctg.Count()
            
        Next i
    Else
        For i = 1 To psn.Count()
        
            If isEndDrawing Then Exit For
            
            Call PrintPerson(psn.Items(i), CStr(i))
            If prpt.InsertBlankRow Then Call NextRow
            
        Next i
    End If
    
    Do While ActiveRow < prpt.RowForBottom
        
        ActiveRow = ActiveRow + 1
        Call CopyRowFormat
        
    Loop
    
    Call inz.Draw(prpt)
    
    Dim mls As New clsMilestones
    Call mls.Draw(prpt)
    
    If flgChangeDrawingRows Then
        prpt.DrawRows = ActiveRow - prpt.RowForChartTop + 1
        prpt.Save
    End If
    
    'Debug.Print "Drawing Complete. " & Now()
    
End Sub

'分類表示
Public Sub PrintCategory(wkCategory As clsCategory, Tag As String)
    
    With sht.Cells(ActiveRow, prpt.ColumnForNo)
        .HorizontalAlignment = xlRight
        .Value = "'" & Tag
    End With
    
    If prpt.UseMemos Then Call PrintMemo(wkCategory.Memo, ActiveRow)
    
    With sht.Cells(ActiveRow, prpt.ColumnForName)
        .Value = wkCategory.Name
        .ShrinkToFit = True
        .IndentLevel = 0
    End With
    
    Call NextRow
    
    Dim i As Long
    For i = 1 To wkCategory.Items.Count()
        If isEndDrawing Then Exit For
        Call PrintItem(wkCategory.Items.Items(i), 1, Tag & "-" & i)
    Next i
    
End Sub

'担当表示
Public Sub PrintPerson(wkPerson As clsPerson, Tag As String)
    
    With sht.Cells(ActiveRow, prpt.ColumnForNo)
        .HorizontalAlignment = xlRight
        .Value = "'" & Tag
    End With
    
    With sht.Cells(ActiveRow, prpt.ColumnForName)
        .Value = wkPerson.Name
        .ShrinkToFit = True
        .IndentLevel = 0
    End With
    
    Call NextRow
    
    Dim i As Long
    For i = 1 To wkPerson.Items.Count()
        If isEndDrawing Then Exit For
        Call PrintItem(wkPerson.Items.Items(i), 1, Tag & "-" & i)
    Next i
    
End Sub

Private Sub PrintMemo(wkMemo As clsMemos, wkRow As Long)
    
    Dim i As Integer
    
    For i = 1 To COUNT_MEMO
        
        sht.Cells(wkRow, prpt.ColumnForNo - COUNT_MEMO + i - 1).Value = wkMemo.Text(i)
        
    Next i
    
End Sub

'項目表示
Public Sub PrintItem(wkItem As clsItem, wkDeps As Integer, Tag As String)
    
    Dim i As Long
    
    '階層は15まで(Excelの機能による制限)
    If 15 < wkDeps Or wkItem.Used Then Exit Sub
    
    If prpt.VisibleAll Or wkItem.haveDrawSchedule Or _
        (prpt.VisibleNotComplete And Not wkItem.Complete) Then
        '描画します
    Else
        '描画しません
        Exit Sub
    End If
    
    'NO列の設定
    'MsgBox "clsGanttChart.PrintItem " & wkItem.No
    'sht.Cells(ActiveRow, prpt.ColumnForNo).Value = wkItem.No
    sht.Cells(ActiveRow, prpt.ColumnForNo).Value = ""
    Call edb.setEditBox(wkItem.No, sht.Cells(ActiveRow, prpt.ColumnForNo))
    
    If prpt.UseMemos Then Call PrintMemo(wkItem.Memo, ActiveRow)
    
    '項目名設定
    sht.Cells(ActiveRow, prpt.ColumnForName).Select
    With Selection
        
        If prpt.ItemNo Then
            .Value = wkItem.Name
        Else
            .Value = Tag & " " & wkItem.Name
        End If
        
        If wkItem.Hyperlink <> "" Then _
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                wkItem.Hyperlink, TextToDisplay:=.Value
        .ShrinkToFit = True
        .IndentLevel = wkDeps
        
    End With
    
    '担当者名・作業時間・進捗状況・備考表示
    If 2 <= wkItem.Person Then
        If wkItem.Person <= psn.Count + 1 Then
            sht.Cells(ActiveRow, prpt.ColumnForPerson).Value = _
                psn.Items(wkItem.Person - 1).Name
        Else
            sht.Cells(ActiveRow, prpt.ColumnForPerson).Value = _
                "担当の設定が間違っています。"
        End If
    End If
            
    If prpt.WorkDaysExceptHoliday Then
        sht.Cells(ActiveRow, prpt.ColumnForWorkDays).Value = _
            wkItem.getWorkDaysWithOutHolidays
    Else
        sht.Cells(ActiveRow, prpt.ColumnForWorkDays).Value = wkItem.getWorkDays
    End If
    
    If wkItem.Complete Then
        'sht.Cells(ActiveRow, prpt.ColumnForStatus).Value = 100
        wkItem.Status = 100
    End If
    
    If prpt.CalcurateStatus And Not wkItem.Complete Then
        sht.Cells(ActiveRow, prpt.ColumnForStatus).Value = wkItem.getStatus
    Else
        sht.Cells(ActiveRow, prpt.ColumnForStatus).Value = wkItem.Status
    End If
    
    Dim rng As Range
    
    If wkItem.Status = 100 Then
        
        '状態が100%ならば行を塗りつぶす
        'Set rng = prpt.getDateRange(i + prpt.ColumnForChart())
        'Set rng = Range(sht.Cells(ActiveRow, prpt.ColumnForStatus + 2),
        Set rng = Range(sht.Cells(ActiveRow, prpt.ColumnForNo), _
            sht.Cells(ActiveRow, prpt.ColumnForNote - 1))
        rng.Select
        
        With Selection.Interior
            
            If prpt.PaintComplete Then
                .color = prpt.Complete.Interior.color
                .Pattern = prpt.Complete.Interior.Pattern
                .PatternColor = prpt.Complete.Interior.PatternColor
                'msg = msg & i & " Painted" & vbCr
            Else
                '.ColorIndex = xlNone
                'msg = msg & i & " Not Painted" & vbCr
            End If
            
        End With
        
    End If
    
    sht.Cells(ActiveRow, prpt.ColumnForNote).Value = wkItem.Note
    
    Call inz.Add(wkItem, ActiveRow)
    
    '工程表示
    For i = 1 To wkItem.Schedules.Count()
        
        Call PrintChartBar(wkItem.Schedules.Items(i))
        
    Next i
    
    wkItem.Used = True
    Call NextRow
    
    '下位項目表示
    For i = 1 To wkItem.SubItems.Count()
        If isEndDrawing Then Exit For
        Call PrintItem(wkItem.SubItems.Items(i), wkDeps + 1, Tag & "-" & i)
    Next i
    
End Sub

'チャート表示
Public Sub PrintChartBar(wkSchedule As clsSchedule)
    
    If Not wkSchedule.Draw Then Exit Sub
    
    Dim s As Shape
    Dim cl As Long
    Dim flgLabel As Boolean
    
    flgLabel = Not prpt.PrintLabel
    
    With wkSchedule
        
        '予定用
        If prpt.isDateOnChart(.PlanBegin, .PlanEnd) Then
        
            Select Case .ChartType
                
                Case gntChartRect
                    Set s = chb.getRect
                    s.Line.ForeColor.RGB = prpt.PlanLineColor
                Case gntChartLine
                    Set s = chb.getLine
                    
            End Select
            
            If .PlanColor < 0 Then
                cl = prpt.PlanFillColor
            Else
                cl = .PlanColor
            End If
            
            Call setChartBar(.PlanBegin, .PlanEnd, s, prpt.PlanPosition, cl)
            s.Name = CHARTBAR & "Plan" & Format(.No, "00000")
            
            If Not flgLabel Then
                sht.Cells(ActiveRow, Int(prpt.getColumnForDate(.PlanBegin))).Value = .Name
                flgLabel = True
            End If
            
        End If
        
        '実績用
        If prpt.isDateOnChart(.ActBegin, .ActEnd) Then
        
            Select Case .ChartType
                
                Case gntChartRect
                    Set s = chb.getRect
                    s.Line.ForeColor.RGB = prpt.ActLineColor
                Case gntChartLine
                    Set s = chb.getLine
            End Select
            
            If .ActColor < 0 Then
                cl = prpt.ActFillColor
            Else
                cl = .ActColor
            End If
            
            Call setChartBar(.ActBegin, .ActEnd, s, prpt.ActPosition, cl)
            s.Name = CHARTBAR & "Act" & Format(.No, "00000")
            
            If Not flgLabel Then _
                sht.Cells(ActiveRow, Int(prpt.getColumnForDate(.ActBegin))).Value = .Name
            
        End If
        
    End With
    
End Sub

'チャート設定
Public Sub setChartBar(wkBegin As Date, wkEnd As Date, wkShape As Shape, _
    wkPosition As Byte, wkColor As Long)
    
    Dim wkPos As Double
    Dim wkTop As Single
    Dim wkLeft As Double
    Dim wkHeight As Single
    Dim wkWidth As Double
    
    wkShape.Visible = msoTrue
    wkShape.Select
    With Selection
        
        '左端設定
        wkPos = prpt.getColumnForDate(wkBegin)
        wkLeft = sht.Columns(Int(wkPos)).Left
        wkWidth = sht.Columns(Int(wkPos)).Width
        
        .Left = wkLeft + wkWidth * (wkPos - Int(wkPos))
        
        '幅設定
        wkPos = prpt.getColumnForDate(wkEnd)
        wkLeft = sht.Columns(Int(wkPos)).Left
        wkWidth = sht.Columns(Int(wkPos)).Width
        
        .Width = wkLeft + wkWidth * (wkPos - Int(wkPos)) - .Left
        
        '縦位置設定
        wkTop = sht.Rows(ActiveRow).Top
        wkHeight = sht.Rows(ActiveRow).Height
        
        .Top = wkTop + wkHeight * (wkPosition / 100)
        
        Select Case wkShape.Type
            
            Case msoAutoShape
                .Height = prpt.ChartWidth
                .ShapeRange.Fill.ForeColor.RGB = wkColor
            Case msoLine
                .ShapeRange.Line.Weight = prpt.ChartWidth * 2 / 3
                .ShapeRange.Line.ForeColor.RGB = wkColor
                .Height = 0
        End Select
        
        .OnAction = "ShowScheduleForm"
        
    End With
    
End Sub

'休日表示
Public Sub PaintHoliday()
    
    Select Case prpt.CellType
        
        Case CELLTYPE_TIME, CELLTYPE_DAY
        
        Case Else
            Exit Sub
        
    End Select
    'MsgBox "clsGanttChart.PaintHoliday"
    
    Dim hld As New clsHoliday
    Dim rng As Range
    Dim i As Integer
    Dim wkDate As Date
    
    Dim msg As String
    
    For i = 0 To prpt.DrawColumns - 1
        
        Set rng = prpt.getDateRange(i + prpt.ColumnForChart())
        rng.Select
        
        With Selection.Interior
            
            If hld.isHoliday(prpt.getDate(i)) Then
                .color = prpt.HOLIDAY.Interior.color
                .Pattern = prpt.HOLIDAY.Interior.Pattern
                .PatternColor = prpt.HOLIDAY.Interior.PatternColor
                msg = msg & i & " Painted" & vbCr
            Else
                .ColorIndex = xlNone
                msg = msg & i & " Not Painted" & vbCr
            End If
            
        End With
        
    Next i
    
    'MsgBox "clsGanttChart.PaintHoliday " & vbCr & msg
    
End Sub

'日付表示更新
Public Sub UpdateDate()
    
    Dim i As Integer
    Dim Row As Integer
    Dim col As Integer
    
    Dim wkDate As Date
    Dim wkTmp As Integer
    
    Row = prpt.RowForDate
    col = prpt.ColumnForChart
    
    sht.Cells(Row, col).Value = prpt.BeginDate
            
    Select Case prpt.CellType
        
        '時間
        Case CELLTYPE_TIME
        
        '日付、週
        Case CELLTYPE_DAY, CELLTYPE_WEEK
            
            wkTmp = 0
            For i = 0 To prpt.DrawColumns - 1
                wkDate = prpt.getDate(i)
                With sht.Cells(Row - 1, col + i)
                    If wkTmp <> Month(wkDate) Then
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Value = "'" & Year(wkDate) & "年" & Month(wkDate) & "月"
                    Else
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Value = ""
                    End If
                End With
                wkTmp = Month(wkDate)
            Next i
        
        '10日
        Case CELLTYPE_10DAYS
            
            For i = 0 To prpt.DrawColumns - 1
                sht.Cells(Row, col + i).Value = prpt.getDate(i)
            Next i
            
        '月
        Case CELLTYPE_MONTH
            
            wkTmp = 0
            For i = 0 To prpt.DrawColumns - 1
                wkDate = prpt.getDate(i)
                With sht.Cells(Row - 1, col + i)
                    If wkTmp <> Year(wkDate) Then
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Value = Year(wkDate) & "年"
                    Else
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Value = ""
                    End If
                End With
                sht.Cells(Row, col + i).Value = prpt.getDate(i)
                wkTmp = Year(wkDate)
            Next i
            
    End Select
    
End Sub

'初期化
Public Function Initialize(ChartSheet As Worksheet)
    
    Application.StatusBar = "データ読込"
    
    'シート設定
    Set sht = ChartSheet
    
    'プロパティ読込
    Call LoadProperty
    'MsgBox "clsGanttChart.Initialize Property Loaded..."
    
    'データ読込
    Call LoadData
    
    If prpt.isGroupingByCategory Then
        Call MargeDataByCategory
    Else
        Call MargeDataByPerson
    End If
    
End Function

Private Sub SetStatusString(Caption As String, Status As Single)
    
    Dim i As Integer
    Dim msg As String
    
    msg = Caption & " "
    Status = Status * 10
    
    For i = 1 To Int(Status)
        
        msg = msg & "■"
        
    Next
    
    For i = Int(Status) To 10
        
        msg = msg & "□"
        
    Next
    
    Application.StatusBar = msg
    
End Sub
