Attribute VB_Name = "mdlMain"
Option Explicit
Option Base 1

'定数の定義
Public Const FIRSTDATE = #1/1/2001#
Public Const LANGUAGE = "Japanese"
Public Const CHARTBAR = "ChartBar"
Public Const EDITBOX = "EditBox"
Public Const DEFAULT_PROPERTY = 3

Public Const CELLTYPE_TIME = "時間"
Public Const CELLTYPE_DAY = "日"
Public Const CELLTYPE_WEEK = "7日"
Public Const CELLTYPE_10DAYS = "10日"
Public Const CELLTYPE_MONTH = "月"

Public Const PROPTYPE_DEF = "標準"
Public Const PROPTYPE_CST = "個別"

Public Const CHARTBARTYPE_RECT = "長方形"
Public Const CHARTBARTYPE_LINE = "直線"

Public Const GROUPTYPE_CATEGORY = "分類"
Public Const GROUPTYPE_PERSON = "担当"

Public Const SORTTYPE_DATE = "日付順"
Public Const SORTTYPE_DATA = "登録順"

Public Const IMPORT_CANCEL = -1
Public Const IMPORT_FALSE = 0
Public Const IMPORT_TRUE = 1

Public Const gntChartLine = 1
Public Const gntChartRect = 0

Public Const COUNT_INAZUMA = 20
Public Const COUNT_MEMO = 5

Public Const ITEMSELECT_LINK = 1
Public Const ITEMSELECT_UNLINK = -1
Public Const ITEMSELECT_CANCEL = 0

Public Const MODE_PLAN = 0
Public Const MODE_ACT = 1

Public MenuBar As clsMenuBar
Public flgClose As Boolean

Private Const MSG_IMPOK = "データ取込成功"
Private Const MSG_IMPNG = "データ取込失敗"
Private Const CALENDAR = "日付変更"
Private Const HOLIDAY = "休日設定"

Public Sub test()
    Dim frm As New frmDataManager
    frm.Show
End Sub

'マイルストーン編集
Public Sub EditMilestone()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmMilestone
    
    With frm
        
        .Show
        
        If .flgEdit Then Call UpdateChart
        
    End With
    
End Sub

'分類移動
Public Sub MoveCategory()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmMoveCategory
    
    With frm
        
        .Show
        
        If .flgEdit Then Call UpdateChart
        
    End With
    
End Sub

'データ取込
Public Sub ImportChartData()
    
    Dim wk As New clsImport
    
    Call wk.Copy
    
    Select Case wk.Status
        
        Case IMPORT_CANCEL
        
        Case IMPORT_FALSE
            MsgBox MSG_IMPNG
        Case IMPORT_TRUE
            MsgBox MSG_IMPOK
            
    End Select
    
End Sub

'データ統合
Public Sub MergeChartData()
    
    Dim wk As New clsImport
    
    Call wk.Merge
    
    Select Case wk.Status
        
        Case IMPORT_CANCEL
        
        Case IMPORT_FALSE
            MsgBox "失敗"
        Case IMPORT_TRUE
            MsgBox "成功"
            
    End Select
    
End Sub

'メニューバー表示
Public Sub ShowMenuBar()
    
    If Not IsSet(MenuBar) Then
        Set MenuBar = New clsMenuBar
    End If
    
    If isChartSheet(ActiveSheet) Then
    
        MenuBar.CREATE
        
    Else
        
        MenuBar.Hide
        
    End If
    
End Sub

'メニューバー非表示
Public Sub HideMenuBar()
    
    If flgClose Then Exit Sub
    
    If Not IsSet(MenuBar) Then
        Set MenuBar = New clsMenuBar
    End If
    
    MenuBar.Hide
    
End Sub

'メニューバー削除
Public Sub DelMenuBar()

    If Not IsSet(MenuBar) Then Set MenuBar = New clsMenuBar
    Set MenuBar = Nothing
    
End Sub

'終了
Public Sub Quit()
    
    ThisWorkbook.Close
    
End Sub

'工程表更新
Public Sub UpdateChart()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    If 1 < ActiveWindow.SelectedSheets.Count Then
        MsgBox "複数のシートが選択されているため、処理を継続できません"
        Exit Sub
    End If
    
    Dim s As New clsGanttChart
    
    Dim buf As Range
    Dim zoom As Integer
    
    Set buf = ActiveCell
    zoom = ActiveWindow.zoom
    
    Application.ScreenUpdating = False
    ActiveWindow.zoom = 100
    Application.StatusBar = "描画開始"
    
    Call s.Initialize(ActiveSheet)
    Call s.DrawChart
    
    buf.Select
    Application.StatusBar = False
    ActiveWindow.zoom = zoom
    Application.ScreenUpdating = True
    
End Sub

'項目編集
Public Sub EditItem(Num As Long)

    Dim frm As New frmItem
    
    With frm
        
        If 2 <= Num Then Call .Load(Num)
        .Show
        
        If .flgEdit Then
            Call UpdateChart
        End If
        
    End With
    
End Sub

'項目並替
Public Sub ShowDataManager()
    
    Dim frm As New frmDataManager
    
    With frm
        .Show
    End With
    
End Sub

'工程編集
Public Sub EditSchedule(Num As Long)

    Dim frm As New frmSchedule
    
    With frm
        
        .Initialize Num
        .Show
        
        If .flgEditSchedule Then
            Call UpdateChart
        End If
        
    End With
    
End Sub

'カレンダー表示
Public Sub ShowCalendarForm()
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = HOLIDAY
        .mode = ""
        .Show
        
    End With
    
End Sub

'項目編集
Public Sub ShowItemForm()
    
    Call EditItem(CLng(Right(Application.Caller, 5)))
    
End Sub

'新規項目編集
Public Sub ShowNewItemForm()
    
    Call EditItem(0)
    
End Sub

'行程編集
Public Sub ShowScheduleForm()
    
    Call EditSchedule(CLng(Right(Application.Caller, 5)))
    
End Sub

'日付変更
Public Sub ChangeDate()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmCalendar
    Dim prt As New clsProperty
    Dim ChartNo As Integer
    
    ChartNo = val(Right(ActiveSheet.codeName, 2)) + 3
    Call prt.Load(CLng(ChartNo))
    
    With frm
    
        .Caption = CALENDAR
        .mode = "SELECT"
        Call .SetDate(prt.BeginDate)
        .Show
        
        If .flgSelect Then
            prt.BeginDate = .SelectedDate
            prt.Save
            Call UpdateChart
        End If
        
    End With
    
End Sub

'休日編集
Public Sub EditHoliday()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = HOLIDAY
        .Show
        
        If .flgSelect Then
            Call UpdateChart
        End If
        
    End With
    
End Sub

'プロパティ編集
Public Sub EditProperty()
    
    Dim i As Integer
    i = mdlTools.ChartNum(ActiveSheet)
    'MsgBox "mdlMain.EditProperty " & i
    If i < 0 Then Exit Sub
    
    Dim frm As New frmProperty
    
    With frm
        
        Call .Init(i)
        .Show
        
        If .flgEdit Then
            Call UpdateChart
        End If
        
    End With
    
End Sub

'イナズマ線編集
Public Sub EditInazuma()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmInazuma
    
    With frm
        
        .Show
        
        If .flgEdit Then UpdateChart
        
    End With
    
End Sub

'自動更新
Public Sub CheckAutoUpdate()
    
    Dim i As Integer
    i = mdlTools.ChartNum(ActiveSheet)
    If i < 0 Then Exit Sub
    
    Dim prpt As New clsProperty
    
    With prpt
        .Load (i)
        If .AutoUpdate Then
            Call UpdateChart
        End If
    End With
    
End Sub

'工程表出力
Public Sub ChartSheetExport()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    ActiveSheet.Copy
    
    Dim s As Shape
    
    For Each s In ActiveSheet.Shapes
        
        s.OnAction = ""
        If s.Name Like EDITBOX & "*" Then s.Visible = msoFalse
        
    Next s
    
End Sub
