Attribute VB_Name = "mdlMain"
Option Explicit
Option Base 1

'�萔�̒�`
Public Const FIRSTDATE = #1/1/2001#
Public Const LANGUAGE = "Japanese"
Public Const CHARTBAR = "ChartBar"
Public Const EDITBOX = "EditBox"
Public Const DEFAULT_PROPERTY = 3

Public Const CELLTYPE_TIME = "����"
Public Const CELLTYPE_DAY = "��"
Public Const CELLTYPE_WEEK = "7��"
Public Const CELLTYPE_10DAYS = "10��"
Public Const CELLTYPE_MONTH = "��"

Public Const PROPTYPE_DEF = "�W��"
Public Const PROPTYPE_CST = "��"

Public Const CHARTBARTYPE_RECT = "�����`"
Public Const CHARTBARTYPE_LINE = "����"

Public Const GROUPTYPE_CATEGORY = "����"
Public Const GROUPTYPE_PERSON = "�S��"

Public Const SORTTYPE_DATE = "���t��"
Public Const SORTTYPE_DATA = "�o�^��"

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

Private Const MSG_IMPOK = "�f�[�^�捞����"
Private Const MSG_IMPNG = "�f�[�^�捞���s"
Private Const CALENDAR = "���t�ύX"
Private Const HOLIDAY = "�x���ݒ�"

Public Sub test()
    Dim frm As New frmDataManager
    frm.Show
End Sub

'�}�C���X�g�[���ҏW
Public Sub EditMilestone()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmMilestone
    
    With frm
        
        .Show
        
        If .flgEdit Then Call UpdateChart
        
    End With
    
End Sub

'���ވړ�
Public Sub MoveCategory()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmMoveCategory
    
    With frm
        
        .Show
        
        If .flgEdit Then Call UpdateChart
        
    End With
    
End Sub

'�f�[�^�捞
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

'�f�[�^����
Public Sub MergeChartData()
    
    Dim wk As New clsImport
    
    Call wk.Merge
    
    Select Case wk.Status
        
        Case IMPORT_CANCEL
        
        Case IMPORT_FALSE
            MsgBox "���s"
        Case IMPORT_TRUE
            MsgBox "����"
            
    End Select
    
End Sub

'���j���[�o�[�\��
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

'���j���[�o�[��\��
Public Sub HideMenuBar()
    
    If flgClose Then Exit Sub
    
    If Not IsSet(MenuBar) Then
        Set MenuBar = New clsMenuBar
    End If
    
    MenuBar.Hide
    
End Sub

'���j���[�o�[�폜
Public Sub DelMenuBar()

    If Not IsSet(MenuBar) Then Set MenuBar = New clsMenuBar
    Set MenuBar = Nothing
    
End Sub

'�I��
Public Sub Quit()
    
    ThisWorkbook.Close
    
End Sub

'�H���\�X�V
Public Sub UpdateChart()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    If 1 < ActiveWindow.SelectedSheets.Count Then
        MsgBox "�����̃V�[�g���I������Ă��邽�߁A�������p���ł��܂���"
        Exit Sub
    End If
    
    Dim s As New clsGanttChart
    
    Dim buf As Range
    Dim zoom As Integer
    
    Set buf = ActiveCell
    zoom = ActiveWindow.zoom
    
    Application.ScreenUpdating = False
    ActiveWindow.zoom = 100
    Application.StatusBar = "�`��J�n"
    
    Call s.Initialize(ActiveSheet)
    Call s.DrawChart
    
    buf.Select
    Application.StatusBar = False
    ActiveWindow.zoom = zoom
    Application.ScreenUpdating = True
    
End Sub

'���ڕҏW
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

'���ڕ���
Public Sub ShowDataManager()
    
    Dim frm As New frmDataManager
    
    With frm
        .Show
    End With
    
End Sub

'�H���ҏW
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

'�J�����_�[�\��
Public Sub ShowCalendarForm()
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = HOLIDAY
        .mode = ""
        .Show
        
    End With
    
End Sub

'���ڕҏW
Public Sub ShowItemForm()
    
    Call EditItem(CLng(Right(Application.Caller, 5)))
    
End Sub

'�V�K���ڕҏW
Public Sub ShowNewItemForm()
    
    Call EditItem(0)
    
End Sub

'�s���ҏW
Public Sub ShowScheduleForm()
    
    Call EditSchedule(CLng(Right(Application.Caller, 5)))
    
End Sub

'���t�ύX
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

'�x���ҏW
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

'�v���p�e�B�ҏW
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

'�C�i�Y�}���ҏW
Public Sub EditInazuma()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    Dim frm As New frmInazuma
    
    With frm
        
        .Show
        
        If .flgEdit Then UpdateChart
        
    End With
    
End Sub

'�����X�V
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

'�H���\�o��
Public Sub ChartSheetExport()
    
    If Not isChartSheet(ActiveSheet) Then Exit Sub
    
    ActiveSheet.Copy
    
    Dim s As Shape
    
    For Each s In ActiveSheet.Shapes
        
        s.OnAction = ""
        If s.Name Like EDITBOX & "*" Then s.Visible = msoFalse
        
    Next s
    
End Sub
