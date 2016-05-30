VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSchedule 
   Caption         =   "UserForm1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   OleObjectBlob   =   "frmSchedule.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���ڕҏW�p�t�H�[��
Option Explicit
Option Base 1

Private ScheduleDateBuffer As String    '�H���p���t�o�b�t�@

Public flgEditSchedule As Boolean  '�H���ҏW�t���O

Private Const CALENDAR = "���t�I��"

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()

    Label9.Caption = "����"
    Label10.Caption = "�\��J�n"
    Label11.Caption = "�\��I��"
    Label12.Caption = "���ъJ�n"
    Label13.Caption = "���яI��"
    Label14.Caption = "�l��"
    
    With lstChartType
        .AddItem "�����`"
        .AddItem "����"
        '.AddItem = "�_��"
    End With
    
    lblPlanBegin.Text = ""
    lblPlanEnd.Text = ""
    lblActBegin.Text = ""
    lblActEnd.Text = ""
    
    chkManual.Caption = "���t�̎����"
    
    btnSet.Caption = "�ݒ�"
    btnCancel.Caption = "���"
    
    Me.Caption = "����"
    
End Sub

'************************************************
' �H���֌W
'************************************************
Public Sub Initialize(Num As Long)
    
    Dim sdl As New clsSchedule
    
    Call sdl.Load(Num)
    Call SetScheduleValuesToForm(sdl)
    
End Sub

'�t�H�[���ɍH���f�[�^��ݒ�
Private Sub SetScheduleValuesToForm(obj As clsSchedule)
    
    With obj
        
        txtScheduleNo.Text = .No
        txtSchedule.Text = .Name
        txtItem.Text = .Item
        lblPlanBegin.Text = DATEFORMAT(.PlanBegin)
        lblPlanEnd.Text = DATEFORMAT(.PlanEnd)
        lblActBegin.Text = DATEFORMAT(.ActBegin)
        lblActEnd.Text = DATEFORMAT(.ActEnd)
        
        Call SetChartBarColor(imgPlanColor, .PlanColor)
        Call SetChartBarColor(imgActColor, .ActColor)
        
        lstChartType.ListIndex = .ChartType
        spbWeight.Value = .Weight
        
    End With
    
    flgEditSchedule = False
    
End Sub

'�ۑ�
Private Sub Save()
    
    Dim sdl As New clsSchedule
    
    With sdl
        
        .No = val(txtScheduleNo.Text)
        .Name = txtSchedule.Text
        .Item = val(txtItem.Text)
        .PlanBegin = getDateFromText(lblPlanBegin.Text)
        .PlanEnd = getDateFromText(lblPlanEnd.Text)
        .ActBegin = getDateFromText(lblActBegin.Text)
        .ActEnd = getDateFromText(lblActEnd.Text)
        
        .PlanColor = GetChartBarColor(imgPlanColor)
        .ActColor = GetChartBarColor(imgActColor)
        
        .ChartType = lstChartType.ListIndex
        .Weight = spbWeight.Value
        
        .Delete = False
        
    End With
    
    sdl.Save
    
End Sub

'���t�擾
Private Function getDateFromText(txt As String) As Date
    
    If txt = "" Then
        getDateFromText = 0
    Else
        getDateFromText = CDate(txt)
    End If
    
End Function

'�`���[�g�F�̕\��
Private Sub SetChartBarColor(img As Control, color As Long)
    
    If color = -2 Then Exit Sub
    
    If color < 0 Then
        img.BackStyle = fmBackStyleTransparent
    Else
        img.BackStyle = fmBackStyleOpaque
        img.BackColor = color
    End If
    
    'MsgBox img.BackStyle
    
End Sub

'�`���[�g�F�̎擾
Private Function GetChartBarColor(img As Control) As Long
    
    If img.BackStyle = fmBackStyleTransparent Then
        GetChartBarColor = -1
    Else
        GetChartBarColor = img.BackColor
    End If
    
End Function

'���t�t�H�[�}�b�g
Private Function DATEFORMAT(mDate As Date) As String
    
    If mDate < FIRSTDATE Then
        DATEFORMAT = ""
    Else
        DATEFORMAT = Format(mDate, "yyyy/mm/dd hh:mm")
    End If
    
End Function

'�H���̓��t�ҏW���@�I��
Private Sub EnableEditDate(mode As Boolean)
    
    'lblPlanBegin.Enabled = mode
    'lblPlanEnd.Enabled = mode
    'lblActBegin.Enabled = mode
    'lblActEnd.Enabled = mode
    
    btnPlanBegin.Visible = Not mode
    btnPlanEnd.Visible = Not mode
    btnActBegin.Visible = Not mode
    btnActEnd.Visible = Not mode
    
End Sub

Private Sub btnCancel_Click()
    flgEditSchedule = False
    Me.Hide
End Sub

'���t�{�^������
Private Sub btnPlanBegin_Click(): Call DateSelect("PlanBegin"): End Sub
Private Sub btnPlanEnd_Click(): Call DateSelect("PlanEnd"): End Sub
Private Sub btnActBegin_Click(): Call DateSelect("ActBegin"): End Sub
Private Sub btnActEnd_Click(): Call DateSelect("ActEnd"): End Sub

'���t�I��
Private Function DateSelect(mode As String) As String
    
    Dim btn As Control
    Dim ActiveDate As Date
    
    Set btn = Controls("lbl" & mode)
    
    ActiveDate = getDateFromText(btn.Text)
    
    Dim frm As New frmCalendar
    
    With frm
        
        .Caption = CALENDAR
        .mode = "SELECT"
        Call .SetDate(ActiveDate)
        .Show
        
        If .flgSelect Then
            ActiveDate = .SelectedDate
            If mode Like "*End" Then
                ActiveDate = ActiveDate + #11:59:59 PM#
            End If
            btn.Text = DATEFORMAT(ActiveDate)
        Else
        
        End If
        
    End With
    
    'DateSelect = DateFormat(ActiveDate)
    
End Function

'�H�����e�ݒ�
Private Sub btnSet_Click()
        
    Call Save
    flgEditSchedule = True
    Me.Hide
        
End Sub

'���t�̎蓮�ݒ�
Private Sub chkManual_Click()
    
    Call EnableEditDate(chkManual.Value)
    
End Sub

'�F�ݒ�
Private Sub imgPlanColor_Click(): Call mdlTools.EditColor(imgPlanColor): End Sub
Private Sub imgActColor_Click(): Call mdlTools.EditColor(imgActColor): End Sub

'�ݒ�ς̓��t���擾
Private Sub setScheduleDateBuffer(Value As String)
    ScheduleDateBuffer = Value
End Sub

'���͒l���s�K�؂ȏꍇ�A�l�𕜋�
Private Sub checkScheduleDate(ControlName As String)

    On Error GoTo RESTORE
    
    Dim obj As Control
    Dim tmp As Date
    
    'MsgBox "checkScheduleData"
    Set obj = Controls(ControlName)
    
    If obj.Text <> "" Then
    
        tmp = CDate(obj.Text)
        obj.Text = DATEFORMAT(tmp)
        
    End If
    
    Exit Sub
    
RESTORE:
    
    obj.Text = ScheduleDateBuffer
    
End Sub

Private Sub lblWeight_Click()
    
    Call mdlTools.SetSpinButtonValue(spbWeight)
    
End Sub

Private Sub lstChartType_Change()
    flgEditSchedule = True
End Sub

Private Sub spbWeight_Change()
    lblWeight.Caption = spbWeight.Value
    flgEditSchedule = True
End Sub

'���O�̕ҏW
Private Sub txtSchedule_Change()
    flgEditSchedule = True
End Sub

'���t�̎蓮�ҏW
'�\��J�n
Private Sub lblPlanBegin_Enter(): Call setScheduleDateBuffer(lblPlanBegin.Text): End Sub
Private Sub lblPlanBegin_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblPlanBegin"): End Sub
Private Sub lblPlanBegin_Change(): flgEditSchedule = True: End Sub
'�\��I��
Private Sub lblPlanEnd_Enter(): Call setScheduleDateBuffer(lblPlanEnd.Text): End Sub
Private Sub lblPlanEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblPlanEnd"): End Sub
Private Sub lblPlanEnd_Change(): flgEditSchedule = True: End Sub
'���ъJ�n
Private Sub lblActBegin_Enter(): Call setScheduleDateBuffer(lblActBegin.Text): End Sub
Private Sub lblActBegin_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblActBegin"): End Sub
Private Sub lblActBegin_Change(): flgEditSchedule = True: End Sub
'���яI��
Private Sub lblActEnd_Enter(): Call setScheduleDateBuffer(lblActEnd.Text): End Sub
Private Sub lblActEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblActEnd"): End Sub
Private Sub lblActEnd_Change(): flgEditSchedule = True: End Sub

'************************************************
' �t�H�[���֌W
'************************************************

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    
    Call mdlTools.setFontOnForm(Me)
    
End Sub
