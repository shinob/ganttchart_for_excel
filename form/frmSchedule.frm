VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSchedule 
   Caption         =   "UserForm1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   OleObjectBlob   =   "frmSchedule.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'項目編集用フォーム
Option Explicit
Option Base 1

Private ScheduleDateBuffer As String    '工程用日付バッファ

Public flgEditSchedule As Boolean  '工程編集フラグ

Private Const CALENDAR = "日付選択"

'インターフェース設定
Private Sub InitInterFace()

    Label9.Caption = "名称"
    Label10.Caption = "予定開始"
    Label11.Caption = "予定終了"
    Label12.Caption = "実績開始"
    Label13.Caption = "実績終了"
    Label14.Caption = "人数"
    
    With lstChartType
        .AddItem "長方形"
        .AddItem "直線"
        '.AddItem = "点線"
    End With
    
    lblPlanBegin.Text = ""
    lblPlanEnd.Text = ""
    lblActBegin.Text = ""
    lblActEnd.Text = ""
    
    chkManual.Caption = "日付の手入力"
    
    btnSet.Caption = "設定"
    btnCancel.Caption = "取消"
    
    Me.Caption = "項目"
    
End Sub

'************************************************
' 工程関係
'************************************************
Public Sub Initialize(Num As Long)
    
    Dim sdl As New clsSchedule
    
    Call sdl.Load(Num)
    Call SetScheduleValuesToForm(sdl)
    
End Sub

'フォームに工程データを設定
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

'保存
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

'日付取得
Private Function getDateFromText(txt As String) As Date
    
    If txt = "" Then
        getDateFromText = 0
    Else
        getDateFromText = CDate(txt)
    End If
    
End Function

'チャート色の表示
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

'チャート色の取得
Private Function GetChartBarColor(img As Control) As Long
    
    If img.BackStyle = fmBackStyleTransparent Then
        GetChartBarColor = -1
    Else
        GetChartBarColor = img.BackColor
    End If
    
End Function

'日付フォーマット
Private Function DATEFORMAT(mDate As Date) As String
    
    If mDate < FIRSTDATE Then
        DATEFORMAT = ""
    Else
        DATEFORMAT = Format(mDate, "yyyy/mm/dd hh:mm")
    End If
    
End Function

'工程の日付編集方法選択
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

'日付ボタン操作
Private Sub btnPlanBegin_Click(): Call DateSelect("PlanBegin"): End Sub
Private Sub btnPlanEnd_Click(): Call DateSelect("PlanEnd"): End Sub
Private Sub btnActBegin_Click(): Call DateSelect("ActBegin"): End Sub
Private Sub btnActEnd_Click(): Call DateSelect("ActEnd"): End Sub

'日付選択
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

'工程内容設定
Private Sub btnSet_Click()
        
    Call Save
    flgEditSchedule = True
    Me.Hide
        
End Sub

'日付の手動設定
Private Sub chkManual_Click()
    
    Call EnableEditDate(chkManual.Value)
    
End Sub

'色設定
Private Sub imgPlanColor_Click(): Call mdlTools.EditColor(imgPlanColor): End Sub
Private Sub imgActColor_Click(): Call mdlTools.EditColor(imgActColor): End Sub

'設定済の日付を取得
Private Sub setScheduleDateBuffer(Value As String)
    ScheduleDateBuffer = Value
End Sub

'入力値が不適切な場合、値を復旧
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

'名前の編集
Private Sub txtSchedule_Change()
    flgEditSchedule = True
End Sub

'日付の手動編集
'予定開始
Private Sub lblPlanBegin_Enter(): Call setScheduleDateBuffer(lblPlanBegin.Text): End Sub
Private Sub lblPlanBegin_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblPlanBegin"): End Sub
Private Sub lblPlanBegin_Change(): flgEditSchedule = True: End Sub
'予定終了
Private Sub lblPlanEnd_Enter(): Call setScheduleDateBuffer(lblPlanEnd.Text): End Sub
Private Sub lblPlanEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblPlanEnd"): End Sub
Private Sub lblPlanEnd_Change(): flgEditSchedule = True: End Sub
'実績開始
Private Sub lblActBegin_Enter(): Call setScheduleDateBuffer(lblActBegin.Text): End Sub
Private Sub lblActBegin_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblActBegin"): End Sub
Private Sub lblActBegin_Change(): flgEditSchedule = True: End Sub
'実績終了
Private Sub lblActEnd_Enter(): Call setScheduleDateBuffer(lblActEnd.Text): End Sub
Private Sub lblActEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call checkScheduleDate("lblActEnd"): End Sub
Private Sub lblActEnd_Change(): flgEditSchedule = True: End Sub

'************************************************
' フォーム関係
'************************************************

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    
    Call mdlTools.setFontOnForm(Me)
    
End Sub
