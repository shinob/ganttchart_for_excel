VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItem 
   Caption         =   "Item"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   OleObjectBlob   =   "frmItem.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'項目編集用フォーム
Option Explicit
Option Base 1

Public itm As New clsItem   '項目
Public flgEdit As Boolean  '編集フラグ

Private ScheduleDateBuffer As String    '工程用日付バッファ

Private ActiveSchedule As Long      '選択中の工程#
Private flgEditSchedule As Boolean  '工程編集フラグ

Private flgActivate As Boolean      '表示済フラグ

Private VALUES_ERROR As String
Private Const MSG_NONAME = "項目の名称が設定されていません"
Private Const MSG_SETSCHEDULE = "工程が編集されています。" & vbCr & "変更を反映させますか?"

Private Const SDL_NEW = "新規登録"
Private Const SDL_NONAME = "[名称未設定]"
Private Const SDL_DEL = "[削除]"
Private Const SDL_DATASELECT = "日付選択"

'インターフェース設定
Private Sub InitInterFace()

    Label1.Caption = "項目名"
    Label2.Caption = "備考"
    Label3.Caption = "リンク先"
    Label4.Caption = "進捗状況"
    
    Me.frmCategory.Caption = "分類"
    Me.frmPerson.Caption = "担当"
    Me.frmSchedule.Caption = "工程"
    
    btnCategory.Caption = "編集"
    btnPerson.Caption = "編集"
    
    chkComplete.Caption = "完了"
    
    btnFile.Caption = "ファイル選択"
    btnLink.Caption = "上位項目の設定"
    btnCalcActDates.Caption = "実績データ生成"
    
    btnEditMemos.Caption = "備考編集"
    
    btnSave.Caption = "保存"
    btnNew.Caption = "新規"
    btnDelete.Caption = "削除"
    btnCancel.Caption = "取消"
    
    Me.frmDetail.Caption = "詳細"
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
    btnDelSchedule.Caption = "削除"
    
    Me.Caption = "項目"
    
End Sub

'************************************************
' 項目関係
'************************************************

'読込
Public Sub Load(Num As Long)
    
    itm.Load Num
    itm.LoadSchedule
    
End Sub

Private Function CheckValues() As Boolean
    
    CheckValues = True
    
    If txtName.Text = "" Then
        CheckValues = False
        VALUES_ERROR = MSG_NONAME
    End If
    
End Function

'フォームの値を取得
Private Sub SetValuesFromForm()
    
    With itm
        
        .No = val(txtNo.Text)
        .Name = txtName.Text
        .Category = cmbCategory.ListIndex + 2
        .Person = cmbPerson.ListIndex + 2
        .Note = txtNote.Text
        .Complete = chkComplete.Value
        .Status = scbStatus.Value
        .Hyperlink = txtHyperlink.Text
        
        If 2 <= .LinkItem Then .Category = 0
        
    End With
    
End Sub

'フォームに値を設定
Private Sub SetValuesToForm()
    
    With itm
        
        txtNo.Text = .No
        txtName.Text = .Name
        
        If (2 < .Category) And (.Category < cmbCategory.ListCount + 2) Then
            cmbCategory.ListIndex = .Category - 2
        Else
            cmbCategory.ListIndex = 0
        End If
        
        'MsgBox .Person & vbCr & cmbPerson.ListCount + 2
        'MsgBox (2 < .Person)
        If (2 < .Person) And (.Person < cmbPerson.ListCount + 2) Then
            cmbPerson.ListIndex = .Person - 2
            'MsgBox "??"
        Else
            cmbPerson.ListIndex = 0
        End If
        
        txtNote.Text = .Note
        chkComplete.Value = .Complete
        scbStatus.Value = .Status
        txtHyperlink.Text = .Hyperlink
        
        If 2 <= .LinkItem Then frmCategory.Visible = False
        
        
    End With
    
    flgEdit = False
    
End Sub

'空データの設定
Private Sub SetValuesAsEmpty()
    
    Dim i As Long
    Dim wkSchedule As clsSchedules
    
    With itm.Schedules
    
        Dim cnt As Long
        
        cnt = .Count()
        
        For i = 1 To cnt
            
            .Items(i).Delete = True
            
        Next i
    
        Call SaveSchedule
    
    End With
    
    i = itm.No
    
    Set itm = Nothing
    Set itm = New clsItem
    
    itm.No = i
    itm.Save
    flgEdit = True
    
    Me.Hide
    
End Sub

'保存
Private Sub SaveAndExit()
    
    Call SetValuesFromForm
    itm.Save
    
    Call ConfirmUpdateSchedule
    Call SaveSchedule
    
    flgEdit = True
    Me.Hide
    
End Sub

'リストの更新
Private Sub UpdateList(obj As Object, cmb As Control)
    
    Dim i As Long
    Dim cnt As Long
    
    cnt = obj.Count()
    
    cmb.Clear
    
    For i = 1 To cnt
        
        cmb.AddItem obj.Items(i).Name
        
    Next i
    
    
End Sub

'分類一覧の更新
Private Sub UpdateCategoryList()
    Dim obj As New clsCategorys
    Call UpdateList(obj, cmbCategory)
End Sub

'担当一覧の更新
Private Sub UpdatePersonList()
    Dim obj As New clsPersons
    Call UpdateList(obj, cmbPerson)
End Sub

Private Sub btnCalcActDates_Click()
    
    If flgEditSchedule Then
        
        MsgBox "工程を編集中です。" & vbCr & "設定後に再度操作して下さい。"
        
    Else
        
        If MsgBox("実績データが変更されます" & vbCr & "よろしいですか？", _
            vbYesNo + vbInformation, "実績データの生成") = vbYes Then
            
            Call itm.Schedules.SetActionDateByStatus(scbStatus.Value)
            
        End If
        
    End If
    
End Sub

Private Sub btnEditMemos_Click()
    
    Dim wk As New frmMemos
    
    Call wk.Load(itm.Memo)
    wk.Show
    
    Set wk = Nothing
    
End Sub

Private Sub btnFile_Click()
    
    Dim tmp As String
    
    tmp = Application.GetOpenFilename()
    
    If tmp <> "" And tmp <> "False" Then
        
        txtHyperlink.Text = tmp
        
    End If
    
End Sub

Private Sub btnLink_Click()
    
    Dim frm As New frmItemSelect
    
    With frm
        
        .ActiveItem = itm.No
        .Show
        
        Select Case .mode
            
            Case ITEMSELECT_LINK
                
                If itm.No <> .SelectedItem Then
                    
                    itm.LinkItem = .SelectedItem
                    frmCategory.Visible = False
                    flgEdit = True
                    
                Else
                
                End If
                
            Case ITEMSELECT_UNLINK
                
                If itm.LinkItem <> 0 Then
                    itm.LinkItem = 0
                    frmCategory.Visible = True
                    flgEdit = True
                End If
                
            Case ITEMSELECT_CANCEL
            
        End Select
        
    End With
    
End Sub

Private Sub lblWeight_Click()
    
    Call mdlTools.SetSpinButtonValue(spbWeight)
    
End Sub

Private Sub lstChartType_Change()
    flgEditSchedule = True
End Sub

'進捗状況設定
Private Sub scbStatus_Change()
    
    lblStatus.Caption = scbStatus.Value & "%"
    flgEdit = True
    
End Sub

'取消操作
Private Sub btnCancel_Click()
    flgEdit = False
    Me.Hide
End Sub

'分類編集
Private Sub btnCategory_Click()
    
    Dim frm As New frmCategory
    
    frm.Show
    
    Dim i As Long
    
    i = cmbCategory.ListIndex
    Call UpdateCategoryList
    cmbCategory.ListIndex = i
    
End Sub

'担当編集
Private Sub btnPerson_Click()

    Dim frm As New frmPerson
    frm.Show
    
    Dim i As Long
    
    i = cmbPerson.ListIndex
    Call UpdatePersonList
    cmbPerson.ListIndex = i
    
End Sub

Private Sub spbWeight_Change()
    lblWeight.Caption = spbWeight.Value
    flgEditSchedule = True
End Sub

'名称の編集
Private Sub txtName_Change(): flgEdit = True: End Sub
Private Sub cmbCategory_Change(): flgEdit = True: End Sub
Private Sub cmbPerson_Change(): flgEdit = True: End Sub
Private Sub txtNote_Change(): flgEdit = True: End Sub
Private Sub txtHyperlink_Change(): flgEdit = True: End Sub
Private Sub chkComplete_Change(): flgEdit = True: End Sub

'新規保存
Private Sub btnNew_Click()
    
    If Not CheckValues Then
        
        MsgBox VALUES_ERROR
        Exit Sub
        
    End If
    
    Dim i As Long
    Dim cnt As Long
    
    cnt = itm.Schedules.Count()
    
    For i = 1 To cnt
        
        itm.Schedules.Items(i).No = 0
        
    Next i
    
    txtNo.Text = ""
    Call SaveAndExit
    
End Sub

'保存
Private Sub btnSave_Click()
    
    If CheckValues Then
        Call SaveAndExit
    Else
        MsgBox VALUES_ERROR
    End If
    
End Sub

'削除
Private Sub btnDelete_Click()
    
    Call PreparationForDelete
    Call SetValuesAsEmpty
    
End Sub

Private Sub PreparationForDelete()
    
    If itm.No < 2 Then Exit Sub
    
    Dim wkItems As New clsItems
    Dim i As Long
    
    wkItems.Load
    
    For i = 1 To wkItems.Count()
        
        With wkItems.Items(i)
            
            If itm.No = .LinkItem Then
                
                .LinkItem = itm.LinkItem
                .Category = itm.Category
                .Save
                
            End If
            
        End With
        
    Next i
    
End Sub

'************************************************
' 工程関係
'************************************************

'工程保存
Private Sub SaveSchedule()
    
    Dim i As Long
    Dim cnt As Long
    
    cnt = itm.Schedules.Count()
    
    For i = 1 To cnt
        
        With itm.Schedules.Items(i)
            
            .Item = itm.No
            .Save
            
        End With
        
    Next i
    
    'MsgBox "SaveSchedule"
    
End Sub

'フォームに工程データを設定
Private Sub SetScheduleValuesToForm(obj As clsSchedule)
    
    With obj
        
        txtScheduleNo.Text = .No
        txtSchedule.Text = .Name
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

'フォームの工程データを取得
Private Function CreateSchedule() As clsSchedule
    
    Dim sdl As New clsSchedule
    
    With sdl
        
        .No = val(txtScheduleNo.Text)
        .Name = txtSchedule.Text
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
    
    Set CreateSchedule = sdl
    
End Function

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

'工程リストの更新
Private Sub UpdateScheduleList()

    Dim i As Long
    Dim cnt As Long
    
    Dim Name As String
    
    With itm.Schedules
        
        .Sort
        
        cnt = .Count
        lstSchedule.Clear
        lstSchedule.AddItem SDL_NEW
        
        For i = 1 To cnt
            
            Name = .Items(i).Name
            
            If Name = "" Then Name = SDL_NONAME
            If .Items(i).Delete Then Name = Name & " " & SDL_DEL
            
            lstSchedule.AddItem Name
            
        Next i
        
    End With
    
End Sub

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
        
        .Caption = SDL_DATASELECT
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
        
    Call UpdateSchedule
        
End Sub

'工程更新
Private Sub UpdateSchedule()

    Dim buf As clsSchedule
    
    Dim i As Long
    
    'i = lstSchedule.ListIndex
    i = ActiveSchedule
    Set buf = CreateSchedule()
    
    With itm.Schedules
    
        If i < 1 Then
            .Add
            .Items(.Count) = buf
        Else
            .Items(i) = Nothing
            .Items(i) = buf
        End If
        
    End With
    
    Call UpdateScheduleList
    
    frmDetail.Visible = False
    flgEditSchedule = False
    flgEdit = True
    
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

'工程更新確認
Private Sub ConfirmUpdateSchedule()

    If flgEditSchedule Then
        
        If vbYes = MsgBox(MSG_SETSCHEDULE, vbYesNo) Then _
            Call UpdateSchedule
        
    End If
    
End Sub

'工程選択
Private Sub lstSchedule_Click()
    
    'MsgBox "lstSchedule_Click()"
    
    Dim obj As clsSchedule
    Dim i As Long
    
    Call ConfirmUpdateSchedule
    
    i = lstSchedule.ListIndex
    
    If i < 1 Then
        Set obj = New clsSchedule
        obj.PlanColor = -1
        obj.ActColor = -1
    Else
        Set obj = itm.Schedules.Items(i)
    End If
        
    Call SetScheduleValuesToForm(obj)
    ActiveSchedule = i
    
    frmDetail.Visible = True
    
End Sub

'工程削除
Private Sub btnDelSchedule_Click()
    
    Dim i As Long
    
    i = lstSchedule.ListIndex
    
    If i < 1 Then
    
    Else
        
        itm.Schedules.Items(i).Delete = True
        
    End If
    
    Call UpdateScheduleList
    
    frmDetail.Visible = False
    flgEdit = True
    
End Sub

'************************************************
' フォーム関係
'************************************************

'フォーム表示
Private Sub UserForm_Activate()

    If flgActivate = False Then
    
        Call SetValuesToForm
        Call UpdateScheduleList
        
        flgEditSchedule = False
        
        If 0 < itm.Schedules.Count Then
            lstSchedule.ListIndex = 1
            'Call SetScheduleValuesToForm(itm.Schedules.Items(1))
        End If
        
        flgActivate = True
        
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call UpdateCategoryList
    Call UpdatePersonList
    Call EnableEditDate(False)
    
    Call mdlTools.setFontOnForm(Me)
    
    frmDetail.Visible = False
    
    flgActivate = False
    
End Sub
