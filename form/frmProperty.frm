VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProperty 
   Caption         =   "Property"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmProperty.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�v���p�e�B�ҏW�p�t�H�[��
Option Explicit
Option Base 1

Private prpt As New clsProperty '�ʃv���p�e�B
Private def As New clsProperty  '�W���v���p�e�B

Public flgInit As Boolean

Public flgEdit As Boolean

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()
    
    With Me
        
        .chkDefault.Caption = "�W���ݒ���g�p"
        .Label1.Caption = "�V�[�g����"
        .Label2.Caption = "�H���J�n��"
        .Label3.Caption = "�Z���P��"
        .Label4.Caption = "�`���"
        .Label5.Caption = "�`��s��"
        .Label6.Caption = "�`��J�n��"
        .Label7.Caption = "�`��J�n�s"
        
        .chkAutoUpdate.Caption = "�����X�V"
        
        btnBeginDate.Caption = "�I��"
        
        .Label14.Caption = "�`���[�g����"
        .Label15.Caption = "�\��`��ʒu[%]"
        .Label16.Caption = "���ѕ`��ʒu[%]"
        .Label17.Caption = "���̎��"
        
        .Label18.Caption = "�\��̐��̐F"
        .Label19.Caption = "�\��̓h��F"
        .Label21.Caption = "���т̐��̐F"
        .Label20.Caption = "���т̓h��F"
        
        .Label27.Caption = "�x���̐F"
        .Label31.Caption = "�������ڂ̔w�i�F"
        
        .lblPlanLineColor.Caption = ""
        .lblPlanFillColor.Caption = ""
        .lblActLineColor.Caption = ""
        .lblActFillColor.Caption = ""
        
        .lblHoliday.Caption = ""
        .lblComplete.Caption = ""
        
        .Label28.Caption = "�C�i�Y�}��"
        
        .chkInazumaDraw.Caption = "�C�i�Y�}���̕`��"
        .Label24.Caption = "�C�i�Y�}���̐F"
        .Label25.Caption = "�C�i�Y�}���̕�"
        .lblInazumaColor.Caption = ""
        
        .Label29.Caption = "�}�C���X�g�[��"
        .Label30.Caption = "����"
        
        .Label22.Caption = "�`���[�g�̕��ޕ��@"
        .Label23.Caption = "�`���[�g�̐�����@"
        
        .chkVisibleAll.Caption = "���ԊO�̍��ڂ��\������"
        .chkVisibleNotComplete.Caption = "���ԊO�̖����̍��ڂ��\������"
        .chkInsertBlankRow.Caption = "���ފԂɋ󔒍s��}������"
        .chkPrintLabel.Caption = "�H�����̂����x���Ƃ��ĕ\������"
        .chkCalcurateStatus.Caption = "�i���󋵂����т���v�Z����"
        .chkWorkDaysExceptHoliday.Caption = "�x���������čH�����v�Z����"
        .chkPaintComplete.Caption = "�������ڂ̔w�i��h��Ԃ�"
        .chkUseMemo.Caption = "�ǉ��̔��l�𗘗p����"
        .chkItemNo.Caption = "����NO���o�͂��Ȃ�"
        
        .MultiPage1.Page1.Caption = "�`��ʒu"
        .MultiPage1.Page2.Caption = "�H������"
        .MultiPage1.Page3.Caption = "�F�ݒ�"
        .MultiPage1.Page4.Caption = "�C�i�Y�}��"
        .MultiPage1.Page5.Caption = "�`��ݒ�"
        
        .btnSet.Caption = "�ݒ�"
        .btnCancel.Caption = "���"
        
        .btnImport.Caption = "�f�[�^�捞"
        .btnMerge.Caption = "�f�[�^����"
        
        .Caption = "�v���p�e�B�ҏW"
        
    End With
    
End Sub

Private Sub getValuesOnForm(wkProperty As clsProperty)
    
    On Error Resume Next
    
    With wkProperty
        
        .RefTypeDefault = chkDefault.Value
        .SheetName = lblSheetName.Caption
        .BeginDate = CDate(txtBeginDate.Text)
        .CellType = cmbCellType.Text
        '.DrawColumns = CInt(lblDrawColumns.Caption)
        .DrawRows = scbDrawRows.Value
        .Top = CInt(lblTop.Caption)
        .Left = CInt(lblLeft.Caption)
        
        .AutoUpdate = chkAutoUpdate.Value
        
        .ChartWidth = scbChartWidth.Value / 10
        .PlanPosition = scbPlanPosition.Value
        .ActPosition = scbActPosition.Value
        
        .PlanLineColor = lblPlanLineColor.BackColor
        .PlanFillColor = lblPlanFillColor.BackColor
        .ActLineColor = lblActLineColor.BackColor
        .ActFillColor = lblActFillColor.BackColor
        
        .InazumaDraw = chkInazumaDraw.Value
        .InazumaWidth = scbInazumaWidth.Value
        .InazumaColor = lblInazumaColor.BackColor
        
        .MilestoneWidth = scbMilestoneWidth.Value
        
        .HOLIDAY.Interior.color = lblHoliday.BackColor
        .Complete.Interior.color = lblComplete.BackColor
        
        .GroupingType = cmbGroupingType.Value
        .SortType = cmbSortType.Value
        
        .VisibleAll = chkVisibleAll.Value
        .VisibleNotComplete = chkVisibleNotComplete.Value
        .InsertBlankRow = chkInsertBlankRow.Value
        .PrintLabel = chkPrintLabel.Value
        .CalcurateStatus = chkCalcurateStatus.Value
        .WorkDaysExceptHoliday = chkWorkDaysExceptHoliday.Value
        .PaintComplete = chkPaintComplete.Value
        .UseMemos = chkUseMemo.Value
        .ItemNo = chkItemNo.Value
        
    End With
    
End Sub

'�t�H�[���ɒl��ݒ�
Private Sub setValuesOnForm(wkProperty As clsProperty)
    
    With wkProperty
        
        chkDefault.Value = .IsDefault()
        lblSheetName.Caption = .SheetName
        txtBeginDate.Text = Format(.BeginDate, "yyyy/mm/dd")
        cmbCellType.Text = .CellType
        'lblDrawColumns.Caption = .DrawColumns & _
        '    "[ " & wkProperty.EndDate & " ]"
        scbDrawRows.Value = .DrawRows
        lblTop.Caption = .Top
        lblLeft.Caption = .Left
        
        chkAutoUpdate.Value = .AutoUpdate
        
        If chkDefault.Value Then
            Call setValuesOnFormSub(def)
        Else
            Call setValuesOnFormSub(wkProperty)
        End If
        
    End With
    
End Sub

Private Sub setValuesOnFormSub(wkProperty As clsProperty)

    With wkProperty
        
        scbChartWidth.Value = .ChartWidth * 10
        scbPlanPosition.Value = .PlanPosition
        scbActPosition.Value = .ActPosition
        
        lblPlanLineColor.BackColor = .PlanLineColor
        lblPlanFillColor.BackColor = .PlanFillColor
        lblActLineColor.BackColor = .ActLineColor
        lblActFillColor.BackColor = .ActFillColor
        
        lblHoliday.BackColor = .HOLIDAY.Interior.color
        lblComplete.BackColor = .Complete.Interior.color
        
        chkInazumaDraw.Value = .InazumaDraw
        scbInazumaWidth.Value = .InazumaWidth
        lblInazumaWidth.Caption = .InazumaWidth
        lblInazumaColor.BackColor = .InazumaColor
        
        scbMilestoneWidth.Value = .MilestoneWidth
        lblMilestoneWidth.Caption = .MilestoneWidth
        
        cmbGroupingType.Value = .GroupingType
        cmbSortType.Value = .SortType
        
        chkVisibleAll.Value = .VisibleAll
        chkVisibleNotComplete.Value = .VisibleNotComplete
        chkInsertBlankRow.Value = .InsertBlankRow
        chkPrintLabel.Value = .PrintLabel
        chkCalcurateStatus.Value = .CalcurateStatus
        chkWorkDaysExceptHoliday.Value = .WorkDaysExceptHoliday
        chkPaintComplete.Value = .PaintComplete
        chkUseMemo.Value = .UseMemos
        chkItemNo.Value = .ItemNo
        
        Call .MakeChartBarTypeList(lstChartBarType)
        
    End With
    
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnImport_Click()
    
    Call mdlMain.ImportChartData
    
End Sub

Private Sub btnMerge_Click()
    
    Call mdlMain.MergeChartData
    
End Sub

Private Sub btnSet_Click()
    Call getValuesOnForm(prpt)
    Call prpt.Save
    
    flgEdit = True
    Me.Hide
    
End Sub

Private Sub chkDefault_Click()
    
    If flgInit = False Then Exit Sub
    
    If chkDefault.Value Then
        Call getValuesOnForm(prpt)
        Call setValuesOnFormSub(def)
    Else
        Call getValuesOnForm(def)
        Call setValuesOnFormSub(prpt)
    End If
    
End Sub

Private Sub cmbCellType_Change()
    
    With prpt
        .CellType = cmbCellType.Value
        lblDrawColumns.Caption = .DrawColumns & " [ " & .EndDate & " ]"
    End With
    
End Sub

Private Sub lblHoliday_Click()
    Call mdlTools.EditColor(lblHoliday)
End Sub

Private Sub lblComplete_Click()
    Call mdlTools.EditColor(lblComplete)
End Sub

Private Sub lblInazumaColor_Click()
    Call mdlTools.EditColor(lblInazumaColor)
End Sub

Private Sub lblPlanLineColor_Click()
    Call mdlTools.EditColor(lblPlanLineColor)
End Sub
Private Sub lblPlanFillColor_Click()
    Call mdlTools.EditColor(lblPlanFillColor)
End Sub
Private Sub lblActLineColor_Click()
    Call mdlTools.EditColor(lblActLineColor)
End Sub
Private Sub lblActFillColor_Click()
    Call mdlTools.EditColor(lblActFillColor)
End Sub

Private Sub scbDrawRows_Change()
    
    lblDrawRows.Caption = scbDrawRows.Value
    
End Sub

Private Sub scbChartWidth_Change()
    
    lblChartWidth.Caption = Format(scbChartWidth.Value / 10, "0.0")
    
End Sub

Private Sub scbInazumaWidth_Change()
    lblInazumaWidth.Caption = scbInazumaWidth.Value
End Sub

Private Sub scbMilestoneWidth_Change()
    lblMilestoneWidth.Caption = scbMilestoneWidth.Value
End Sub

Private Sub scbPlanPosition_Change()
    
    lblPlanPosition.Caption = scbPlanPosition.Value
    
End Sub

Private Sub scbActPosition_Change()
    
    lblActPosition.Caption = scbActPosition.Value
    
End Sub

Public Sub Init(Num As Integer)

    With prpt
        Call .Load(Num)
        Call .MakeCellTypeList(cmbCellType)
        Call .MakeGroupingTypeList(cmbGroupingType)
        Call .MakeSortTypeList(cmbSortType)
    End With
    
    Call def.Load(DEFAULT_PROPERTY)
    
    Call setValuesOnForm(prpt)
    
    MultiPage1.Value = 0
    
    flgInit = True
    
End Sub

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call mdlTools.setFontOnForm(Me)
    
    flgInit = False
    
End Sub
