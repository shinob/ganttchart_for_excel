VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProperty 
   Caption         =   "Property"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmProperty.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'プロパティ編集用フォーム
Option Explicit
Option Base 1

Private prpt As New clsProperty '個別プロパティ
Private def As New clsProperty  '標準プロパティ

Public flgInit As Boolean

Public flgEdit As Boolean

'インターフェース設定
Private Sub InitInterFace()
    
    With Me
        
        .chkDefault.Caption = "標準設定を使用"
        .Label1.Caption = "シート名称"
        .Label2.Caption = "工程開始日"
        .Label3.Caption = "セル単位"
        .Label4.Caption = "描画列数"
        .Label5.Caption = "描画行数"
        .Label6.Caption = "描画開始列"
        .Label7.Caption = "描画開始行"
        
        .chkAutoUpdate.Caption = "自動更新"
        
        btnBeginDate.Caption = "選択"
        
        .Label14.Caption = "チャート太さ"
        .Label15.Caption = "予定描画位置[%]"
        .Label16.Caption = "実績描画位置[%]"
        .Label17.Caption = "線の種類"
        
        .Label18.Caption = "予定の線の色"
        .Label19.Caption = "予定の塗り色"
        .Label21.Caption = "実績の線の色"
        .Label20.Caption = "実績の塗り色"
        
        .Label27.Caption = "休日の色"
        .Label31.Caption = "完了項目の背景色"
        
        .lblPlanLineColor.Caption = ""
        .lblPlanFillColor.Caption = ""
        .lblActLineColor.Caption = ""
        .lblActFillColor.Caption = ""
        
        .lblHoliday.Caption = ""
        .lblComplete.Caption = ""
        
        .Label28.Caption = "イナズマ線"
        
        .chkInazumaDraw.Caption = "イナズマ線の描画"
        .Label24.Caption = "イナズマ線の色"
        .Label25.Caption = "イナズマ線の幅"
        .lblInazumaColor.Caption = ""
        
        .Label29.Caption = "マイルストーン"
        .Label30.Caption = "線幅"
        
        .Label22.Caption = "チャートの分類方法"
        .Label23.Caption = "チャートの整列方法"
        
        .chkVisibleAll.Caption = "期間外の項目も表示する"
        .chkVisibleNotComplete.Caption = "期間外の未完の項目も表示する"
        .chkInsertBlankRow.Caption = "分類間に空白行を挿入する"
        .chkPrintLabel.Caption = "工程名称をラベルとして表示する"
        .chkCalcurateStatus.Caption = "進捗状況を実績から計算する"
        .chkWorkDaysExceptHoliday.Caption = "休日を除いて工数を計算する"
        .chkPaintComplete.Caption = "完了項目の背景を塗りつぶす"
        .chkUseMemo.Caption = "追加の備考を利用する"
        .chkItemNo.Caption = "項目NOを出力しない"
        
        .MultiPage1.Page1.Caption = "描画位置"
        .MultiPage1.Page2.Caption = "工程調整"
        .MultiPage1.Page3.Caption = "色設定"
        .MultiPage1.Page4.Caption = "イナズマ線"
        .MultiPage1.Page5.Caption = "描画設定"
        
        .btnSet.Caption = "設定"
        .btnCancel.Caption = "取消"
        
        .btnImport.Caption = "データ取込"
        .btnMerge.Caption = "データ結合"
        
        .Caption = "プロパティ編集"
        
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

'フォームに値を設定
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
