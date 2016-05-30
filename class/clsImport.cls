VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'データ取り込み用マクロ
Option Explicit
Option Base 1

Private src As Workbook
Private dst As Workbook
Private sht As Worksheet

Public Status As Integer

'結合
Public Function Merge() As Boolean
    
    Status = -1
    Merge = False
    
    Set sht = ActiveSheet
    
    Status = 0
    Application.ScreenUpdating = False
    
    If Not SelectFile Then
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    Dim sSht As Worksheet
    Dim dSht As Worksheet
    
    Dim ItemRow As Long
    Dim i As Long
    Dim j As Long
    
    '項目の取込
    Set sSht = getSheetObject(src, "shtItem")
    Set dSht = getSheetObject(dst, "shtItem")
    
    i = dSht.Cells.SpecialCells(xlCellTypeLastCell).Row
    If i < 2 Then
        i = 1
    Else
        Do While dSht.Cells(i, 1) = "" Or dSht.Cells(i, 2) = ""
            i = i - 1
        Loop
        i = i + 1
    End If
    ItemRow = i
    
    j = 2
    Do While sSht.Cells(j, 1) <> ""
        
        sSht.Activate
        sSht.Rows(j).Copy
        
        With dSht
            .Activate
            .Cells(i + j - 2, 1).Select
            .Paste
            
            .Cells(i + j - 2, 1) = i + j - 2
            If .Cells(i + j - 2, 9) > 0 Then _
                .Cells(i + j - 2, 9) = .Cells(i + j - 2, 9) + i - 2
                
        End With
        
        Application.CutCopyMode = False
        
        j = j + 1
        
    Loop
    
    '行程の取込
    Set sSht = getSheetObject(src, "shtSchedule")
    Set dSht = getSheetObject(dst, "shtSchedule")
    
    i = dSht.Cells.SpecialCells(xlCellTypeLastCell).Row
    If i < 2 Then
        i = 1
    Else
        Do While dSht.Cells(i, 1) = "" Or dSht.Cells(i, 3) = ""
            i = i - 1
        Loop
        i = i + 1
    End If
    
    j = 2
    Do While sSht.Cells(j, 1) <> ""
        
        sSht.Activate
        sSht.Rows(j).Copy
        
        With dSht
            .Activate
            .Cells(i + j - 2, 1).Select
            .Paste
            
            .Cells(i + j - 2, 1) = i + j - 2
            If .Cells(i + j - 2, 3) <> "" Then _
                If .Cells(i + j - 2, 3) > 0 Then _
                    .Cells(i + j - 2, 3) = .Cells(i + j - 2, 3) + ItemRow - 2
            
        End With
        
        Application.CutCopyMode = False
        
        j = j + 1
        
    Loop
    'MsgBox dSht.Name
    
    src.Close savechanges:=False
    Status = 1
    sht.Activate
    Application.ScreenUpdating = True
    
    Merge = True
    
End Function

'取込
Public Function Copy() As Boolean
    
    Status = -1
    Copy = False
    
    Set sht = ActiveSheet
    
    Status = 0
    Application.ScreenUpdating = False
    
    If Not SelectFile Then
        Application.ScreenUpdating = True
        Exit Function
    End If
    
    Call CopyData("shtCategory")
    Call CopyData("shtItem")
    Call CopyData("shtSchedule")
    Call CopyData("shtPerson")
    Call CopyData("shtHoliday")
    Call CopyData("shtMilestone")
    
    src.Close savechanges:=False
    Status = 1
    sht.Activate
    Application.ScreenUpdating = True
    
    Copy = True
    
End Function

'ファイル選択
Private Function SelectFile() As Boolean
    
    Dim wkFile As String
    
    wkFile = Application.GetOpenFilename()
    
    If wkFile = "" Then
        Exit Function
        
    ElseIf Not wkFile Like "*.xls" Then
        MsgBox "Excelのファイルではありません"
        SelectFile = False
        Exit Function
    End If
    
    Set src = Workbooks.Open(FileName:=wkFile, ReadOnly:=True)
    Set dst = ThisWorkbook
    
    SelectFile = True
    
End Function

'シート取得
Private Function getSheetObject(wkBook As Workbook, codeName As String) As Worksheet
    
    Dim s As Worksheet
    
    For Each s In wkBook.Worksheets
        
        If s.codeName = codeName Then
            
            Set getSheetObject = s
            Exit Function
            
        End If
        
    Next s
    
End Function

'シート内データの複製
Private Sub CopyData(codeName As String)
    
    On Error GoTo FIN
    
    Dim sSheet As Worksheet
    Dim dSheet As Worksheet
    
    Set sSheet = getSheetObject(src, codeName)
    Set dSheet = getSheetObject(dst, codeName)
    
    src.Activate
    sSheet.UsedRange.Copy
    
    dst.Activate
    dSheet.Select
    dSheet.Cells(1, 1).Select
    dSheet.Paste
    dSheet.Cells(1, 1).Select
    
    Application.CutCopyMode = False
    
FIN:
    
End Sub
