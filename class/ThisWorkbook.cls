VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


'
Option Explicit
Option Base 1

Private Sub Workbook_Activate()
    
    'MsgBox "Workbook_Activate"
    mdlMain.ShowMenuBar
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    'MsgBox "Workbook_BeforeClose"
    mdlMain.DelMenuBar
    mdlMain.flgClose = True
    
End Sub

Private Sub Workbook_Deactivate()
    
    'MsgBox "Workbook_Deactivate"
    mdlMain.HideMenuBar
    
End Sub

'ブックを開いた時メニューバーを生成
Private Sub Workbook_Open()
    
    'MsgBox "Workbook_Open"
    mdlMain.ShowMenuBar
    
End Sub

'シートを選択した時、メニューバーを確認し、自動更新を処理する
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    'MsgBox "Workbook_SheetActivate"
    mdlMain.ShowMenuBar
    mdlMain.CheckAutoUpdate
    
End Sub
