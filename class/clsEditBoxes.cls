VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEditBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'編集ボタン作成クラス
Option Explicit
Option Base 1

Private Boxes() As Shape
Private cnt As Long
Private sht As Worksheet

'初期化
Public Sub Initialize(Num As Long)
'num : 必要なEditBoxの数 = 項目数

    Dim s As Shape
    Dim i As Long
    Dim j As Long
    
    '描画済を取得
    cnt = 0
    For Each s In sht.Shapes
        
        If s.Name Like EDITBOX & "*" Then
            
            cnt = cnt + 1
            
            ReDim Preserve Boxes(cnt)
            Set Boxes(cnt) = s
            Boxes(cnt).Name = EDITBOX & Format(cnt + 1, "00000")
            Boxes(cnt).Visible = msoFalse
            
        End If
        
    Next s
    
    If Num <= cnt Then
        GoTo FIN
    End If
    
    '不足分を描画
    j = Num - cnt
    For i = 1 To j
        
        cnt = cnt + 1
        
        ReDim Preserve Boxes(cnt)
        Set Boxes(cnt) = sht.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
        Boxes(cnt).Name = EDITBOX & Format(cnt + 1, "00000")
        Boxes(cnt).Visible = msoFalse
        
    Next i
    
FIN:
    '個数を初期値に設定
    'MsgBox "clsEditBoxes.Initialize cnt = " & cnt
    cnt = 1
    
End Sub

'編集ボタン取得
Public Property Get EditBoxes(Num As Long) As Shape
    
    Set EditBoxes = Boxes(Num)
    
End Property

'編集ボタン設定
Public Sub setEditBox(ItemNum As Long, Cell As Range)
    
    Boxes(ItemNum - 1).Visible = msoTrue
    Boxes(ItemNum - 1).Select
    
    With Selection
        
        .Top = Cell.Top
        .Left = Cell.Left
        .Width = Cell.Width
        .Height = Cell.Height
        
        '.Name = EDITBOX & Format(ItemNum, "00000")
        .ShapeRange.Fill.Transparency = 1#
        .ShapeRange.Line.Visible = msoFalse
        .OnAction = "ShowItemForm"
        
    End With
    
    'cnt = cnt + 1
    
End Sub

'EditBox数を取得
Private Function Count() As Long
    
    On Error Resume Next
    
    Count = UBound(Boxes)
    
End Function

Private Sub Class_Initialize()
    Set sht = ActiveSheet
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    cnt = Count()
    
    For i = 1 To cnt
        Set Boxes(i) = Nothing
    Next i
    
End Sub
