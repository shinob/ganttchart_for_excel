VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'分類クラス
Option Explicit
Option Base 1

Public No As Long           '行番号
Public Name As String       '名称
Public Visible As Boolean   '表示・非表示

Public Memo As New clsMemos

Private sht As Worksheet    'データ用ワークシート

Public Delete As Boolean

Public Items As clsItems

'保存
Public Function Save() As Boolean
    
    If No < 2 Then
        No = FindBlankRow(sht, 2, 2)
    End If
    
    With sht
        
        Dim i As Integer
            
        If Delete Then
            
            For i = 1 To 3 + COUNT_MEMO
                
                .Cells(No, i).Value = ""
                
            Next i
        
        Else
        
            .Cells(No, 1) = No
            .Cells(No, 2) = Name
            .Cells(No, 3) = Visible
            For i = 1 To COUNT_MEMO
                .Cells(No, 3 + i).Value = Memo.Text(i)
            Next i
            
        End If
        
    End With
    
    Save = True
    
End Function

'読込
Public Function Load(Row As Long) As Boolean
    
    If sht.Cells(Row, 1) = "" Then
        Load = False
        Exit Function
    End If
    
    No = sht.Cells(Row, 1)
    Name = sht.Cells(Row, 2)
    Visible = sht.Cells(Row, 3)
    
    Dim i As Integer
    
    For i = 1 To COUNT_MEMO
        
        Memo.Text(i) = sht.Cells(Row, 3 + i).Value
        
    Next i
    
    Load = True
    
End Function

'項目の分類番号更新
Public Sub UpdateNo()
    
    Dim i As Long
    
    For i = 1 To Items.Count()
        
        Items.Items(i).Category = No
        
    Next i
    
End Sub

Public Sub ClearItems()

    Set Items = Nothing
    Set Items = New clsItems

End Sub

Private Sub Class_Initialize()
    Set sht = shtCategory
    Visible = True
    Set Items = New clsItems
End Sub

Private Sub Class_Terminate()
    
    Set Items = Nothing
    
End Sub
