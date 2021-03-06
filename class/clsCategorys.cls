VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCategorys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'分類用クラス
Option Explicit
Option Base 1

Private itm() As clsCategory

'登録済み分類数
Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    i = UBound(itm)
    
    Count = i
    
End Property

'削除以外の分類数
Public Property Get CountNotDeleted() As Long
    
    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    
    cnt = Count
    j = 0
    
    For i = 1 To cnt
        If Not itm(i).Delete Then
            j = j + 1
        End If
    Next i
    
    CountNotDeleted = j
    
End Property

'分類取得
Public Property Get Items(Number As Long) As clsCategory
    
    On Error Resume Next
    
    Set Items = itm(Number)
    
End Property

'分類設定
Public Property Let Items(Number As Long, newData As clsCategory)
    
    On Error Resume Next
    
    Set itm(Number) = newData
    
End Property

Public Property Get getByIndex(Num As Long) As clsCategory
    On Error Resume Next
    Set getByIndex = itm(Num)
End Property

Public Property Let getByIndex(Num As Long, newData As clsCategory)
    On Error Resume Next
    Set itm(Num) = newData
End Property

'保存
Public Sub Save()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        itm(i).Save
        
    Next i
    
End Sub

'読込
Public Sub Load()
    
    Dim buf As clsCategory
    Dim i As Long
    i = 2
    
    Do
    
        Set buf = New clsCategory
        
        If buf.Load(i) Then
            
            Add
            Set itm(Count) = buf
            
        Else
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
End Sub

'追加
Public Function Add() As clsCategory
    
    Dim i As Long
    i = Count() + 1
    
    ReDim Preserve itm(i)
    'Set itm(i) = New clsCategory
    
    Set Add = itm(i)
    
End Function

'入換
Public Sub Exchange(src As Long, dst As Long)
    
    Dim wk As clsCategory
    
    Set wk = itm(src)
    Set itm(src) = itm(dst)
    Set itm(dst) = wk
    
    Dim i As Long
    
    i = itm(src).No
    itm(src).No = itm(dst).No
    itm(dst).No = i
    
    itm(src).UpdateNo
    itm(dst).UpdateNo
    
End Sub

'項目読込
Public Sub LoadItems()
    
    Dim i As Long
    Dim wkItems As New clsItems
    
    wkItems.Load
    
    For i = 1 To wkItems.Count
        
        If wkItems.Items(i).Category > 1 Then
            
            itm(i).Items.Add
            Set itm(i).Items.Items(itm(i).Items.Count) = wkItems.Items(i)
            
        End If
        
    Next i
    
End Sub

Private Sub Class_Initialize()
    Call Load
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    
    For i = 1 To Me.Count()
    
        Set itm(i) = Nothing
        
    Next i
    
End Sub
