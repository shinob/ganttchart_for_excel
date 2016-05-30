VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'項目クラス
Option Explicit
Option Base 1

Private itm() As clsItem
Private Inazuma(COUNT_INAZUMA) As Date

'登録数
Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    i = UBound(itm)
    'If IsObject(itm(1)) Then i = 0
    
    Count = i
    
End Property

'取得
Public Property Get Items(Number As Long) As clsItem
    
    On Error Resume Next
    
    Set Items = itm(Number)
    
End Property

'設定
Public Property Let Items(Number As Long, newData As clsItem)
    
    On Error Resume Next
    
    Set itm(Number) = newData
    
End Property

'取得
Public Property Get getByIndex(Num As Long) As clsItem
    On Error Resume Next
    Set getByIndex = itm(Num)
End Property

'設定
Public Property Let getByIndex(Num As Long, newData As clsItem)
    On Error Resume Next
    Set itm(Num) = newData
End Property

'保存
Public Sub Save()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        itm(i).Save
        
    Next i
    
    For i = 1 To COUNT_INAZUMA
        
        shtItem.Cells(1, i + 10).Value = Inazuma(i)
        
    Next i
    
End Sub

'読込
Public Sub Load()
    
    Dim buf As clsItem
    Dim i As Long
    i = 2
    
    Do
    
        Set buf = New clsItem
        
        If buf.Load(i) Then
            
            Add
            Set itm(Count) = buf
            
        Else
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
    For i = 1 To COUNT_INAZUMA
        
        Inazuma(i) = shtItem.Cells(1, i + 10).Value
        
    Next i
    
End Sub

'分類別読込
Public Sub Load_By_Category(Num As Long)
    
    Dim buf As clsItem
    Dim i As Long
    i = 2
    
    Do
        
        Set buf = New clsItem
        If buf.Load(i) Then
            
            If buf.Category = Num Then
                
                Add
                Set itm(Count) = buf
                
            End If
            
        Else
            Exit Do
        End If
        
        i = i + 1
        
    Loop
    
    
End Sub

'追加
Public Function Add() As clsItem
    
    Dim i As Long
    i = Count() + 1
    
    ReDim Preserve itm(i)
    'Set itm(i) = New clsCategory
    
    Set Add = itm(i)
    
End Function

'並替
Public Sub Sort()
    
    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    
    Dim wkItem As clsItem
    
    cnt = Count()
    
    For i = 1 To cnt
        
        For j = 1 To cnt - i
            
            If itm(j).BeginDate <= itm(j + 1).BeginDate Then
            
            Else
                
                Set wkItem = itm(j)
                Set itm(j) = itm(j + 1)
                Set itm(j + 1) = wkItem
                
            End If
            
        Next j
        
    Next i
    
    For i = 1 To cnt
        itm(i).SubItems.Sort
    Next i
    
End Sub

'イナズマ線
Public Property Get InazumaDate(Num As Integer) As Date
    
    InazumaDate = Inazuma(Num)
    
End Property

'
Public Property Let InazumaDate(Num As Integer, newDate As Date)
    
    Inazuma(Num) = newDate
    
End Property

'イナズマ線設定
Public Sub setInazuma(Num As Integer, newDate As Date)
    
    Dim i As Long
    
    Inazuma(Num) = newDate
    
    For i = 1 To Count()
        
        Call itm(i).setInazuma(Num, newDate)
        
    Next i
    
End Sub

'イナズマ線保存
Public Sub SaveInazuma()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        Call itm(i).SaveInazuma
        
    Next i
    
    For i = 1 To COUNT_INAZUMA
        
        shtItem.Cells(1, i + 10).Value = Inazuma(i)
        
    Next i
    
End Sub

'分類データのみ保存
Public Sub SaveCateogory()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        Call itm(i).SaveCategory
        
    Next i
    
End Sub

Private Sub Class_Initialize()
    'Call Load
End Sub

Private Sub Class_Terminate()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        Set itm(i) = Nothing
        
    Next i
    
End Sub
