VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************
' データリスト
'****************************************
Option Explicit
Option Base 1

Private Type DataListBin
    ValueType As Integer
    No As Long
    Row As Long
    Deps As Integer
End Type

Private Const DT_Category = 1
Private Const DT_Item = 2

Private Bins() As DataListBin
Private BinsPos As Long

Private Categorys As clsCategorys
Private Items As clsItems
Private Schedules As clsSchedules

Private Sub setList()
    
    Set Categorys = Nothing
    Set Items = Nothing
    Set Schedules = Nothing
    
    'Call Categorys.Load
    Set Categorys = New clsCategorys
    
    Set Items = New clsItems
    Call Items.Load
    
    Set Schedules = New clsSchedules
    Call Schedules.Load(-1)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    For i = 1 To Categorys.Count()
        Categorys.getByIndex(i).ClearItems
        'With Categorys.getByIndex(i).Items
        '    .Clear
        '    MsgBox .Count()
        'End With
    Next i
    
    For i = 1 To Items.Count()
        
        j = Items.getByIndex(i).LinkItem
        If j < 2 Then
        
        Else
            With Items.getByIndex(j - 1).SubItems
                .Add
                .getByIndex(.Count()) = Items.getByIndex(i)
            End With
        End If
        
        j = Items.getByIndex(i).Category
        If j < 2 Then
        
        Else
            With Categorys.getByIndex(j - 1).Items
                .Add
                .getByIndex(.Count()) = Items.getByIndex(i)
            End With
        End If
        
    Next i
    
    For i = 1 To Schedules.Count()
        
        j = Schedules.getByIndex(i).Item
        If j < 2 Then
        Else
            With Items.getByIndex(j - 1).Schedules
                .Add
                .getByIndex(.Count()) = Schedules.getByIndex(i)
            End With
        End If
        
    Next i
    
End Sub

Public Sub makeListBox(ListBox As Control)
    
    'MsgBox TypeName(ListBox)
    If TypeName(ListBox) <> "ListBox" Then Exit Sub
    
    ListBox.Clear
    setList
    
    Dim i As Long
    
    ReDim Bins(Categorys.Count() + Items.Count())
    BinsPos = 0
    
    For i = 1 To Categorys.Count()
        With Categorys.getByIndex(i)
            ListBox.AddItem .Name
            addBins DT_Category, .No, ListBox.ListCount(), 0
            'ListBox.AddItem i
            addItemsOnListBox ListBox, .Items, 1
        End With
    Next i
    
End Sub

Private Sub addItemsOnListBox(ListBox As Control, SubItems As clsItems, Deps As Long)
    
    'MsgBox "addItemsOnListBox" & SubItems.Count()
    
    Dim i As Long
    Dim indent As String
    
    For i = 0 To Deps * 4
        indent = indent & " "
    Next i
    
    For i = 1 To SubItems.Count()
        With SubItems.getByIndex(i)
            ListBox.AddItem indent & .Name
            addBins DT_Item, .No, ListBox.ListCount(), Deps
            addItemsOnListBox ListBox, .SubItems, Deps + 1
        End With
    Next i
    
End Sub

Private Function Count() As Long
    On Error Resume Next
    'Count = UBound(Bins)
    Count = BinsPos
End Function

Private Sub addBins(ValueType As Integer, No As Long, Row As Long, Deps As Long)
    
    Dim i As Long
    
    BinsPos = BinsPos + 1
    i = Count()
    
    With Bins(i)
        .ValueType = ValueType
        .No = No
        .Row = Row
        .Deps = Deps
    End With
    
End Sub

Public Function MoveToUp(Num As Long) As Boolean
    
    'Categorys.getByIndex(1).Items.Clear
    'Exit Function
    
    MoveToUp = False
    If Num < 1 Then Exit Function
    
    Dim i As Integer
    
    Dim d1 As Long
    Dim d2 As Long
    'Dim wk
    Dim flg As Boolean
    
    d1 = Num + 1
    flg = False
    
    For i = 1 To Num
        
        d2 = d1 - i
        'Debug.Print d2
        
        If Bins(d1).Deps > Bins(d2).Deps Then
            Exit Function
        ElseIf Bins(d1).Deps = Bins(d2).Deps Then
            flg = True
            Exit For
        End If
        
    Next i
    
    If Not flg Then Exit Function
    
    Select Case Bins(d1).ValueType
        
        Case DT_Category
            MoveToUp = ExchangeCategory(d1, d2)
        Case DT_Item
            MoveToUp = ExchangeItem(d1, d2)
    End Select
    
End Function

Public Function MoveToDown(Num As Long) As Boolean
    
    'Categorys.getByIndex(1).Items.Clear
    'Exit Function
    
    MoveToDown = False
    If Num < 1 Then Exit Function
    
    Dim i As Integer
    
    Dim d1 As Long
    Dim d2 As Long
    'Dim wk
    Dim flg As Boolean
    
    d1 = Num + 1
    flg = False
    
    For i = d1 + 1 To Count()
        
        d2 = i
        'Debug.Print d2
        
        If Bins(d1).Deps > Bins(d2).Deps Then
            Exit Function
        ElseIf Bins(d1).Deps = Bins(d2).Deps Then
            flg = True
            Exit For
        End If
        
    Next i
    
    If Not flg Then Exit Function
    'MsgBox "d1 = " & d1 & Chr(13) & "d2 = " & d2
    'If d2 < 1 Then Exit Function
    'If d2 > Count() Then Exit Function
    'MsgBox "d2 = " & d2 & Chr(13) & "Count() = " & Count()
    'Exit Function
    
    Select Case Bins(d1).ValueType
        
        Case DT_Category
            MoveToDown = ExchangeCategory(d1, d2)
        Case DT_Item
            MoveToDown = ExchangeItem(d1, d2)
    End Select
    
End Function

Private Function ExchangeCategory(d1 As Long, d2 As Long) As Boolean
    
    MsgBox "分類を選択されています。"
    ExchangeCategory = False
    
End Function

Private Function ExchangeItem(d1 As Long, d2 As Long) As Boolean
    
    Dim wk
    
    'MsgBox Bins(d1).ValueType
    'Exit Function
    
    'MsgBox "D1 = " & d1 & Chr(13) & "D2 = " & d2
    d1 = Bins(d1).No - 1
    d2 = Bins(d2).No - 1
    
    wk = Items.getByIndex(d1).No
    Items.getByIndex(d1).No = Items.getByIndex(d2).No
    Items.getByIndex(d2).No = wk
    
    'MsgBox "d1 = " & Items.getByIndex(d1).Name & Chr(13) _
    '    & "d2 = " & Items.getByIndex(d2).Name
    With Items.getByIndex(d1)
        .Save
        .Schedules.UpdateItemAll .No
        .UpdateLinkItemOfSubItems
    End With
    
    With Items.getByIndex(d2)
        .Save
        .Schedules.UpdateItemAll .No
        .UpdateLinkItemOfSubItems
    End With
    
    ExchangeItem = True
    'wk = Items.getByIndex(d1)
    'Items.getByIndex(d1) = Items.getByIndex(d2)
    'Items.getByIndex(d1) = wk
    
End Function

Private Sub Class_Initialize()
    'Call setList
End Sub
