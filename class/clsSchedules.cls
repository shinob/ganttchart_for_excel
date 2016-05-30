VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'工程用クラス
Option Explicit
Option Base 1

Private itm() As clsSchedule

Public Property Get ActLatest()
    
    Dim i As Long
    Dim wkDate As Date
    
    For i = 1 To Count()
        
        If Items(i).ActEnd < Items(i).ActBegin Then
            wkDate = Items(i).ActBegin + 1
        Else
            wkDate = Items(i).ActEnd
        End If
        
        If ActLatest < wkDate Then ActLatest = wkDate
        
    Next i
    
    If ActLatest < FIRSTDATE Then ActLatest = BeginDate
    
End Property

Public Property Get BeginDate() As Date

    Dim i As Long
    
    BeginDate = #1/1/2100#
    
    For i = 1 To Count()
        
        If Items(i).BeginDate < BeginDate Then
            BeginDate = Items(i).BeginDate
        Else
            
        End If
        
    Next i
    
End Property

Public Property Get EndDate() As Date

    Dim i As Long
    
    For i = 1 To Count()
        
        If EndDate < Items(i).EndDate Then
            EndDate = Items(i).EndDate
        Else
        
        End If
        
    Next i
    
End Property

'登録済み工程数
Public Property Get Count() As Long
    
    On Error Resume Next
    
    Dim i As Long
    i = UBound(itm)
    
    Count = i
    
End Property

'取得
Public Property Get Items(Number As Long) As clsSchedule
    
    On Error Resume Next
    
    Set Items = itm(Number)
    
End Property

'設定
Public Property Let Items(Number As Long, newData As clsSchedule)
    
    On Error Resume Next
    
    Set itm(Number) = newData
    
End Property

Public Property Get getByIndex(Num As Long) As clsSchedule
    On Error Resume Next
    Set getByIndex = itm(Num)
End Property

Public Property Let getByIndex(Num As Long, newData As clsSchedule)
    On Error Resume Next
    Set itm(Num) = newData
End Property

'工程を更新
Public Sub UpdateItemAll(No)
    
    Dim i As Long
    
    For i = 1 To Count()
        With itm(i)
            .Item = No
            .SaveItem
        End With
    Next i
    
End Sub

'保存
Public Sub Save()
    
    Dim i As Long
    
    For i = 1 To Count()
        
        itm(i).Save
        
    Next i
    
End Sub

'読込
Public Sub Load(ItemNum As Long)
'   -1 : 全て
    
    Dim buf As clsSchedule
    Dim i As Long
    i = 2
    
    Do
    
        Set buf = New clsSchedule
        
        If buf.Load(i) Then
            
            If ItemNum < 0 Or ItemNum = buf.Item Then
                
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
Public Function Add() As clsSchedule
    
    Dim i As Long
    i = Count() + 1
    
    ReDim Preserve itm(i)
    'Set itm(i) = New clsSchedule
    
    Set Add = itm(i)
    
End Function

'並替
Public Sub Sort()
    
    Dim i As Long
    Dim j As Long
    Dim cnt As Long
    
    Dim No As Long
    
    Dim wkSchedule As clsSchedule
    
    cnt = Count()
    
    For i = 1 To cnt
        
        For j = 1 To cnt - i
            
            If itm(j).BeginDate <= itm(j + 1).BeginDate Then
                
            Else
                
                'Call wkSchedule.Copy(itm(j))
                'Call itm(j).Copy(itm(j + 1))
                'Call itm(j + 1).Copy(wkSchedule)
                Set wkSchedule = itm(j)
                Set itm(j) = itm(j + 1)
                Set itm(j + 1) = wkSchedule
                
                No = itm(j).No
                itm(j).No = itm(j + 1).No
                itm(j + 1).No = No
                
            End If
            
        Next j
        
    Next i
    
End Sub

Public Function getTotalWorkDays()
    
    Dim i As Long
    Dim j As Long
    
    Dim wkDays As Double
    
    i = Count()
    For j = 1 To i
        
        wkDays = wkDays + itm(j).getWorkDays
        
    Next j
    
    getTotalWorkDays = wkDays
    
End Function


'進捗状況から実績データを生成
Public Sub SetActionDateByStatus(mStatus As Integer)
    
    Dim i As Long
    Dim j As Long
    
    Dim wkTotalDays As Double
    Dim wkWorkDays As Double
    
    wkTotalDays = getTotalWorkDays * (mStatus / 100)
    
    i = Count()
    
    For j = 1 To i
        
        With itm(j)
            
            wkWorkDays = .getWorkDays
            
            If wkTotalDays <= 0 Then
                
                .ActBegin = 0
                .ActEnd = 0
                
            ElseIf wkWorkDays < wkTotalDays Then
                
                .ActBegin = .PlanBegin
                .ActEnd = .PlanEnd
                
            ElseIf wkWorkDays > wkTotalDays Then
                
                .ActBegin = .PlanBegin
                .ActEnd = .PlanBegin + wkTotalDays
                
            End If
            
            wkTotalDays = wkTotalDays - wkWorkDays
            
        End With
        
    Next j
    
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
