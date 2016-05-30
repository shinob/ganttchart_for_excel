Attribute VB_Name = "mdlTools"
Option Explicit
Option Base 1

'スピンボタンの値をダイアログで設定
Public Sub SetSpinButtonValue(spb As Control)
    
    On Error GoTo FIN:
    
    Dim buf As Long
    Dim val As Long
    Dim str As String
    
    buf = spb.Value
    
    str = InputBox("数値を入力して下さい", "数値設定", buf)
    
    If str = "" Then Exit Sub
    
    val = CLng(str)
    
    If spb.Min <= val And val <= spb.Max Then
        spb.Value = val
    Else
        MsgBox "範囲外の数値が入力されました"
    End If
    
    Exit Sub
    
FIN:
    MsgBox "無効な数値が入力されました"
    
End Sub

'プロパティと日付からX座標を取得
Public Function GetXonChart(p As clsProperty, TargetDate As Date) As Single
    
    Dim Position As Double
    Dim Column As Range
    
    Position = p.getColumnForDate(TargetDate)
    Set Column = ActiveSheet.Columns(CInt(Position))
    
    GetXonChart = Column.Left + Column.Width * (Position - CInt(Position))
    
End Function

'オブジェクトにメモリが割り当てられているか
Public Function IsSet(var As Variant) As Boolean
    
    If TypeName(var) = "Nothing" Then
        IsSet = False
    Else
        IsSet = True
    End If
    
End Function

'色設定
Public Function EditColor(obj As Control) As Boolean
    
    On Error Resume Next
    
    Dim frm As New frmColor
    
    With frm
        
        .Initialize obj
        .Show
        
    End With
    
    If IsSet(frm) Then
    
        EditColor = frm.isColorChanged()
        
    End If
    
End Function

'空白列の取得
Public Function FindBlankRow(sht As Worksheet, _
    beginRow As Long, Column As Integer) As Long
    
    Dim i As Long
    i = beginRow
    
    Do While sht.Cells(i, Column) <> ""
        i = i + 1
    Loop
    
    FindBlankRow = i
    
End Function

'チャート用シートの確認
Public Function isChartSheet(sht As Worksheet) As Boolean
    
    If sht.codeName Like "shtChart*" Then
        isChartSheet = True
    Else
        isChartSheet = False
    End If
    
End Function

'チャート#取得
Public Function ChartNum(sht As Worksheet) As Integer
    
    If Not isChartSheet(sht) Then
        ChartNum = -1
        Exit Function
    End If
    
    Dim No As Integer
    No = val(Right(sht.codeName, 2)) + DEFAULT_PROPERTY
    ChartNum = No
    
End Function

'フォーム上のコントロールに使用するフォントを指定
Public Sub setFontOnForm(frm As UserForm)
    
    On Error Resume Next
    
    Dim obj As Control
    Dim FontName As String
    Dim Size As Double
    
    'InputBox "Version of this Excel is", "Version", Application.Version
    'MsgBox TypeName(Application.Version)
    
    Select Case Application.Version
        
        Case "10.0", "11.0"
            FontName = "ＭＳ Ｐゴシック"
            Size = 10
        Case "10.1"
            FontName = "Osaka"
            'FontName = "Arial"
            Size = 10.5
        Case Else
            FontName = "ＭＳ ゴシック"
            Size = 10
        
    End Select
    
    For Each obj In frm.Controls
        
        obj.Font.Name = FontName
        obj.Font.Size = Size
        
    Next obj
    
End Sub
