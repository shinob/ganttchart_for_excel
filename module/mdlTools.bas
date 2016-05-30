Attribute VB_Name = "mdlTools"
Option Explicit
Option Base 1

'�X�s���{�^���̒l���_�C�A���O�Őݒ�
Public Sub SetSpinButtonValue(spb As Control)
    
    On Error GoTo FIN:
    
    Dim buf As Long
    Dim val As Long
    Dim str As String
    
    buf = spb.Value
    
    str = InputBox("���l����͂��ĉ�����", "���l�ݒ�", buf)
    
    If str = "" Then Exit Sub
    
    val = CLng(str)
    
    If spb.Min <= val And val <= spb.Max Then
        spb.Value = val
    Else
        MsgBox "�͈͊O�̐��l�����͂���܂���"
    End If
    
    Exit Sub
    
FIN:
    MsgBox "�����Ȑ��l�����͂���܂���"
    
End Sub

'�v���p�e�B�Ɠ��t����X���W���擾
Public Function GetXonChart(p As clsProperty, TargetDate As Date) As Single
    
    Dim Position As Double
    Dim Column As Range
    
    Position = p.getColumnForDate(TargetDate)
    Set Column = ActiveSheet.Columns(CInt(Position))
    
    GetXonChart = Column.Left + Column.Width * (Position - CInt(Position))
    
End Function

'�I�u�W�F�N�g�Ƀ����������蓖�Ă��Ă��邩
Public Function IsSet(var As Variant) As Boolean
    
    If TypeName(var) = "Nothing" Then
        IsSet = False
    Else
        IsSet = True
    End If
    
End Function

'�F�ݒ�
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

'�󔒗�̎擾
Public Function FindBlankRow(sht As Worksheet, _
    beginRow As Long, Column As Integer) As Long
    
    Dim i As Long
    i = beginRow
    
    Do While sht.Cells(i, Column) <> ""
        i = i + 1
    Loop
    
    FindBlankRow = i
    
End Function

'�`���[�g�p�V�[�g�̊m�F
Public Function isChartSheet(sht As Worksheet) As Boolean
    
    If sht.codeName Like "shtChart*" Then
        isChartSheet = True
    Else
        isChartSheet = False
    End If
    
End Function

'�`���[�g#�擾
Public Function ChartNum(sht As Worksheet) As Integer
    
    If Not isChartSheet(sht) Then
        ChartNum = -1
        Exit Function
    End If
    
    Dim No As Integer
    No = val(Right(sht.codeName, 2)) + DEFAULT_PROPERTY
    ChartNum = No
    
End Function

'�t�H�[����̃R���g���[���Ɏg�p����t�H���g���w��
Public Sub setFontOnForm(frm As UserForm)
    
    On Error Resume Next
    
    Dim obj As Control
    Dim FontName As String
    Dim Size As Double
    
    'InputBox "Version of this Excel is", "Version", Application.Version
    'MsgBox TypeName(Application.Version)
    
    Select Case Application.Version
        
        Case "10.0", "11.0"
            FontName = "�l�r �o�S�V�b�N"
            Size = 10
        Case "10.1"
            FontName = "Osaka"
            'FontName = "Arial"
            Size = 10.5
        Case Else
            FontName = "�l�r �S�V�b�N"
            Size = 10
        
    End Select
    
    For Each obj In frm.Controls
        
        obj.Font.Name = FontName
        obj.Font.Size = Size
        
    Next obj
    
End Sub
