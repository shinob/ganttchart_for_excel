VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'�J�����_�[�t�H�[��
Option Explicit
Option Base 1

Private hd As New clsHoliday    '�x���f�[�^
Public mode As String           '����ݒ� SELECT : ���t�I��
Public SelectedDate As Date     '�I���
Private ActiveDate As Date      '�I��O��
Public flgSelect As Boolean     '�I���t���O

Private Const DATEFORMAT = "yyyy�Nmm��"

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()

    lblMonth.Caption = ""
    btnExit.Caption = "�I��"
    
End Sub

'�{�^������
Private Sub btnDate1_Click(): btnClick btnDate1.Caption: End Sub
Private Sub btnDate2_Click(): btnClick btnDate2.Caption: End Sub
Private Sub btnDate3_Click(): btnClick btnDate3.Caption: End Sub
Private Sub btnDate4_Click(): btnClick btnDate4.Caption: End Sub
Private Sub btnDate5_Click(): btnClick btnDate5.Caption: End Sub
Private Sub btnDate6_Click(): btnClick btnDate6.Caption: End Sub
Private Sub btnDate7_Click(): btnClick btnDate7.Caption: End Sub
Private Sub btnDate8_Click(): btnClick btnDate8.Caption: End Sub
Private Sub btnDate9_Click(): btnClick btnDate9.Caption: End Sub
Private Sub btnDate10_Click(): btnClick btnDate10.Caption: End Sub
Private Sub btnDate11_Click(): btnClick btnDate11.Caption: End Sub
Private Sub btnDate12_Click(): btnClick btnDate12.Caption: End Sub
Private Sub btnDate13_Click(): btnClick btnDate13.Caption: End Sub
Private Sub btnDate14_Click(): btnClick btnDate14.Caption: End Sub
Private Sub btnDate15_Click(): btnClick btnDate15.Caption: End Sub
Private Sub btnDate16_Click(): btnClick btnDate16.Caption: End Sub
Private Sub btnDate17_Click(): btnClick btnDate17.Caption: End Sub
Private Sub btnDate18_Click(): btnClick btnDate18.Caption: End Sub
Private Sub btnDate19_Click(): btnClick btnDate19.Caption: End Sub
Private Sub btnDate20_Click(): btnClick btnDate20.Caption: End Sub
Private Sub btnDate21_Click(): btnClick btnDate21.Caption: End Sub
Private Sub btnDate22_Click(): btnClick btnDate22.Caption: End Sub
Private Sub btnDate23_Click(): btnClick btnDate23.Caption: End Sub
Private Sub btnDate24_Click(): btnClick btnDate24.Caption: End Sub
Private Sub btnDate25_Click(): btnClick btnDate25.Caption: End Sub
Private Sub btnDate26_Click(): btnClick btnDate26.Caption: End Sub
Private Sub btnDate27_Click(): btnClick btnDate27.Caption: End Sub
Private Sub btnDate28_Click(): btnClick btnDate28.Caption: End Sub
Private Sub btnDate29_Click(): btnClick btnDate29.Caption: End Sub
Private Sub btnDate30_Click(): btnClick btnDate30.Caption: End Sub
Private Sub btnDate31_Click(): btnClick btnDate31.Caption: End Sub
Private Sub btnDate32_Click(): btnClick btnDate32.Caption: End Sub
Private Sub btnDate33_Click(): btnClick btnDate33.Caption: End Sub
Private Sub btnDate34_Click(): btnClick btnDate34.Caption: End Sub
Private Sub btnDate35_Click(): btnClick btnDate35.Caption: End Sub
Private Sub btnDate36_Click(): btnClick btnDate36.Caption: End Sub
Private Sub btnDate37_Click(): btnClick btnDate37.Caption: End Sub

'�{�^������
Private Sub btnClick(Value As String)
    
    If Value = "" Then Exit Sub
    
    Select Case (mode)
        
        '���t�I��
        Case "SELECT"
        
            SelectedDate = GetSelectedDate(val(Value))
            'MsgBox flgSelect
            Me.Hide
            
        '�x���ݒ�
        Case Else
        
            Call ChangeHoliday(Value)
        
    End Select
    
    flgSelect = True
    
End Sub

'�x���ݒ�
Private Sub ChangeHoliday(Value As String)
    
    '���t�ȊO�Ȃ�I��
    If Value = "" Then Exit Sub
    
    'MsgBox value
    'lblMonth.Caption = value
        
    Dim Today As Date
    
    'Today = ActiveMonth() - 1 + Val(value)
    Today = GetSelectedDate(val(Value))
    
    If hd.isHoliday(Today) Then
        hd.setHoliday(Today) = False
    Else
        hd.setHoliday(Today) = True
    End If
    
    Call UpdateColors
    
End Sub

'�I����擾
Private Function GetSelectedDate(Value As Integer) As Date
    
    GetSelectedDate = ActiveMonth - 1 + Value
    
End Function

'�ݒ茎�l�擾
Private Function getMonthValue(Value As Date) As Long

    getMonthValue = (Year(Value) - Year(FIRSTDATE)) * 12 + _
        Month(Value) - 1
    
End Function

'���t�ݒ�
Public Sub SetDate(ByVal Value As Date)
    
    'MsgBox "frmCalendar.SetDate = " & Value
    If Value < FIRSTDATE Then Value = Date
    ActiveDate = Value
    
    Dim i As Long
    i = getMonthValue(ActiveDate)
    
    If spbMonth.Value = i Then
        Call UpdateCalendar
    Else
        spbMonth.Value = i
    End If
    
End Sub

'�I��
Private Sub btnExit_Click()
    'MsgBox flgSelect
    Me.Hide
End Sub

'�N���̕ύX
Private Sub spbMonth_Change()
    
    Dim s As String
    
    s = Format(ActiveMonth(), DATEFORMAT)
    
    lblMonth.Caption = s
    UpdateCalendar
    
End Sub

'�I������1�����擾
Private Function ActiveMonth() As Date
    
    Dim y As Integer
    Dim m As Integer
    
    y = Int(spbMonth.Value / 12) + Year(hd.getFirstDate)
    m = spbMonth.Value Mod 12 + 1
    
    ActiveMonth = DateSerial(y, m, 1)
    
End Function

'�J�����_�[�\���̍X�V
Private Sub UpdateCalendar()
    
    Dim TargetDate As Date
    Dim w As Integer
    
    'MsgBox "frmCalendar.UpdateCalendar " & ActiveDate
    
    TargetDate = ActiveMonth()
    w = WeekDay(TargetDate)
    
    Dim i As Integer
    Dim D As Integer
    Dim obj As Control
    
    For i = 1 To 37
        
        D = Day(TargetDate)
        Set obj = Controls("btnDate" & i)
        
        If i < w Or (D + w) < i Then
            
            'Controls("btnDate" & i).Caption = ""
            obj.Caption = ""
            obj.Visible = False
            
        Else
            'Controls("btnDate" & i).Caption = d
            obj.Caption = D
            obj.Visible = True
            TargetDate = TargetDate + 1
        End If
        
    Next i
    
    UpdateColors
    'lblMonth = w
    
End Sub

'���x���\��
Private Sub UpdateColors()
    
    Dim i As Integer
    Dim Today As Date
    
    Dim msg As String
    
    Today = ActiveMonth()
    'MsgBox Today & " = " & ActiveDate
    'MsgBox "frmCalendar.UpdateColors " & ActiveDate
    msg = "frmCalendar.UpdateColors()" & vbCr
    
    For i = 1 To 37
        
        With Controls("btnDate" & i)
        
            If .Caption <> "" Then
                
                msg = msg & Today & " = " & ActiveDate & vbCr
                
                If Today = Int(ActiveDate) Then
                    .ForeColor = RGB(0, 0, 255)
                    'MsgBox "Today = ActiveDate"
                ElseIf hd.isHoliday(Today) Then
                    .ForeColor = RGB(255, 0, 0)
                Else
                    .ForeColor = RGB(0, 0, 0)
                End If
                
                Today = Today + 1
                
            End If
            
        End With
        
    Next i
    
    'MsgBox msg
    
End Sub

'������
Private Sub UserForm_Initialize()
    
    Dim Today As Integer
    
    Call InitInterFace
    
    'Today = (Year(Date) - Year(hd.getFirstDate)) * 12 + Month(Date) - 1
    'spbMonth.Value = Today
    spbMonth.Value = getMonthValue(Date)
    
    Call mdlTools.setFontOnForm(Me)

End Sub
