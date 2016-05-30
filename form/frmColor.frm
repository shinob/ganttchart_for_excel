VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColor 
   Caption         =   "UserForm1"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   OleObjectBlob   =   "frmColor.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'�F�I��p�t�H�[��
Option Explicit
Option Base 1

Public SelectedColor As Long    '�I��F
Private SetColorObject As Control

'�C���^�[�t�F�[�X�ݒ�
Private Sub InitInterFace()
    
    btnOK.Caption = "�ݒ�"
    btnCancel.Caption = "���"
    btnNoColor.Caption = "���ݒ�"
    
    frmPalette.Caption = "�p���b�g"
    frmBlue.Caption = "��"
    
    Me.Caption = "�F�ݒ�"
    
End Sub

'������
Public Sub Initialize(obj As Control)
    
    Set SetColorObject = obj
    
    With imgColor
        
        .BackStyle = obj.BackStyle
        .BackColor = obj.BackColor
        
    End With
    
End Sub

'�I��F�����{�ɐݒ�
Public Sub setColor(color As Long)
    
    With imgColor
        
        .BackStyle = fmBackStyleOpaque
        .BackColor = color
        
    End With
    
End Sub

'�I���{�^���F�ύX
Private Sub SetImgBgColor(img As Image, color As Long)
    
    img.BackColor = color
    img.BackStyle = fmBackStyleOpaque
    
End Sub

'�I���{�^���F�ݒ�
Private Sub SetPaletteColor(blue As Integer)
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 4
        
        For j = 0 To 4
            
            Controls("imgPalette" & (i * 5 + j + 1)).BackColor = _
                RGB(64 * i, 64 * j, blue)
            
        Next j
        
    Next i
    
End Sub

'���x���ݒ�
Private Sub SetBlueColor()
    
    Dim i As Integer
    
    For i = 0 To 4
        
        Controls("imgBlue" & (i + 1)).BackColor = _
            RGB(0, 0, 64 * i)
        
    Next i
    
End Sub

'����I��
Private Sub btnCancel_Click()

    SelectedColor = -2
    Me.Hide
    
End Sub

'���F�I��
Private Sub btnNoColor_Click()
    
    SetColorObject.BackStyle = fmBackStyleTransparent
    SelectedColor = -1
    Me.Hide
    
End Sub

'�I��
Private Sub btnOK_Click()
    
    If imgColor.BackStyle = fmBackStyleTransparent Then
        SetColorObject.BackStyle = fmBackStyleTransparent
        SelectedColor = -1
    Else
        With SetColorObject
            .BackStyle = fmBackStyleOpaque
            .BackColor = imgColor.BackColor
        End With
        SelectedColor = imgColor.BackColor
    End If
    
    Me.Hide
    
End Sub

'�F���ύX���ꂽ���ۂ�
Public Function isColorChanged() As Boolean
    
    If SelectedColor < -1 Then
        isColorChanged = False
    Else
        isColorChanged = True
    End If
    
End Function

'���x���ύX
Private Sub imgBlue1_Click(): SetPaletteColor 0: End Sub
Private Sub imgBlue2_Click(): SetPaletteColor 256 / 4: End Sub
Private Sub imgBlue3_Click(): SetPaletteColor 256 / 2: End Sub
Private Sub imgBlue4_Click(): SetPaletteColor 256 / 4 * 3: End Sub
Private Sub imgBlue5_Click(): SetPaletteColor 256: End Sub

'�F�I��
Private Sub imgPalette1_Click(): setColor imgPalette1.BackColor: End Sub
Private Sub imgPalette2_Click(): setColor imgPalette2.BackColor: End Sub
Private Sub imgPalette3_Click(): setColor imgPalette3.BackColor: End Sub
Private Sub imgPalette4_Click(): setColor imgPalette4.BackColor: End Sub
Private Sub imgPalette5_Click(): setColor imgPalette5.BackColor: End Sub
Private Sub imgPalette6_Click(): setColor imgPalette6.BackColor: End Sub
Private Sub imgPalette7_Click(): setColor imgPalette7.BackColor: End Sub
Private Sub imgPalette8_Click(): setColor imgPalette8.BackColor: End Sub
Private Sub imgPalette9_Click(): setColor imgPalette9.BackColor: End Sub
Private Sub imgPalette10_Click(): setColor imgPalette10.BackColor: End Sub
Private Sub imgPalette11_Click(): setColor imgPalette11.BackColor: End Sub
Private Sub imgPalette12_Click(): setColor imgPalette12.BackColor: End Sub
Private Sub imgPalette13_Click(): setColor imgPalette13.BackColor: End Sub
Private Sub imgPalette14_Click(): setColor imgPalette14.BackColor: End Sub
Private Sub imgPalette15_Click(): setColor imgPalette15.BackColor: End Sub
Private Sub imgPalette16_Click(): setColor imgPalette16.BackColor: End Sub
Private Sub imgPalette17_Click(): setColor imgPalette17.BackColor: End Sub
Private Sub imgPalette18_Click(): setColor imgPalette18.BackColor: End Sub
Private Sub imgPalette19_Click(): setColor imgPalette19.BackColor: End Sub
Private Sub imgPalette20_Click(): setColor imgPalette20.BackColor: End Sub
Private Sub imgPalette21_Click(): setColor imgPalette21.BackColor: End Sub
Private Sub imgPalette22_Click(): setColor imgPalette22.BackColor: End Sub
Private Sub imgPalette23_Click(): setColor imgPalette23.BackColor: End Sub
Private Sub imgPalette24_Click(): setColor imgPalette24.BackColor: End Sub
Private Sub imgPalette25_Click(): setColor imgPalette25.BackColor: End Sub

Private Sub UserForm_Initialize()
    
    Call InitInterFace
    Call SetBlueColor
    Call SetPaletteColor(0)
    
    Call mdlTools.setFontOnForm(Me)
    
    SelectedColor = -2
    
End Sub
