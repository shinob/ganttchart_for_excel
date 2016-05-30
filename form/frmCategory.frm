VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCategory 
   Caption         =   "Category"
   ClientHeight    =   4060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   OleObjectBlob   =   "frmCategory.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'分類編集用フォーム
Option Explicit
Option Base 1

Private ctg As New clsCategorys '分類データ
Private flgEdit As Boolean

Private Memo As clsMemos

Private Const NEW_ITEM = "新規登録"
Private Const MSG_NONAME = "名称が設定されていません。"
Private Const MSG_ISSETCHANGE = "変更内容設定しますか?"

'インターフェース設定
Private Sub InitInterFace()
    
    Label1.Caption = "名称"
    
    btnEditMemo.Caption = "備考編集"
    
    btnSet.Caption = "設定"
    btnSave.Caption = "保存"
    btnCancel.Caption = "取消"
    chkVisible.Caption = "表示"
    
    frmDetail.Caption = "詳細"
    
    Me.Caption = "分類"
    
End Sub

'詳細転写
Private Sub SetDetailData(Category As clsCategory)
    
    With Category
        
        txtNo.Text = .No
        txtName.Text = .Name
        chkVisible.Value = .Visible
        
    End With
    
End Sub

'データ生成
Private Function CREATE() As clsCategory
    
    Dim tmp As New clsCategory
    
    With tmp
        
        .No = val(txtNo.Text)
        .Name = txtName.Text
        .Visible = chkVisible.Value
        
    End With
    
    Set CREATE = tmp
    
End Function

'一覧更新
Private Function UpdateList()
    
    With lstCategory
        
        .Clear
        
        .AddItem NEW_ITEM
        
        Dim i As Long
        
        For i = 1 To ctg.Count()
            
            .AddItem ctg.Items(i).Name
            
        Next i
        
    End With
    
    flgEdit = False
    
End Function

Private Sub btnEditMemo_Click()
    
    Dim frm As New frmMemos
    
    Call frm.Load(Memo)
    
    frm.Show
    
    Set frm = Nothing
    
End Sub

'設定
Private Sub btnSet_Click()
    
    Dim buf As clsCategory
    
    If txtName.Text = "" Then
        MsgBox MSG_NONAME
        Exit Sub
    End If
    
    Set buf = CREATE()
    
    Set buf.Memo = Memo
    
    If buf.No < 2 Then
    
        buf.No = ctg.Count() + 2
        txtNo.Text = buf.No
        
        ctg.Add
        ctg.Items(ctg.Count()) = buf
        
    Else
        
        ctg.Items(buf.No - 1) = Nothing
        ctg.Items(buf.No - 1) = buf
        
    End If
    
    UpdateList
    
End Sub

'保存
Private Sub btnSave_Click()
    
    If flgEdit Then
    
        If MsgBox(MSG_ISSETCHANGE, vbYesNo) = vbYes Then
            Call btnSet_Click
        End If
        
    End If
    
    ctg.Save
    Me.Hide
    
End Sub

'取消
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub lstCategory_Click()
    
    Dim i As Long
    i = lstCategory.ListIndex
    
    If flgEdit Then
        
        If MsgBox(MSG_ISSETCHANGE, vbYesNo) = vbYes Then
            Call btnSet_Click
            lstCategory.ListIndex = i
        End If
        
    End If
    
    If i < 1 Then
        
        Dim buf As New clsCategory
        SetDetailData buf
        Set Memo = New clsMemos
        
    Else
        
        SetDetailData ctg.Items(i)
        Set Memo = ctg.Items(i).Memo
        
    End If
    
    flgEdit = False
    
End Sub

'変更フラグ
Private Sub txtName_Change(): flgEdit = True: End Sub
Private Sub chkVisible_Click(): flgEdit = True: End Sub

Private Sub UserForm_Initialize()

    Call InitInterFace
    Call UpdateList
    
    Set Memo = New clsMemos
    
    Call mdlTools.setFontOnForm(Me)
    
End Sub
