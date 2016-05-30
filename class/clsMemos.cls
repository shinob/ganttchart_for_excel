VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'追加備考用クラス
Option Explicit
Option Base 1

Private Memo(COUNT_MEMO) As String

Public Property Get Text(Num As Integer) As String
    Text = Memo(Num)
End Property

Public Property Let Text(Num As Integer, Value As String)
    Memo(Num) = Value
End Property
