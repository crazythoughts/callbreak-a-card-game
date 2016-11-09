VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private card_name As String
Private card_value As Integer
Private card_img As StdPicture
Private card_id As Integer
Public Property Get id() As Integer
id = card_id
End Property
Public Property Let id(value As Integer)
card_id = value
End Property
Public Property Get name() As String
name = card_name
End Property
Public Property Let name(value As String)
card_name = value
End Property
Public Property Get cvalue() As Integer
cvalue = card_value
End Property
Public Property Let cvalue(value As Integer)
card_value = value
End Property
Public Property Get img() As StdPicture
Set img = card_img
End Property
Public Property Let img(value As StdPicture)
Set card_img = value
End Property

