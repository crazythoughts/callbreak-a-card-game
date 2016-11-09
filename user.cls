VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private user_id As Integer
Private current_call As Integer
Private won_hands As Integer
Private points As Double
Private total_points As Double
Private name As String
Public Property Get id() As Integer
id = user_id
End Property
Public Property Let id(value As Integer)
user_id = value
End Property
Public Property Get ccall() As Integer
ccall = current_call
End Property
Public Property Let ccall(value As Integer)
current_call = value
End Property
Public Property Get win() As Integer
win = won_hands
End Property
Public Property Let win(value As Integer)
won_hands = value
End Property
Public Property Get cpoints() As Double
cpoints = points
End Property
Public Property Let cpoints(value As Double)
points = value
End Property
Public Property Get tpoints() As Double
tpoints = total_points
End Property
Public Property Let tpoints(value As Double)
total_points = value
End Property
Public Property Get player_name() As String
player_name = name
End Property
Public Property Let player_name(value As String)
name = value
End Property


