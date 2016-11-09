VERSION 5.00
Begin VB.Form frmstartup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   10140
   ClientTop       =   3960
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdjoin 
      Caption         =   "Command1"
      Height          =   1092
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   4572
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Command1"
      Height          =   1092
      Left            =   600
      TabIndex        =   1
      Top             =   4080
      Width           =   4572
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "Command1"
      Height          =   1092
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   4572
   End
End
Attribute VB_Name = "frmstartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public choice As Integer
Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdjoin_Click()
    frmjoin.Show
    choice = 2
    Unload Me
End Sub

Private Sub cmdnew_Click()
    Load Server
    choice = 1
    Unload Me
    frmnamecol.Show
End Sub

Private Sub Form_Load()
    frmstartup.BackColor = RGB(36, 84, 10)
    cmdnew.Font = "Comic Sans MS"
    cmdnew.FontSize = 15
    cmdnew.Caption = "New Game"
    cmdjoin.Font = "Comic Sans MS"
    cmdjoin.FontSize = 15
    cmdjoin.Caption = "Join Game"
    cmdexit.Font = "Comic Sans MS"
    cmdexit.FontSize = 15
    cmdexit.Caption = "Exit"
End Sub
