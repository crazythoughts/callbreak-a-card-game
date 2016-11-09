VERSION 5.00
Begin VB.Form frmnamecol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name"
   ClientHeight    =   2190
   ClientLeft      =   9210
   ClientTop       =   6015
   ClientWidth     =   7140
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Command1"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtname 
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblname 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmnamecol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmnamecol.BackColor = RGB(36, 84, 10)
    lblname.BackColor = RGB(36, 84, 10)
    lblname.ForeColor = vbWhite
    lblname.Font = "Comic Sans MS"
    lblname.FontSize = 20
    lblname.Alignment = 2
    lblname.Caption = "Your Name"
    txtname.Font = "Comic Sans MS"
    txtname.MaxLength = 8
    txtname.FontSize = 18
    txtname.Text = ""
    cmdok.Font = "Comic Sans MS"
    cmdok.FontSize = 15
    cmdok.FontBold = True
    cmdok.Caption = "OK"
    cmdcancel.Font = "Comic Sans MS"
    cmdcancel.FontSize = 15
    cmdcancel.FontBold = True
    cmdcancel.Caption = "CANCEL"
End Sub
Private Sub txtname_Change()
    If txtname.Text = "" Then
        cmdok.Enabled = False
    Else
        cmdok.Enabled = True
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
    Unload Server
    frmstartup.Show
End Sub
Private Sub cmdok_Click()
    player_name = txtname.Text
    Unload Me
    frmstart.Show
End Sub
