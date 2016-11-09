VERSION 5.00
Begin VB.Form frmjoin 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   7635
   ClientTop       =   3690
   ClientWidth     =   7695
   FillColor       =   &H00004000&
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00404040&
      Caption         =   "Command1"
      Height          =   852
      Left            =   4200
      MaskColor       =   &H00004000&
      TabIndex        =   5
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   3012
   End
   Begin VB.CommandButton cmdjoingame 
      BackColor       =   &H000080FF&
      Caption         =   "Command1"
      Height          =   852
      Left            =   360
      MaskColor       =   &H00008000&
      TabIndex        =   4
      Top             =   3000
      Width           =   3012
   End
   Begin VB.TextBox txthostip 
      Height          =   612
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   3252
   End
   Begin VB.TextBox txtname 
      Height          =   612
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label lblhostip 
      Caption         =   "Label1"
      Height          =   972
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3972
   End
   Begin VB.Label lblname 
      Caption         =   "Label1"
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "frmjoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ipadd As String
Private Sub cmdback_Click()
    frmstartup.Show
    Unload Me
End Sub
Private Sub cmdjoingame_Click()
    player_name = txtname.Text
    ipadd = txthostip.Text
    Unload Me
    frmstart.Show
End Sub
Private Sub Form_Load()
    'form code
        frmjoin.BackColor = RGB(36, 84, 10)
    'label name code
        lblname.BackColor = RGB(36, 84, 10)
        lblname.FontSize = 17
        lblname.Font = "Comic Sans MS"
        lblname.FontItalic = True
        lblname.ForeColor = vbWhite
        lblname.AutoSize = True
        lblname.Caption = "Player Name"
    'label hostip code
        lblhostip.Font = "Comic Sans MS"
        lblhostip.BackColor = RGB(36, 84, 10)
        lblhostip.FontSize = 17
        lblhostip.ForeColor = vbWhite
        lblhostip.FontItalic = True
        lblhostip.Caption = "Enter the server IP"
        lblhostip.AutoSize = True
    'text name code
        txtname.MaxLength = 8
        txtname.Font = "Comic Sans MS"
        txtname.Font.Size = 17
        txtname.Text = ""
    'text hostip code
        txthostip.Font = "Comic Sans MS"
        txthostip.FontSize = 17
        txthostip.Text = ""
        txthostip.MaxLength = 15
        txthostip.CausesValidation = True
    'command join code
        cmdjoingame.Font = "Comic Sans MS"
        cmdjoingame.FontSize = 15
        cmdjoingame.Caption = "Join Game"
    'command back code
        cmdback.Font = "Comic Sans MS"
        cmdback.FontSize = 15
        cmdback.Caption = "Back"
End Sub
Private Sub txthostip_Change()
    If txthostip.Text = "" Or txtname.Text = "" Then
        cmdjoingame.Enabled = False
    Else
        cmdjoingame.Enabled = True
    End If
End Sub

Private Sub txthostip_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKey0 To vbKey9
      Case vbKeyBack, vbKeyClear, vbKeyDelete
      Case vbKeyLeft, vbKeyRight
      Case vbKeyDecimal
      Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub txtname_Change()
    If txtname.Text = "" Or txthostip.Text = "" Then
        cmdjoingame.Enabled = False
    Else
        cmdjoingame.Enabled = True
    End If
End Sub

