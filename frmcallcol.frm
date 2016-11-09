VERSION 5.00
Begin VB.Form frmcallcol 
   Caption         =   "Calls"
   ClientHeight    =   1485
   ClientLeft      =   9555
   ClientTop       =   6690
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   5910
   Begin VB.CommandButton cmdcall 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtcall 
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblcall 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmcallcol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public current_calls As Integer
Public pressed As Integer
Private Sub cmdcall_Click()
    current_calls = Val(txtcall.Text)
    pressed = 1
    frmstart.tcpclient.SendData "call|" & current_calls
    frmstart.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    pressed = 0
    frmcallcol.BackColor = RGB(36, 84, 10)
    lblcall.BackColor = RGB(36, 84, 10)
    lblcall.ForeColor = vbWhite
    lblcall.Alignment = 2
    lblcall.FontSize = 15
    lblcall.Font = "Comic Sans MS"
    lblcall.Caption = "Enter Your Call"
    txtcall.Font = "Comic Sans MS"
    txtcall.FontSize = 15
    txtcall.Text = ""
    cmdcall.Font = "Comic Sans MS"
    cmdcall.FontSize = 15
    cmdcall.FontBold = True
    cmdcall.Caption = "OK"
    frmstart.Enabled = False
    txtcall.MaxLength = 1
End Sub

Private Sub txtcall_Change()
    If txtcall.Text = "" Then
        cmdcall.Enabled = False
    Else
        cmdcall.Enabled = True
    End If
End Sub

Private Sub txtcall_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey1 To vbKey8
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

