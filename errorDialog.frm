VERSION 5.00
Begin VB.Form errorDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   1470
   ClientLeft      =   10845
   ClientTop       =   8205
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5895
   Begin VB.CommandButton btngoback 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnretry 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblerrormsg 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "errorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btngoback_Click()
    frmjoin.Show
    Unload frmstart
    Unload Me
End Sub
Private Sub btnretry_Click()
    Unload frmstart
    Unload Me
    Load frmstart
    frmstart.Show
End Sub
Private Sub Form_Load()
    'label error code
        lblerrormsg.FontSize = 18
        lblerrormsg.ForeColor = RGB(50, 50, 50)
        lblerrormsg.Caption = "Cannot connect to the given server"
    'command button code
        btnretry.Caption = "Retry"
        btngoback.Caption = "Go Back"
End Sub
