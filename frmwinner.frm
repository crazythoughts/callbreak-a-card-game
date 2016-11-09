VERSION 5.00
Begin VB.Form frmwinner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WINNER"
   ClientHeight    =   12450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   23235
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmwinner.frx":0000
   ScaleHeight     =   12450
   ScaleWidth      =   23235
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblwinner 
      Caption         =   "Label1"
      Height          =   2295
      Left            =   1680
      TabIndex        =   0
      Top             =   4560
      Width           =   20535
   End
End
Attribute VB_Name = "frmwinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmwinner.BackColor = RGB(36, 84, 10)
lblwinner.Font = "Brush Script MT"
lblwinner.BackColor = RGB(36, 84, 10)
lblwinner.ForeColor = vbWhite
lblwinner.Alignment = 2
lblwinner.FontSize = 90
lblwinner.Caption = "The winner is " & winner_name
Unload Server
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmstart
frmstartup.Show
End Sub
