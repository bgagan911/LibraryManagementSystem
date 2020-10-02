VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   4260
   ClientLeft      =   4335
   ClientTop       =   3270
   ClientWidth     =   6330
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "in_pic.frx":0000
      ScaleHeight     =   4335
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim wd As Integer, hd As Integer
wd = Me.Width
hd = Me.Height * 1.5
Me.Move (Screen.Width - wd) / 2, _
        (Screen.Height - hd) / 2
End Sub

