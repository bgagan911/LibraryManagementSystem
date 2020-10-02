VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install/Uninstall Database"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "iu_db.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUn 
      Caption         =   "Un-Install Database"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "Install Database"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdIn_Click()
Call install_db
End Sub

Private Sub cmdUn_Click()
Call un_db
End Sub

Private Sub Form_Load()

End Sub
