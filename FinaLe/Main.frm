VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "<:: Library Management System ::>"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8670
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4935
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/27/2005"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:06 PM"
         EndProperty
      EndProperty
      Enabled         =   0   'False
      MouseIcon       =   "Main.frx":11C2
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Index           =   0
      NegotiatePosition=   1  'Left
      WindowList      =   -1  'True
      Begin VB.Menu iu_db 
         Caption         =   "Install/Uninstall DataBase"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu bk_db 
         Caption         =   "Book's Database"
      End
      Begin VB.Menu s_db 
         Caption         =   "Student's Database"
      End
   End
   Begin VB.Menu a_dd 
      Caption         =   "&Add/Remove"
      Begin VB.Menu nw_bk 
         Caption         =   "Add Book"
      End
      Begin VB.Menu nw_stu 
         Caption         =   "Add Student"
      End
      Begin VB.Menu rm_bk 
         Caption         =   "Remove Book"
      End
      Begin VB.Menu rm_stu 
         Caption         =   "Rmove Student"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub bk_db_Click()
Form2.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub iu_db_Click()
Form1.Show
End Sub


Private Sub MDIForm_Load()
Me.Caption = App.Title
Form9.Show
End Sub

Private Sub nw_bk_Click()
Form4.Show
End Sub

Private Sub nw_stu_Click()
Form5.Show
End Sub

Private Sub rm_bk_Click()
Form6.Show
End Sub

Private Sub rm_stu_Click()
Form7.Show
End Sub

Private Sub s_db_Click()
Form3.Show
End Sub

