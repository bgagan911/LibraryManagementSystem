VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Splash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LMS - Library Management System : Installation"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdC 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "frmStart.frx":11C2
      ScaleHeight     =   855
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   4080
      Top             =   1680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strURL As String
Private intProgressBarMax As Integer

Private Sub cmdC_Click()
frmLogin.Show
Unload Me
End Sub

Private Sub Form_Load()
'This is a built-in Property to avoid Multiple Instance
If App.PrevInstance Then
Unload Me
Exit Sub
End If
'Continue with load code here

cmdC.Enabled = False
cmdC.Caption = "Loading"
Label1.Caption = Chr(10) & "Please Wait..."
intProgressBarMax = 100
ConfigureBar ProgressBar1
With Timer1
   .Enabled = False
   .Interval = 500
   End With
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static counter As Integer
   ' Test the variable named counter. If it's less
   ' then the module variable intProgressBarMax then
   ' show the ProgressBar control.
   If counter = 25 Then
    Label1.Caption = Chr(10) & "Initializing" & _
                " Variables..."
    Else
    If counter = 50 Then
       Label1.Caption = Chr(10) & "Initializing" & _
                        " Forms..."
    Else
    If counter = 75 Then
    Label1.Caption = Chr(10) & "Setting up" & _
                " Environment..."
    End If
    End If
   End If
   '===============================================
   If counter = intProgressBarMax Then
      Timer1.Enabled = False
      counter = 0
    Label1.Caption = Chr(10) & "LMS Successfully" & _
                " Loaded"
      cmdC.Caption = "&Continue"
      cmdC.Enabled = True
      'MsgBox "PAth : " & App.Path
   Else
      counter = counter + 5
      ProgressBar1.Value = counter
   End If
End Sub
Private Sub ConfigureBar(prgBar As ProgressBar)
   With ProgressBar1
      .Max = intProgressBarMax
      .Visible = False
   End With
End Sub


