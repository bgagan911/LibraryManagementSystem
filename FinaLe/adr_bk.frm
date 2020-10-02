VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<:: Add Book ::>"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "adr_bk.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db_copy.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\db_copy.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "book"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      DataField       =   "bk_author"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "bk_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "isbn_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Author :"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Book Title :"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ISBN No. :"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
cmdCancel.Enabled = True
cmdAdd.Enabled = False
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Adodc1.Recordset.AddNew
Text1.SetFocus
End Sub

Private Sub cmdCancel_Click()
Adodc1.Recordset.Cancel
Adodc1.Refresh
cmdAdd.Enabled = True
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo msg:
Adodc1.Recordset.Update
Adodc1.Refresh
cmdAdd.Enabled = True
cmdCancel.Enabled = False
Adodc1.Recordset.Update
Adodc1.Refresh
msg:
If Err.Number = 0 Then
Exit Sub
Else
    MsgBox "Error No >" & Err.Number & Chr(10) & Err.Description, vbExclamation, "Save Error"
    'Debug.Print Err.Number
    'Text1.Text = ""
    Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh

Adodc1.Visible = False
cmdCancel.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
End Sub
