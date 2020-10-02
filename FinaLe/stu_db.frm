VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Browse >> Student's Database ::"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "stu_db.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleMode       =   0  'User
   ScaleWidth      =   2000
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1080
      Top             =   3600
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "SELECT * FROM student ORDER BY roll_no ASC"
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox Text9 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, d MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Text            =   "Text9"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Find Book"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "E&dit"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "stu_branch"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      DataField       =   "stu_course"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      DataField       =   "stu_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "roll_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Issue Status : "
      Height          =   375
      Left            =   2280
      TabIndex        =   30
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "ISBN No. :"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Title : "
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Author : "
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Date of Issue : "
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Roll No. :"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "<< Search Results >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   24
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   57.762
      X2              =   1963.899
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Caption         =   "Roll No. :"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Course :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Branch :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
cmdSave.Enabled = False
Adodc1.Recordset.Cancel
Adodc1.Refresh
Call de_studb
Call save_studb
End Sub

Private Sub cmdClose_Click()
Call hide_stusrch
End Sub

Private Sub cmdEdit_Click()
cmdSave.Enabled = True
Call edit_studb
Call en_studb
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdInfo_Click()
Call bkget_info
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
'Move to next item - if at EOF, backup one item
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "Last Record", vbInformation, "Record Browser"
Adodc1.Recordset.MovePrevious
End If
Text1.SetFocus
End Sub

Private Sub cmdPrev_Click()
'Move to Prev item - If at BOF, backup one item
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
MsgBox "First Record", vbInformation, "Record Browser"
Adodc1.Recordset.MoveNext
End If
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Call shw_stusrch
End Sub

Private Sub cmdSave_Click()
On Error GoTo msg
Adodc1.Recordset.Update
Adodc1.Refresh
Call save_studb
Call de_studb
Adodc1.Recordset.Update
Adodc1.Refresh
cmdSave.Enabled = False
msg:
If Err.Number = 0 Then
         Exit Sub
    Else
MsgBox "Error : " & Err.Number & Chr(10) & _
    Err.Description, vbExclamation, "UpDate Error"
    End If
End Sub

Private Sub Form_Load()
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh

Call hide_stusrch
  With Adodc2
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=C:\db_copy.mdb;Persist Security Info=False"
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .CommandType = adCmdText
      .RecordSource = "SELECT * from book"
      .Refresh
   End With

Adodc1.Visible = False
Adodc2.Visible = False
cmdSave.Enabled = False
cmdCancel.Enabled = False
Call de_studb
End Sub
