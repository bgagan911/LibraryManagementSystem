VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Browse >> Book Database ::"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "bk_db.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6837.215
   ScaleMode       =   0  'User
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker text4 
      Bindings        =   "bk_db.frx":11C2
      DataField       =   "issue_date"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dddd, dd MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dddd, dd MMMM yyyy"
      Format          =   150470659
      CurrentDate     =   38474
   End
   Begin VB.TextBox Text1 
      DataField       =   "isbn_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "bk_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      DataField       =   "bk_author"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      DataField       =   "roll_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Text            =   "Text9"
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Find Student"
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "E&dit"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      DataField       =   "issue_status"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3360
      Top             =   1680
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   2880
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "SELECT * FROM book ORDER BY isbn_no ASC"
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
   Begin VB.Label Label1 
      Caption         =   "ISBN No. :"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Title : "
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Author : "
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Issue : "
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Roll No. :"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Issue Status : "
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Left            =   1200
      TabIndex        =   19
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "Branch :"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Course :"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Roll No. :"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   3630.379
      Y2              =   3630.379
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.Cancel
Adodc1.Refresh
Call de_bkdb
Call save_bkdb
End Sub

Private Sub cmdClose_Click()
Call hide_srch
End Sub

Private Sub cmdEdit_Click()
cmdSave.Enabled = True
Call hide_srch
Call edit_bkdb
Call en_bkdb
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdInfo_Click()
Call stuget_info
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
End Sub

Private Sub cmdPrev_Click()
'Move to Prev item - If at BOF, backup one item
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
    MsgBox "First Record", vbInformation, "Record Browser"
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub cmdSave_Click()
Dim flag As Boolean
Call isdt_check(flag)
If flag = True Then
    Exit Sub
End If

With Adodc1.Recordset
    .Fields("roll_no") = Me.Text5.Text & ""
    .Fields("issue_date") = Me.Check1.Value
End With
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh
Call hide_srch
Call save_bkdb
Call de_bkdb
End Sub

Private Sub Form_Load()
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh

Call hide_srch
   With Adodc2
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=C:\db_copy.mdb;Persist Security Info=False"
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .CommandType = adCmdText
      .RecordSource = "SELECT * from student"
      .Refresh
   End With

cmdSave.Enabled = False
Adodc1.Visible = False
Adodc2.Visible = False
cmdCancel.Enabled = False
Call de_bkdb
End Sub
