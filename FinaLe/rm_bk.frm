VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<:: Delete Book ::>"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "rm_bk.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      DataField       =   "issue_status"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Record"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
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
   Begin VB.TextBox Text4 
      DataField       =   "issue_date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, dd MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2040
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "    Select the Record you want to delete"
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
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Title : "
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Author : "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Issue : "
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Roll No. :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Issue Status : "
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
'***************************************
'Display Confirmation Message
'If answer is NO then Exit subroutine
'Deletes the Current Record,
'Then Move to the Next Record
'If it is the End of File then Refresh
'But if it is also the Begining of File
'Then Display Message <Data Base Emtpy>
'and call to < cmdAdd > is made
'***************************************
Dim Rvalue
Rvalue = MsgBox("Are you sure you want to delete this Record?", vbQuestion + vbYesNo, "Delete Item")
If Rvalue = vbNo Then
   Exit Sub
End If
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
    Adodc1.Refresh
  If Adodc1.Recordset.BOF = True Then
    MsgBox "You must add a record.", vbOKOnly + vbInformation, "Empty file"
  Else
    Adodc1.Recordset.MoveFirst
  End If
End If
Call stu_refresh
Text1.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Public Sub stu_refresh()
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Check1.Enabled = False
End Sub
