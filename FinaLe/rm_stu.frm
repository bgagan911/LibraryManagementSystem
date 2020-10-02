VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<:: Delete Student ::>"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "rm_stu.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Record"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "roll_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      DataField       =   "stu_name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      DataField       =   "stu_course"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      DataField       =   "stu_branch"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   2040
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "Select the Record you want to delete"
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
   Begin VB.Label Label4 
      Caption         =   "Branch :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Course :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Roll No. :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
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

Private Sub Form_Load()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
End Sub

Public Sub stu_refresh()
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub
