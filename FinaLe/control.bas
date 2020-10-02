Attribute VB_Name = "control"
'Option Explicit

Public Sub install_db()
On Error GoTo dis1
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile (App.Path & "\db.mdb"), "c:\db_copy.mdb", False
MsgBox "Database Installed Successfully", vbInformation, "DataBase Installer"
Exit Sub
dis1:
response = MsgBox("Database Already Exists!.", vbExclamation, "DataBase Installer")
Exit Sub

End Sub

Public Sub un_db()
On Error GoTo dis
Set fs = CreateObject("Scripting.Filesystemobject")
fs.DeleteFile ("c:\db_copy.mdb")
MsgBox "Database Un-Instaled Successfully", vbInformation, "DataBase Installer"
Exit Sub
dis:
response = MsgBox("Error : " & Err.Description, vbCritical, "Database Installer")
End Sub

Public Sub de_bkdb()
Form2.Text1.Locked = True
Form2.Text2.Locked = True
Form2.Text3.Locked = True
Form2.text4.Enabled = False
Form2.Text5.Locked = True
Form2.Check1.Enabled = False
End Sub

Public Sub en_bkdb()
Form2.Text1.Locked = False
Form2.Text2.Locked = False
Form2.Text3.Locked = False
Form2.text4.Enabled = True
Form2.Text5.Locked = False
Form2.Check1.Enabled = True
End Sub

Public Sub save_bkdb()
Form2.cmdEdit.Enabled = True
Form2.cmdInfo.Enabled = True
Form2.cmdExit.Enabled = True
Form2.cmdFirst.Enabled = True
Form2.cmdPrev.Enabled = True
Form2.cmdNext.Enabled = True
Form2.cmdLast.Enabled = True
Form2.cmdCancel.Enabled = False
End Sub

Public Sub edit_bkdb()
Form2.cmdEdit.Enabled = False
Form2.cmdInfo.Enabled = False
Form2.cmdExit.Enabled = False
Form2.cmdFirst.Enabled = False
Form2.cmdPrev.Enabled = False
Form2.cmdNext.Enabled = False
Form2.cmdLast.Enabled = False
Form2.cmdCancel.Enabled = True
End Sub

Public Sub stuget_info()
On Error GoTo msg
    Dim sroll_no As String
    'assigning Srch value
    scroll_no = Form2.Text5.Text
    Dim Srchflag As Boolean
    Srchflag = False 'Initially Search is not found
     
        With Form2.Adodc2
        .RecordSource = "SELECT * from student where roll_no = " & scroll_no
        .Refresh
            With Form2.Adodc2.Recordset
            If .EOF <> True Then
                Call ssrch_res
                'Form7.Show
                Form2.Text7 = .Fields("stu_name") & " "
                Form2.Text8 = .Fields("stu_course") & " "
                Form2.Text9 = .Fields("stu_branch") & " "
                Form2.Text6 = .Fields("roll_no") & " "
                Srchflag = True
                Exit Sub
             End If
    
           If Srchflag = False Then 'Display msg when search not found
            MsgBox "Roll No. does not exists", vbInformation, "Search Result"
            Form2.Text5.Text = ""
            Form2.Text5.SetFocus
           End If
           End With
        End With
msg:
If Err.Number = 0 Then
    Exit Sub
Else
    MsgBox "Error Number :>  " & Err.Number & Chr(10) & Err.Description, vbExclamation, "Error Handler"
    Form2.Text1.SetFocus
End If
End Sub

Public Sub ssrch_res()
Form2.Height = 7180

Form2.Shape1.Visible = True
Form2.Label7.Visible = True
Form2.Label8.Visible = True
Form2.Label9.Visible = True
Form2.Label10.Visible = True
Form2.Label11.Visible = True

Form2.Text6.Visible = True
Form2.Text7.Visible = True
Form2.Text8.Visible = True
Form2.Text9.Visible = True

End Sub

Public Sub hide_srch()
Form2.Height = 3980

Form2.Shape1.Visible = False
Form2.Label7.Visible = False
Form2.Label8.Visible = False
Form2.Label9.Visible = False
Form2.Label10.Visible = False
Form2.Label11.Visible = False

Form2.Text6.Visible = False
Form2.Text7.Visible = False
Form2.Text8.Visible = False
Form2.Text9.Visible = False
End Sub

Public Sub hide_stusrch()
Form3.Height = 3500

Form3.Shape1.Visible = False
Form3.Label7.Visible = False
Form3.Label8.Visible = False
Form3.Label9.Visible = False
Form3.Label10.Visible = False
Form3.Label11.Visible = False

Form3.Text5.Visible = False
Form3.Text6.Visible = False
Form3.Text7.Visible = False
Form3.Text8.Visible = False
Form3.Text9.Visible = False
Form3.Check1.Visible = False
End Sub

Public Sub shw_stusrch()
Form3.Height = 6700

Form3.Shape1.Visible = True
Form3.Label7.Visible = True
Form3.Label8.Visible = True
Form3.Label9.Visible = True
Form3.Label10.Visible = True
Form3.Label11.Visible = True

Form3.Text5.Visible = True
Form3.Text6.Visible = True
Form3.Text7.Visible = True
Form3.Text8.Visible = True
Form3.Text9.Visible = True
Form3.Check1.Visible = True
End Sub

Public Sub de_studb()
Form3.Text1.Locked = True
Form3.Text2.Locked = True
Form3.Text3.Locked = True
Form3.text4.Locked = True
'
Form3.Text5.Locked = True
Form3.Text6.Locked = True
Form3.Text7.Locked = True
Form3.Text8.Locked = True
Form3.Text9.Locked = True
Form3.Check1.Enabled = False
End Sub

Public Sub en_studb()
Form3.Text1.Locked = False
Form3.Text2.Locked = False
Form3.Text3.Locked = False
Form3.text4.Locked = False
End Sub

Public Sub save_studb()
Form3.cmdEdit.Enabled = True
Form3.cmdInfo.Enabled = True
Form3.cmdExit.Enabled = True
'
Form3.cmdFirst.Enabled = True
Form3.cmdPrev.Enabled = True
Form3.cmdNext.Enabled = True
Form3.cmdLast.Enabled = True
'
Form3.cmdCancel.Enabled = False
End Sub

Public Sub edit_studb()
Form3.cmdEdit.Enabled = False
Form3.cmdInfo.Enabled = False
Form3.cmdExit.Enabled = False
'
Form3.cmdFirst.Enabled = False
Form3.cmdPrev.Enabled = False
Form3.cmdNext.Enabled = False
Form3.cmdLast.Enabled = False
'
Form3.cmdCancel.Enabled = True
End Sub

Public Sub bkget_info()
Dim sroll_no As String
    'assigning Srch value
    scroll_no = Form3.Text1.Text
    Dim Srchflag As Boolean
    Srchflag = False 'Initially Search is not found
        
        With Form3.Adodc2
        .RecordSource = "SELECT * from book where roll_no = " & scroll_no
        .Refresh
        With Form3.Adodc2.Recordset
            If .EOF <> True Then
                Call shw_stusrch
                Form3.Text5 = .Fields("isbn_no") & " "
                Form3.Text6 = .Fields("bk_name") & " "
                Form3.Text7 = .Fields("bk_author") & " "
                Form3.Text8 = .Fields("roll_no") & " "
                Form3.Text9 = .Fields("issue_date") & " "
                    If .Fields(3) = True Then
                        Form3.Check1.Value = 1
                    Else
                        Form3.Check1.Value = 0
                    End If
                Srchflag = True
                Exit Sub
             End If
    
           If Srchflag = False Then 'Display msg when search not found
            MsgBox "Roll No > " & Form3.Text1.Text & " < don't have any book", vbInformation, "Search Result"
            Form3.Text1.SetFocus
           End If
           End With
        End With
End Sub

