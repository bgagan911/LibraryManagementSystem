Attribute VB_Name = "nw_chk"

Public Sub check_roll()
On Error GoTo msg
If Form8.Text1.Text = "" Then
    MsgBox "Enter a Valid Roll No. ", vbExclamation, "Error Handler"
    Form8.Text1.SetFocus
Else
If Not IsNumeric(Form8.Text1.Text) Then
    MsgBox "Only Numeric Values allowed", vbExclamation
End If
End If
'
'===========================================
'
If Form8.Text2.Text = "" Then
    MsgBox "Enter a Valid ISBN No. ", vbExclamation, "Error Handler"
Else
If Not IsNumeric(Form8.Text2.Text) Then
    MsgBox "Only Numeric Values allowed", vbExclamation
End If
End If
msg:
If Err.Number = 0 Then
    Exit Sub
Else
    MsgBox "Error Number :>  " & Err.Number & Chr(10) & Err.Description, vbExclamation, "Error Handler"
End If
End Sub



Public Sub isdt_check(flag As Boolean)
flag = False
On Error GoTo msg
With Form2
    
    If (Not .Text5.Text = "" And .Check1.Value = 1) _
       Or (.Text5.Text = "" And .Check1.Value = 0) Then
       flag = False
    End If
    
    If (.Text5.Text = "" And .Check1.Value = 1) _
        Or (Not .Text5.Text = "" And .Check1.Value = 0) Then
            MsgBox "A book cannot be issued without a " & _
            "valid Roll No." & Chr(10) & Chr(10) & _
            "-OR-" & Chr(10) & _
            Chr(10) & _
            "Roll No. cannot be there if Book Status is " & _
            "Set to NOT-ISSUED" _
            , vbExclamation, "Issue Error"
            flag = True
     End If

    'End If
'===========================================
    If .Text1.Text = "" Or Not IsNumeric(.Text1.Text) Then
        MsgBox "Enter a Valid ISBN No. " _
        , vbExclamation, "Error Handler"
        flag = True
    End If
End With ' End of with Statement i.e. with Form2

msg:
    If Err.Number = 0 Then
         Exit Sub
    Else
     MsgBox "Error Number :>  " & Err.Number & Chr(10) & _
     Err.Description, vbExclamation, "Error Handler"
     flag = True
    End If

End Sub
