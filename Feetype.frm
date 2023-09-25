Private Sub close_Click()
ex.Show
Unload Me
End Sub
Private Sub delete_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("fee_type") = Text1 Then
Adodc1.Recordset.delete
Adodc1.Recordset.update
MsgBox "record deleted successfully"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
End Sub
Private Sub new_Click()
Text1 = ""
End Sub
Private Sub save_Click()
Adodc1.Refresh
36
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("fee_type") = Text1.Text
Adodc1.Recordset.Fields("freq") = Text2.Text
Adodc1.Recordset.update
MsgBox "record save successfully"
End Sub
Private Sub update_Click()
Adodc1.Refresh
f = 0
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("fee_type") = Text1 Then
Adodc1.Recordset.Fields("fee_type") = Text1.Text
Adodc1.Recordset.Fields("freq") = Text2.Text
f = 1
Adodc1.Recordset.update
GoTo 50
End If
Adodc1.Recordset.MoveNext
Loop
50 If f = 1 Then
MsgBox "record modified successfully"
End If
f = 0

End Sub