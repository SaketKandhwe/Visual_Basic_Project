Private Sub close_Click()
ex.Show
Unload Me
End Sub
Private Sub Combo1_Click()
Adodc1.Refresh
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from add_class where upper(class)='" & UCase(Combo1.Text)

& "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Combo1 Then
Text1 = Adodc1.Recordset.Fields("class")
Combo3 = Adodc1.Recordset.Fields("cltype")
End If
Adodc1.Recordset.MoveNext
Loop
'MsgBox "record search successfully"
Adodc1.Refresh
End Sub
Private Sub Combo2_Click()
Adodc1.Refresh
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from add_class where upper(cltype)='" &
UCase(Combo2.Text) & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False

If Adodc1.Recordset.Fields("cltype") = Combo2 Then
Text1 = Adodc1.Recordset.Fields("class")
Combo3 = Adodc1.Recordset.Fields("cltype")
End If
Adodc1.Recordset.MoveNext
Loop
'MsgBox "record search successfully"
Adodc1.Refresh
End Sub
Private Sub delete_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Text1 Then
Adodc1.Recordset.delete
Adodc1.Recordset.update
MsgBox "record deleted successfully"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
Adodc2.Visible = False

Adodc2.Refresh
Adodc2.Recordset.MoveFirst
Do While Adodc2.Recordset.EOF = False
Combo3.AddItem (Adodc2.Recordset.Fields("cl_type"))
Adodc2.Recordset.MoveNext
Loop
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
Combo2.AddItem (Adodc1.Recordset.Fields("cltype"))
Combo1.AddItem (Adodc1.Recordset.Fields("class"))
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub new_Click()
Text1 = ""
Combo3 = ""
End Sub
Private Sub save_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("class") = Text1.Text
Adodc1.Recordset.Fields("cltype") = Combo3.Text
Adodc1.Recordset.update

MsgBox "record save successfully"
End Sub
Private Sub update_Click()
Adodc1.Refresh
f = 0
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Text1 Or Adodc1.Recordset.Fields("cltype") = Combo3
Then
Adodc1.Recordset.Fields("class") = Text1.Text
Adodc1.Recordset.Fields("cltype") = Combo3.Text
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