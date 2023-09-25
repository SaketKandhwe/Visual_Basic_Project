Private Sub Combo4_Click()
Adodc1.Refresh
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from fee_entry where upper(class)='" & UCase(Combo4.Text)
& "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Combo4 Then
Combo1 = Adodc1.Recordset.Fields("class")
Combo2 = Adodc1.Recordset.Fields("fee_type")
Combo3 = Adodc1.Recordset.Fields("sem")
Text5 = Adodc1.Recordset.Fields("fee")
End If
Adodc1.Recordset.MoveNext
Loop
'MsgBox "record search successfully"
Adodc1.Refresh
End Sub

Private Sub Combo5_Click()
Adodc1.Refresh
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from fee_entry where upper(fee_type)='" &
UCase(Combo5.Text) & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("fee_type") = Combo5 Then
Combo1 = Adodc1.Recordset.Fields("class")
Combo2 = Adodc1.Recordset.Fields("fee_type")
Combo3 = Adodc1.Recordset.Fields("sem")
Text5 = Adodc1.Recordset.Fields("fee")
End If
Adodc1.Recordset.MoveNext
Loop
'MsgBox "record search successfully"
Adodc1.Refresh
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
Adodc2.Visible = False
Adodc3.Visible = False
Adodc2.Refresh
Adodc2.Recordset.MoveFirst
Do While Adodc2.Recordset.EOF = False
Combo1.AddItem (Adodc2.Recordset.Fields("class"))
Adodc2.Recordset.MoveNext
Loop
Adodc3.Refresh
Adodc3.Recordset.MoveFirst
Do While Adodc3.Recordset.EOF = False
Combo2.AddItem (Adodc3.Recordset.Fields("fee_type"))
Adodc3.Recordset.MoveNext
Loop
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
Combo4.AddItem (Adodc1.Recordset.Fields("class"))
'Combo5.AddItem (Adodc1.Recordset.Fields("fee_type"))
Adodc1.Recordset.MoveNext
Loop
'coding for adding name
Dim f_fee_type As Boolean
Adodc1.Refresh
Do While (Adodc1.Recordset.EOF = False)
f_fee_type = False

For i = 0 To Combo5.ListCount - 1
If UCase(Combo5.List(i)) = UCase(Adodc1.Recordset.Fields("fee_type")) Then
f_fee_type = True
Exit For
End If
Next
If f_fee_type = False Then
Combo5.AddItem Adodc1.Recordset.Fields("fee_type")
End If
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub new_Click()
Text5 = ""
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo5 = ""
Combo5.Clear
'coding for adding name
Dim f_fee_type As Boolean
Adodc1.Refresh
Do While (Adodc1.Recordset.EOF = False)

f_fee_type = False
For i = 0 To Combo5.ListCount - 1
If UCase(Combo5.List(i)) = UCase(Adodc1.Recordset.Fields("fee_type")) Then
f_fee_type = True
Exit For
End If
Next
If f_fee_type = False Then
Combo5.AddItem Adodc1.Recordset.Fields("fee_type")
End If
Adodc1.Recordset.MoveNext
Loop
End Sub
Private Sub save_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("class") = Combo1.Text
Adodc1.Recordset.Fields("fee_type") = Combo2.Text
Adodc1.Recordset.Fields("sem") = Combo3.Text
Adodc1.Recordset.Fields("fee") = Text5.Text
Adodc1.Recordset.update
MsgBox "record save successfully"
End Sub
Private Sub update_Click()

Adodc1.Refresh
f = 0
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Combo1 Then
Adodc1.Recordset.Fields("class") = Combo1.Text
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
Private Sub close_Click()
ex.Show
Unload Me
End Sub
Private Sub delete_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst

Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("class") = Combo1 Then
Adodc1.Recordset.delete
Adodc1.Recordset.update
MsgBox "record deleted successfully"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub