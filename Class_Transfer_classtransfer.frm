Private Sub close_Click()
Unload Me
End Sub
Private Sub Form_Load()
Adodc3.Refresh
Adodc3.Recordset.MoveFirst
Do While Adodc3.Recordset.EOF = False
Combo1.AddItem (Adodc3.Recordset.Fields("sess"))
Combo2.AddItem (Adodc3.Recordset.Fields("class"))
Combo3.AddItem (Adodc3.Recordset.Fields("sec"))
Adodc3.Recordset.MoveNext
Loop
End Sub
Private Sub update_Click()
Adodc3.Refresh
f = 0
Adodc3.Recordset.MoveFirst
Do While Adodc3.Recordset.EOF = False
If Adodc3.Recordset.Fields("sess") = Combo1 Or Adodc3.Recordset.Fields("class") = Combo2 Or
Adodc3.Recordset.Fields("sec") = Combo3 Then
Adodc3.Recordset.Fields("class") = Combo5.Text
Adodc3.Recordset.Fields("sec") = Combo4.Text
Adodc3.Recordset.Fields("sess") = Combo6.Text
f = 1
Adodc3.Recordset.update
GoTo 50
End If
Adodc3.Recordset.MoveNext
Loop
50 If f = 1 Then
MsgBox "class transfered successfully"
End If
f = 0
End Sub

