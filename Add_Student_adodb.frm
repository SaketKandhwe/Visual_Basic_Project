Private Sub Form_Load()

'code for adding class
Dim f_cl As Boolean
Adodc2.Refresh
Do While (Adodc2.Recordset.EOF = False)
f_cl = False
For i = 0 To Combo1.ListCount - 1
If UCase(Combo1.List(i)) = UCase(Adodc2.Recordset.Fields("cl")) Then
f_cl = True
Exit For
End If
Next
If f_cl = False Then
Combo1.AddItem Adodc2.Recordset.Fields("cl")
End If
Adodc2.Recordset.MoveNext
Loop
'code for adding sec
Dim f_sec As Boolean
Adodc2.Refresh
Do While (Adodc2.Recordset.EOF = False)
f_sec = False
For i = 0 To Combo2.ListCount - 1
If UCase(Combo2.List(i)) = UCase(Adodc2.Recordset.Fields("sec")) Then
f_sec = True

Exit For
End If
Next
If f_sec = False Then
Combo2.AddItem Adodc2.Recordset.Fields("sec")
End If
Adodc2.Recordset.MoveNext
Loop
End Sub
Private Sub new_Click()
Text1 = ""
Combo1 = ""
Combo2 = ""
End Sub
Private Sub save_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("name") = Text1.Text
Adodc1.Recordset.Fields("class") = Combo1.Text
Adodc1.Recordset.Fields("sec") = Combo2.Text
Adodc1.Recordset.Fields("sess") = Text2.Text
Adodc1.Recordset.update
MsgBox "record save successfully"
End Sub

Private Sub update_Click()
Adodc1.Refresh
f = 0
Adodc1.Recordset.MoveFirst
Do While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("name") = Text1 Then
Adodc1.Recordset.Fields("name") = Text1.Text
Adodc1.Recordset.Fields("class") = Combo1.Text
Adodc1.Recordset.Fields("sec") = Combo2.Text
Adodc1.Recordset.Fields("sess") = Text2.Text
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
If Adodc1.Recordset.Fields("name") = Text1 Then
Adodc1.Recordset.delete
Adodc1.Recordset.update
MsgBox "record deleted successfully"
End If
Adodc1.Recordset.MoveNext
Loop
End Sub