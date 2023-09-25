Private Sub Command3_Click()
Command4.Visible = True
Command3.Visible = False
Text2.PasswordChar = "*"
End Sub
Private Sub Command4_Click()
Command3.Visible = True
Command4.Visible = False
Text2.PasswordChar = ""
End Sub
Private Sub login_Click()
Adodc1.Refresh
If Text1.Text = Adodc1.Recordset.Fields("use") And Text2.Text =
Adodc1.Recordset.Fields("pass") Then
MsgBox ("congratulation!!!!...you are logged in")
Unload Me
MDIForm1.Show
Else
MsgBox ("Sorry!!!!...you user id /password is wrong plzz enter correct user id or password")
End If
End Sub

Private Sub cancel_Click()
End
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
End Sub