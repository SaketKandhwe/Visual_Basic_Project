Private Sub Image1_Click()
Timer1.Enabled = True
Timer1_Timer
End Sub
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
If ProgressBar1.Value = 0 Then
Label1.Caption = "Loading........"
ElseIf ProgressBar1.Value = 10 Then
Label1.Caption = "Loading coffee shop management system........"
ElseIf ProgressBar1.Value = 20 Then
Label1.Caption = "please wait........"
ElseIf ProgressBar1.Value = 30 Then
Label1.Caption = "Loading coffee shop management system........"
ElseIf ProgressBar1.Value = 40 Then
Label1.Caption = "please wait........"
ElseIf ProgressBar1.Value = 50 Then
Label1.Caption = "Loading........"
ElseIf ProgressBar1.Value = 60 Then

Label1.Caption = "Loading coffee shop management system........"
ElseIf ProgressBar1.Value = 70 Then
Label1.Caption = "please wait........"
ElseIf ProgressBar1.Value = 80 Then
Label1.Caption = "Loading coffee shop management system........"
ElseIf ProgressBar1.Value = 90 Then
Label1.Caption = "Almost done!!!!!........"
ElseIf ProgressBar1.Value = 95 Then
Label1.Caption = "Done process completed!!!!!........"
'ProgressBar1.Value = ProgressBar1.Value + 5
'If ProgressBar1.Value = 80 Then
'ProgressBar1.Value = ProgressBar1 + 20
'Label1.Caption = "Loading complete!!!!!!!!......"
ElseIf ProgressBar1.Value >= ProgressBar1.Max Then
Timer1.Enabled = False
'End If
'MsgBox "Loading complete!"
Unload Me
Student_login.Show
End If
End Sub