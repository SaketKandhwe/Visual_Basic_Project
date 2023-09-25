Private Sub Command1_Click()
CommonDialog1.ShowOpen
Text12 = CommonDialog1.FileName
Image3.Picture = LoadPicture(Text12)
End Sub
Private Sub Command2_Click()
'Unload Me
Form4.Show
End Sub
Private Sub Form_Load()

Adodc1.Visible = False
Text12.Visible = False
Adodc1.Refresh
Text13.Text = "RE00"
Text13 = Text13 & Adodc1.Recordset.RecordCount + 1
End Sub
Private Sub Timer1_Timer()
If Label27.ForeColor = vbRed Then
Label27.ForeColor = vbBlue
Else
Label27.ForeColor = vbRed
End If
End Sub