Private Sub Label22_Click()
End Sub
Private Sub Command1_Click()
CommonDialog1.ShowOpen
Text15 = CommonDialog1.FileName
Image2.Picture = LoadPicture(Text15)
End Sub
Private Sub Command2_Click()
'Unload Me
Form5.Show
End Sub
Private Sub Command3_Click()
'Unload Me
Form2.Show
End Sub
Private Sub Form_Load()
Adodc1.Visible = False

Adodc2.Visible = False
Text15.Visible = False
End Sub