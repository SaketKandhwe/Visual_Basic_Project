Dim str As String
Private Sub Check1_Click()
If Check1.Value = Checked Then
submit.Visible = True
Else
submit.Visible = False
End If
End Sub
Private Sub Command1_Click()
CommonDialog1.ShowOpen
Text1 = CommonDialog1.FileName
Image2.Picture = LoadPicture(Text1)
End Sub
Private Sub Command2_Click()
Unload Form2
Unload Form4
Unload Form5
End Sub

Private Sub Command3_Click()
'Unload Me
Form4.Show
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
Text1.Visible = False
End Sub
Private Sub submit_Click()
Form2.Adodc1.Refresh
With Form2.Adodc1.Recordset
.AddNew
.Fields("studna") = Form2.Text1.Text + " " + Form2.Text2.Text + " " + Form2.Text3.Text
.Fields("regisno") = Form2.Text13.Text
If Form2.Option1.Value = True Then
.Fields("gen") = "Male"
End If
If Form2.Option2.Value = True Then
.Fields("gen") = "Female"
End If
.Fields("dob") = Form2.DTPicker1.Value
.Fields("sess") = Form2.Text14.Text
.Fields("category") = Form2.Combo1
.Fields("stutype") = Form2.Text6.Text

.Fields("mtongue") = Form2.Text5.Text
.Fields("religion") = Form2.Combo2
.Fields("class") = Form2.Text4.Text
.Fields("nationality") = Form2.Text11.Text
.Fields("email") = Form2.Text7.Text
.Fields("mno") = Val(Form2.Text8)
.Fields("tno") = Val(Form2.Text9)
.Fields("preschool") = Form2.Text10.Text
.Fields("stuphoto") = Form2.Text12.Text
.update
End With
Form4.Adodc1.Refresh
With Form4.Adodc1.Recordset
.AddNew
.Fields("regisno") = Form2.Text13.Text
.Fields("pAdd") = Form4.Text9.Text
.Fields("pcity") = Form4.Text1.Text
.Fields("pstate") = Form4.Text2.Text
.Fields("pcode") = Val(Form4.Text3)
.Fields("comadd") = Form4.Text4.Text
.Fields("comcountry") = Form4.Text5.Text
.Fields("comcity") = Form4.Text6.Text
.Fields("comstate") = Form4.Text7.Text
.Fields("comcode") = Val(Form4.Text8)

.update
End With
Form4.Adodc2.Refresh
With Form4.Adodc2.Recordset
.AddNew
.Fields("regisno") = Form2.Text13.Text
.Fields("fna") = Form4.Text10.Text
.Fields("fmno") = Val(Form4.Text11)
.Fields("fqual") = Form4.Combo3
.Fields("foffadd") = Form4.Text12.Text
.Fields("femail") = Form4.Text13.Text
.Fields("ftno") = Val(Form4.Text14)
.Fields("foccu") = Form4.Combo2
.Fields("fmod") = Form4.Combo1
.Fields("fphoto") = Form4.Text15.Text
.update
End With
Adodc1.Refresh
With Adodc1.Recordset
.AddNew
.Fields("regisno") = Form2.Text13.Text
.Fields("mna") = Text4.Text
.Fields("mmno") = Val(Text3)
.Fields("mqual") = Combo3

.Fields("moffadd") = Text2.Text
.Fields("memail") = Text10.Text
.Fields("mtno") = Val(Text5)
.Fields("moccu") = Combo2
.Fields("mmod") = Combo1
.Fields("mphoto") = Text1.Text
.update
End With
Adodc2.Refresh
With Adodc2.Recordset
.AddNew
.Fields("regisno") = Form2.Text13.Text
.Fields("presc") = Text6.Text
.Fields("remarks") = Text7.Text
.Fields("lastexam") = Text8.Text
.Fields("year") = Text9.Text
.Fields("status") = Combo4.Text
.Fields("marks") = Val(Text5)
.Fields("board") = Text12.Text
.Fields("blood") = Text13.Text
.Fields("height") = Text14.Text
'.Fields("weigth") = Text15.Text
If Check2.Value = Checked Then
str = str + "TC"

.Fields("docu") = str
End If
If Check3.Value = Checked Then
str = str + "CC"
.Fields("docu") = str
End If
If Check4.Value = Checked Then
str = str + "Report C"
.Fields("docu") = str
End If
If Check5.Value = Checked Then
str = str + "DOB certificate"
.Fields("docu") = str
End If
.Fields("addhar") = Text16.Text
.update
End With
MsgBox "thanks for completing the form"
End Sub