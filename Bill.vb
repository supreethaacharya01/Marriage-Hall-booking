Imports System.Data.OleDb

PublicClassForm8

Dim cnn As OleDb.OleDbConnection

Dim cmd As OleDb.OleDbCommand

Dim dr As OleDb.OleDbDataReader

Dim str1, str2 AsString

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

cnn.Close()

Me.Close()

Form3.Show()

EndSub

PrivateSub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

DimstrAsString

str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb"

con = New OleDb.OleDbConnection(str) cnn.Open()

str = "select from cdetail where Sino="& Val(TextBox4.Text) & cmd = New OleDb.OleDbCommand

cmd = NewOleDbCommand(str, cnn)

Dim OleDbReader As OleDb.OleDbDataReader = cmd.ExecuteReader()

If OleDbReader. Read = TrueThen

TextBox4.Visible = True

TextBox2.Visible = True

TextBox3.Visible = True

TextBox5.Visible = True

TextBox2.Text = OleDbReader.Item(1)

TextBox3.Text = OleDbReader.Itern(3)

Else MsgBox("no record")

EndIf

EndSub

Private Sub Form8_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

TODO: This line of code loads data into the 'MarriagedataDataSet7.abill table You

can move, or remove it, as needed. Me. AbillTableAdapter. Fill (Me.MarriagedataDataSet7.abill)

cnn = New OleDb.OleDbConnection("Provider=Microsoft ACE.OLEDB 12.0, Data

Source= |DataDirectory|\Marriagedata.accdb")

cnn.Open()

TextBox2.Visible = False

TextBox3.Visible = False

TextBox5.Visible = False

TextBox4.Focus()

DataGridView1.Visible = False

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

str2= "insert into abill values("& TextBox1.Text &","& TextBox2.Text &', & TextBox3.Text &", ""& TextBox6. Text &"", "& TextBox5.Text &")"

cmd = New OleDb.OleDbCommand(str2)

cmd.Connection = cnn

cmd.ExecuteNonQuery()

MsgBox("One row inserted")

TextBox1.Focus()

EndSub

PrivateSub Button5_Click(ByVal sender As System.Object, ByVal e As

System.EventArgs) Handles Button5.Click Dim connection1 AsOleDbConnection = New OleDbConnection()

cnn.Open() OleDb.OleDbConnection(str)

str "select from cdetail where Sino="& Val(TextBox4.Text) &""

cmd = New OleDb.OleDbCommand cmd NewOleDbCommand(str, cnn)

Dim OleDbReader As OleDb.OleDbDataReader = cmd.ExecuteReader() If

OleDbReader.Read = TrueThen

TextBox4.Visible = True

TextBox2.Visible = True TextBox3.Visible = True

TextBox5.Visible = True

TextBox2.Text = OleDbReader.Item(1) TextBox3.Text = OleDbReader.Item(3)

Else MsgBox("no record")

EndIf EndSub

PrivateSub Form8_Load(ByVal sender As System.Object, ByVal e As

System.EventArgs) Handles MyBase.Load TODO: This line of code loads data into the 'MarriagedataDataSet7.abill table. You can move, or remove it, as needed.

Me. AbillTableAdapter.Fill(Me.MarriagedataDataSet7.abill) con = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= |DataDirectory|\Marriagedata.accdb")

cnn.Open()

TextBox2.Visible = False

TextBox3.Visible = False TextBox5.Visible = False

TextBox4.Focus()

DataGridView1.Visible = False

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

str2 = "insert into abill values("& TextBox1.Text &","& TextBox2.Text &"", "& TextBox3.Text &", "& TextBox6.Text &"", "& TextBox5.Text &")"

cmd = New OleDb.OleDbCommand(str2)

cmd.Connection = cnn

cmd.ExecuteNonQuery()

MsgBox("One row inserted")

TextBox1.Focus()

EndSub

PrivateSub Button5 Click(ByVal sender As System.Object, ByVal e As

System.EventArgs) Handles Button5.Click Dim connection1 AsOleDbConnection = NewOleDbConnection()

connection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb connection1.Open()

Dim adp As OleDb.OleDbDataAdapter = NewOleDbDataAdapter("select * from abill", connection1)

Dim ds AsDataSet = NewDataSet

adp.Fill(ds)

DataGridView1.DataSource = ds.Tables(0) DataGridView1.Visible = True

EndSub

PrivateSub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

Form9.TextBox1.Text = Me.TextBox1.Text Form9.TextBox2.Text = Me.TextBox6.Text

Form9.TextBox3.Text = Me.TextBox2.Text Form9.TextBox4.Text = Me.TextBox3.Text.

Form9.TextBox5.Text = Me.TextBox5.Text

Form9.Show()

EndSub

PrivateSub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

Try

DimstrAsString

crin New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb")

cnn.Open()

str = "Delete from abill where Bilno="& TextBox1.Text &""

cmd = New OleDb.OleDbCommand(str) cmd.Connection = cnn

cmd.ExecuteNonQuery()

MsgBox("one record deleted")

Catch ex AsException

MsgBox("no record found")

EndTry

EndSub

PrivateSub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged

TextBox6.Text = Today

EndSub

EndClass