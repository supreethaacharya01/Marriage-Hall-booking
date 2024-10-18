Imports System.Data.OleDb

PublicClassForm4

Dim cnn As OleDb.OleDbConnection

Dim cmd As OleDb.OleDbCommand

Dim str1, str2 AsString

Dim dr As OleDb.OleDbDataReader

Dim ds AsDataTable

PrivateSub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

str1 = "insert into cdetail values("& TextBox1.Text &","& TextBox2.Text & ""& TextBox3.Text &"', "& TextBox4.Text &"," "& TextBox5.Text &","& TextBox6.Text &")"

cmd = New OleDb.OleDbCommand(str1)

cmd.Connection = cnn;

cmd.ExecuteNonQuery()

MsgBox("One row inserted")

TextBox1.Clear()

TextBox2.Clear()

TextBox3.Clear()

TextBox4.Clear()

TextBox5.Clear()

TextBox6.Clear() TextBox1.Focus()

EndSub

Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

TODO: This line of code loads data into the 'MarriagedataDataSet cdetail' table. You can move, or remove it, as needed. Me.CdetailTableAdapter.Fill(Me.MarriagedataDataSet.cdetail) cnn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB 12.0, Data Source=C:\Users\LAB 2\Desktop\BSC project wrk Marriagedata.accdb")

спп.Open()

EndSub

PrivateSub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

Dim connection1 AsOleDbConnection = NewOleDbConnection() connection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12 0; Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb

connection1.Open() Connection. OleDbDataAdapter = NewOleDbDataAdapter("select*from cdetail", connection1)

Dim ds AsDataSet NewDataSet

DataGridView1.DataSource = ds.Tables(0)

adp.Fill(ds)

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

cnn.Close()

Me.Close() Form3.Show()

EndSub

EndClass

