
Imports System.Data.OleDb

PublicClass Form5

Dim cnn As OleDb.OleDbConnection.

Dim cmd As OleDb.OleDbCommand

PrivateSub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

Try

cnn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data

DimstrAsString

Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb")

cnn.Open()

str = "select*from cdetail where Sino="& TextBox1.Text &" "

cmd = New OleDb.OleDbCommand(str)

cmd.Connection = cnn

Dim da AsNew OleDb.OleDbDataAdapter

da.SelectCommand = cmd

Dim ds AsNewDataSet

da.Fill(ds, "cdetail")

s.Tables("cdetail") DataGridView1.DataSource = ds.Tables("

Catch ex AsException

MsgBox("no data found")

EndTry

EndSub

PrivateSub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

TODO: This line of code loads data into the 'MarriagedataDataSet1.cdetail' table. You can move, or remove it, as needed. Me.CdetailTableAdapter. Fill (Me. MarriagedataDataSet1.cdetail) cnn New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb")

cnn.Open()

EndSub

Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

cnn.Close()

Me.Close()

Form3.Show()

EndSub

EndClass


