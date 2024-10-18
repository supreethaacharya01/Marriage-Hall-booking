Imports System.Data.OleDb

PublicClass Form6

Dim cnn As OleDb.OleDbConnection

Dim cmd As OleDb.OleDbCommand

Dim dr As OleDb.OleDbDataReader

PrivateSub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

Try

DimstrAsString

cnn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb")

cnn.Open()

str = "Delete from cdetail where Sino="& TextBox1.Text &""

cmd New OleDb.OleDbCommand(str)

cmd.Connection = cnn

cmd.ExecuteNonQuery()

MsgBox("one record deleted")

Catch ex AsException

MsgBox("no record found")

EndTry

EndSub

PrivateSub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

"TODO: This line of code loads data into the 'MarriagedataDataSet3.cdetail' table. You can move, or remove it, as needed.

Me. CdetailTableAdapter Fill (Me.MarriagedataDataSet3.cdetail)

cnn New OleDb.OleDbConnection("Provider=Microsoft ACE.OLEDB, 12.0;Data Source C:\Users\LAB 2\Desktop\BSC project wrk Marriagedata.accdb")

cnn.Open()

TextBox1.Visible = True

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

cnn.Close()

Me.Close()

Form3.Show()

EndSub

PrivateSub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click Dim connection1 AsOleDbConnection

NewOleDbConnection() connection1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accob connection1.Open()

Dim adp As OleDb.OleDbDataAdapter = New OleDbDataAdapr ("select*from cdetail", connection1)

Dim ds AsDataSet = NewDataSet

adp.Fill(ds)

DataGridView1.DataSource = ds.Tables(0)

End Sub

EndClass

