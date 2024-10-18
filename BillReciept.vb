Imports System.Data.OleDb PublicClassForm9

Dim cnn As OleDb.OleDbConnection

Dim cmd As OleDb.OleDbCommand

Dim dr As OleDb.OleDbDataReader DimstrAsString

PrivateSub Form9_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) HandlesMyBase.Load cnn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb") cnn.Open()

TextBox1.Focus()

EndSub

EndClass