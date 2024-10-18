
Imports System.Data.OleDb

PublicClassForm7

Dim cnn As OleDb.OleDbConnection

Dim cmd As OleDb.OleDbCommand

DimstrAsString

Private Sub Button1Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

DimstrAsString

str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\LAB

2\Desktop\BSC project wrk Marriagedata.accdb"

cnn = New OleDb.OleDbConnection(str)

cnn.Open()

= "select from cdetail where Sino="& Val(TextBox1.Text) &""

str cmd = New OleDb.OleDbCommand

Dim OleDbReader As OleDb.OleDbDataReader = cmd.ExecuteReader()

cmd = NewOleDbCommand(str, cnn)

If OleDbReader. Read = True Then

TextBox1.Visible = True

TextBox2.Visible = True TextBox3.Visible = True

TextBox2.Text = OleDbReader.Itern(2)

Else

TextBox3.Text = OleDbReader.Item(5)

MsgBox("No Record Found")

EndIf

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

Dim connectionstring AsString

DimstrAsString

TextBox2.Visible = True

TextBox3.Visible = True

connectionstring = "Provider=Microsoft.ACE.OL FDB.12.0; Data

Source=C:\Users\LAB 2\Desktop\BSC project wrk Marriagedata.accdb"

cnn New OleDb.OleDbConnection(connectionstring)

str = "update cdetail set Guest="& TextBox3.Text &" where Sino="&

cnn.Open()

TextBox1.Text &" "

cmd = New OleDb.OleDbCommand(str)

cmd.Connection = cnn

cmd.ExecuteNonQuery()

MsgBox("Record Updated")

cmd.Dispose()

cnn.Close()

EndSub

PrivateSub Button2 Click(ByVal sender As System.Object, ByVal System.EventArgs) Handles Button2.clis e As

cnn.Close() Me.Close()

Form 3.Show()

EndSub

PrivateSub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) HandlesMyBase.Load

TODO: This line of code loads data into the 'MarriagedataDataSet4 cdetail' table You can move, or remove it, as needed. Me CdetailTableAdapter.Fill(Me.MarriagedataDataSet4.cdetail)

cnn = New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12 0; Data Source C:\Users\LAB 2\Desktop\BSC project wrk\Marriagedata.accdb")

cnn.Open()

TextBox2.Visible = False

TextBox3.Visible = False

TextBox1.Focus()

EndSub

EndClass



