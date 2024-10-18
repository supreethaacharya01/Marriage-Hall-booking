PublicClassForm2

PrivateSub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

If f(TextBox1.Text = "Adrnin"And TextBox2.Text = "1234") Then MsgBox("Login Succesfull")

Form3.Show()

Elself TextBox1.Text = ""Then

MsgBox("Enter your username ")

Elself TextBox2.Text = ""Then

MsgBox("Enter your Password")

Else

MsgBox("Incorrect username and password")

TextBox1.Clear()

TextBox2.Clear()

TextBox1.Focus()

EndIf

EndSub

PrivateSub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

MsgBox("Logged Out Successfully")

Me.Close()

Form1.Show()

EndSub

PrivateSub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

Me.Hide()

Form1.Show()

EndSub