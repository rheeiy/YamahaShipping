Imports System.Security.Cryptography
Imports System.Text
Public Class frm_Login
    Public login_npk As String
    Public login_name As String
    Public login_admin As String
    Public login_pwd As String

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If myValidasi() Then
            prosesLogin(txtUsername.Text, txtPassword.Text)
        End If
    End Sub

    Shared Function GetMd5Hash(ByVal md5Hash As MD5, ByVal input As String) As String

        ' Convert the input string to a byte array and compute the hash. 
        Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input))

        ' Create a new Stringbuilder to collect the bytes 
        ' and create a string. 
        Dim sBuilder As New StringBuilder()

        ' Loop through each byte of the hashed data  
        ' and format each one as a hexadecimal string. 
        Dim i As Integer
        For i = 0 To data.Length - 1
            sBuilder.Append(data(i).ToString("x2"))
        Next i

        ' Return the hexadecimal string. 
        Return sBuilder.ToString()

    End Function 'GetMd5Hash

    Private Sub prosesLogin(ByVal username As String, ByVal password As String)
        Dim myTemp As New Class_Login
        Dim myDT As New DataTable

        myDT = myTemp.getLogin(username, password)

        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            login_npk = myDT.Rows(0).Item("login_npk").ToString
            login_name = myDT.Rows(0).Item("login_nama").ToString
            login_admin = myDT.Rows(0).Item("login_admin").ToString
            login_pwd = myDT.Rows(0).Item("login_password").ToString
            Dim pwd As String = txtPassword.Text
            Using md5Hash As MD5 = MD5.Create()
                Dim hash As String = GetMd5Hash(md5Hash, pwd)
                If hash = login_pwd Then
                    mycancel()
                    Me.Hide()
                    frm_MAIN.Show()
                Else
                    MessageBox.Show("Incorrect password !", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    txtPassword.Text = ""
                    txtPassword.Focus()
                End If
            End Using
        Else
            MessageBox.Show("Incorrect username", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            txtPassword.Text = ""
            txtUsername.Focus()
        End If
    End Sub

    Private Sub mycancel()
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.Focus()
    End Sub

    Private Function myValidasi() As Boolean
        If (txtUsername.Text = "") Then
            MessageBox.Show("Please enter your username", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtUsername.Focus()
            Return False
        ElseIf (txtPassword.Text = "") Then
            MessageBox.Show("Please enter your password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtPassword.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub txtUsername_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUsername.LostFocus
        If txtUsername.Text <> "" Then
            txtUsername.Text = txtUsername.Text.PadLeft(5, "0")
        End If
    End Sub
End Class