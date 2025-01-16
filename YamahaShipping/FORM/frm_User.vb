Imports System
Imports System.Security.Cryptography
Imports System.Text

Public Class frm_User
    Dim npk As String = ""
    Dim nama As String = ""
    Dim password As String = ""
    Dim admin As String = ""
    Dim aktif As String = ""

    Private Sub save()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        npk = txtNPK.Text
        nama = txtNama.Text
        password = txtPassword.Text
        admin = cbAdmin.SelectedItem
        If admin = "Yes" Then
            admin = "T"
        Else
            admin = "F"
        End If
        aktif = cbAktif.SelectedItem
        If aktif = "Yes" Then
            aktif = "T"
        Else
            aktif = "F"
        End If
        Using md5hash As MD5 = MD5.Create
            password = GetMd5Hash(md5hash, password)
        End Using
        If myTem.insertUser(npk, nama, password, admin, aktif) = False Then
            MsgBox("Gagal Simpan !")
        Else
            MsgBox("Berhasil disimpan !")
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

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        save()
    End Sub
End Class