
Imports System.Data.SqlClient
Imports System.Data

Public Class Class_Login
    Public Function getLogin(ByVal username As String, ByVal password As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            myDA = New SqlDataAdapter("SELECT * FROM aavh_login WHERE login_npk='" & username & "'  and login_aktif='T'", connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getDataLogin(ByVal username As String) As DataTable
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            myDA = New SqlDataAdapter("SELECT * FROM aavh_login WHERE login_npk='" & username & "'", connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertUser(ByVal npk As String, ByVal nama As String, ByVal password As String, ByVal admin As String, ByVal aktif As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = "insert into aavh_login (login_npk, login_nama, login_password, login_admin, login_aktif) values ('" & npk & "','" & nama & "',md5('" & password & "'),'" & admin & "','" & aktif & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateUser(ByVal npk As String, ByVal nama As String, ByVal password As String, ByVal admin As String, ByVal aktif As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = "update aavh_login set login_nama='" & nama & "', login_password=md5('" & password & "'), login_admin='" & admin & "', login_aktif='" & aktif & "' where login_npk='" & npk & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getAllLogin() As DataTable
        Try
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            myDA = New SqlDataAdapter("SELECT * FROM aavh_login order by login_npk asc", connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT

        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function
End Class
