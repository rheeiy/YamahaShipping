Imports System
Imports System.Data
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient


Public Class connectDB
    'CONNECT SQL BGP
    Public Shared sqlConnectionBGP As New SqlConnection("server=DESKTOP-IQLVVGB; uid=sa; pwd=123; database=DELIVERY")
    Public Shared Sub connectSQLBGP()
        sqlConnectionBGP.Open()
    End Sub
    Public Shared Sub disconnectSQLBGP()
        sqlConnectionBGP.Close()
    End Sub
    Public Shared Sub executeDBSQLBGP(ByVal sqlQuery As String)
        If sqlConnectionBGP.State = ConnectionState.Open Then
            Dim cmd As New SqlCommand(sqlQuery, sqlConnectionBGP)
            cmd.ExecuteNonQuery()
            sqlConnectionBGP.Close()
        Else
            MessageBox.Show("Failed to connect SQL Database", "Error Connection", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    'CONNECT SQL
    Public Shared sqlConnection As New SqlConnection("server=DESKTOP-IQLVVGB; uid=sa; pwd=123; database=DELIVERY")
    Public Shared Sub connectSQL()
        sqlConnection.Open()
    End Sub
    Public Shared Sub disconnectSQL()
        sqlConnection.Close()
    End Sub

    Public Shared Sub executeDBSQL(ByVal sqlQuery As String)
        If sqlConnection.State = ConnectionState.Open Then
            Dim cmd As New SqlCommand(sqlQuery, sqlConnection)
            cmd.ExecuteNonQuery()
            sqlConnection.Close()
        Else
            MessageBox.Show("Failed to connect SQL Database", "Error Connection", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class