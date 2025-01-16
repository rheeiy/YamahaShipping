Imports System.Data.OleDb
Imports System.Data
'Imports Microsoft.
Public Class connectQAD

    Protected tblPengguna = New DataTable
    Protected SQL As String
    Protected Cn As OleDbConnection
    Protected Cmd As OleDbCommand
    Protected Da As OleDbDataAdapter
    Protected Ds As DataSet
    Protected Dt As DataTable
    Public Function OpenConn() As Boolean
        Cn = New OleDb.OleDbConnection("DSN=ee_qad_live;HOST=192.168.0.190;DB=mfgtdw;UID=sysprogress;PWD=s;PORT=50004")
        'Cn = New OleDbConnection("server=192.168.0.44; Data Source=ee_qad_live; User ID=sysprogress; Password=s")
        Cn.Open()
        If Cn.State <> ConnectionState.Open Then
            MsgBox("nok")
            Return False
        Else
            MsgBox("ok")
            Return True
        End If
    End Function
    Public Sub CloseConn()
        If Not IsNothing(Cn) Then
            Cn.Close()
            Cn = Nothing
        End If
    End Sub
    Public Function ExecuteQuery3(ByVal Query As String) As DataTable
        If Not OpenConn() Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed")
            Return Nothing
            Exit Function
        End If

        Cmd = New OleDbCommand(Query, Cn)
        Da = New OleDbDataAdapter
        Da.SelectCommand = Cmd

        Ds = New Data.DataSet
        Da.Fill(Ds)

        Dt = Ds.Tables(0)

        Return Dt

        Dt = Nothing
        Ds = Nothing
        Da = Nothing
        Cmd = Nothing

        CloseConn()

    End Function
    Public Sub ExecuteNonQuery3(ByVal Query As String)
        If Not OpenConn() Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed..!!")
            Exit Sub
        End If

        Cmd = New OleDbCommand
        Cmd.Connection = Cn
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = Query
        Cmd.ExecuteNonQuery()
        Cmd = Nothing
        CloseConn()
    End Sub

End Class
