Imports System.Data
Imports System.Data.Sql
Imports System
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Math
Imports ThoughtWorks.QRCode.Codec

Public Class frm_cetakTag

    Public absId As String = ""
    Public kres As String = ""
    Public poAAIJ As String = ""
    Public pn_desc As String = ""
    Public pn_cust As String = ""
    Public qty_del As String = ""

    Private Sub myCancel()
        txtLabel.Text = ""
        txtLabel.Focus()
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If txtLabel.Text = "" Then
            MessageBox.Show("Pilih Tag terlebih dahulu !!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            frm_barcode.qrcode(absId)
            'frm_barcode.printSafetyCase(absId)
        End If
    End Sub

    Private Sub frm_cetak_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnFind2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind2.Click
        frm_LOV.ShowDialog()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        myCancel()
    End Sub
End Class