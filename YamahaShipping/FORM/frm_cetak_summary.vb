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

Public Class frm_cetak_summary

    Public cyc As String = ""
    Public tgl As String = ""

    Private Sub myCancel()
        txtLabel.Text = ""
        txtLabel.Focus()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If txtLabel.Text = "" Then
            MessageBox.Show("Pilih DS terlebih dahulu !!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            printSummary()
        End If
    End Sub

    Private Sub frm_cetak_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
    End Sub

    Private Sub dgData_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub printSummary()
        Dim cryRpt1 As New cetak_Summary_Manual()

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = cyc
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@cycle")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = tgl
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@tgl")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        cryRpt1.SetDatabaseLogon("sa", "r3m04aaij", "HRD", "dotNET")
        Dim prd As New System.Drawing.Printing.PrintDocument

        cryRpt1.PrintOptions.PrinterName = "HP LaserJet M101-M106"

        cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
        cryRpt1.PrintToPrinter(1, True, 0, 0)
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnFind2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind2.Click
        frm_LOV_summary.ShowDialog()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        myCancel()
    End Sub
End Class