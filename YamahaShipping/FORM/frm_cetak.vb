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

Public Class frm_cetak

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

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If txtLabel.Text = "" Then
            MessageBox.Show("Pilih DS terlebih dahulu !!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            frm_barcode.qrcode(absId)
            printBon(absId)
        End If
    End Sub

    Private Sub frm_cetak_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
    End Sub

    Private Sub dgData_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Public Sub printBon(idds)
        Dim cryRpt1 As New cetak_DS_OES()

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = idds
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@ds_no")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = frm_barcode.GetSetting("pic")
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@pic")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        cryRpt1.SetDatabaseLogon("usrdotnet", "r3m04tdw", "dbserver2.akebono-astra.co.id", "DELIVERY")
        Dim prd As New System.Drawing.Printing.PrintDocument

        cryRpt1.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName

        'cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        'cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        cryRpt1.PrintToPrinter(1, True, 0, 0)
    End Sub

    Private Sub printBon()
        Dim cryRpt1 As New cetak_DS()

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue


        crParameterDiscreteValue.Value = absId
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@ds_no")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = kres
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@kres")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = poAAIJ
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@po")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = pn_desc
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@part_desc")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = pn_cust
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@model")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = qty_del
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@jumlah")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = frm_barcode.GetSetting("pic")
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@pic")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        cryRpt1.SetDatabaseLogon("sa", "r3m04aaij", "dbserver2.akebono-astra.co.id", "dotNET")
        Dim prd As New System.Drawing.Printing.PrintDocument

        cryRpt1.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName

        'cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        'cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        cryRpt1.PrintToPrinter(1, True, 0, 0)
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

    Private Sub txtLabel_TextChanged(sender As Object, e As EventArgs) Handles txtLabel.TextChanged

    End Sub
End Class