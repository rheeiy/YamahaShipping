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
Imports System.Configuration
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.IO
Imports System.Xml

Public Class frm_dsOES
    Dim kanban As String = ""
    Dim itemCard As String = ""
    Dim trans_id As String = ""
    Dim pn_cust As String = ""
    Dim pn_cust2 As String = ""
    Dim po As String = ""
    Dim snp As String = ""
    Dim snp_cetak As String = ""
    Dim label As String = ""
    Dim pn_aaij As String = ""
    Public count_scan As Decimal = 0
    Public sum_status As Boolean = False
    Public order_qty As Decimal = 0
    Public num_to As Integer = 0
    Public outstd_qty As Decimal = 0
    Public jmlScan As Integer = 0
    Public jmlScanTotal As Integer = 0
    Public timeUtc As Date = Date.UtcNow
    Public cstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time")
    Public cstTime As Date = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone)
    Public shiftA As DateTime
    Public shiftB As DateTime
    Public pn_his As String = ""
    Public cycle As String = ""
    Public lot As String = ""
    Public lotTime As String = ""
    Public lotDB As String = ""
    Public pnAaijTag As String = ""
    'FOR DS
    Dim SJ As String = ""
    Dim pn_DS As String = ""
    Dim kres As String = ""
    Dim pn_desc As String = ""
    Dim userYim As String = ""
    Dim poAAIJ As String = ""
    Dim poLine As String = ""
    Dim shipto As String = ""
    Dim custpart As String = ""
    Dim myQX As New maintainSalesOrderShipper_OES()
    Dim myQX2 As New confirmDS()
    Public qty_del As Decimal = 0
    Public blnAwal As String = ""
    Public blnAkhir As String = ""
    Public statShipping As Boolean = True
    Public smallpart As Boolean = False
    Public idprint As Integer = 0
    Public hitCreateDS As Integer = 0
    Dim part_weight_act As Decimal = 0
    Dim hasil As String = ""
    Dim nilai As Decimal = 0.0
    Dim ms_weight_min As Decimal = 0
    Dim ms_weight_max As Decimal = 0
    Dim part_deskripsi As String = ""
    Dim tipeTag As String = ""
    Dim formatQR As String = ""
    Dim snp_timbang As Decimal = 0
    Dim ms_1timbang As Decimal = 0
    Dim ms_totTimbang As Decimal = 0
    Dim snpItemCard As Decimal = 0
    Dim tototWeight As Decimal = 0
    Dim ms_weight_min_ms As Decimal = 0
    Dim ms_weight_max_ms As Decimal = 0
    Dim counting As Integer = 1
    Dim paletize As String = ""
    'WSA
    Dim elemList As XmlNodeList = Nothing
    Dim elementXml = New XmlDocument()
    Dim rs As String = ""
    'END WSA
    Private Sub txtPalet_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPalet.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            If txtPalet.Text = "PRINT DS" Then
                insertDS()
                txtItemCard.Text = ""
                txtItemCard.Focus()
            End If
        End If
    End Sub

    Private Sub myCancel()
        txtItemCard.Text = ""
        txtPalet.Text = ""
        txtPalet.Focus()
    End Sub
    Private Sub txtItemCard_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCard.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            txtItemCard.Text = txtItemCard.Text.ToUpper
            If txtPalet.Text = "PRINT DS" And txtItemCard.Text <> "FINISH DS" Then
                insertDSDet()
            ElseIf txtItemCard.Text = "FINISH DS" Then
                createDS()
            Else
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                myCancel()
            End If
        End If
    End Sub

    Private Sub insertDS()
        Dim myTem As New queryDB_OES
        Dim myDT As New DataTable
        myTem.insertDS("-", "", txtItemCard.Text)
        idprint = myTem.getNoDs
    End Sub

    Private Sub createDS()
        Dim myTem As New queryDB_OES
        Dim myDT As New DataTable
        Dim myTem2 As New queryDB_OES
        Dim myDT2 As New DataTable
        userYim = "PRPC"
        poAAIJ = Format(cstTime, "MMyy") & userYim
        If myQX.requestInbound(idprint) = True Then
            'myQX.absId = "Test"
            myDT = myTem.createDS2(idprint)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                For i As Integer = 0 To myDT.Rows.Count - 1
                    trans_id = myDT.Rows(i).Item("idorder").ToString

                    If myTem2.updateLoadingDet(trans_id) = False Then
                        myTem2.updateNGDS(myQX.absId)
                    Else
                        If count_scan = order_qty Then
                            myTem2.updateLoading(trans_id)
                        End If
                    End If
                    'myCancelDS()
                Next
                myTem2.updateDS(idprint, myQX.absId, "OK", myQX.note_xml, myQX.keterangan_qx, myQX.xmlSend)
                myTem2.updateDSDet(idprint, myQX.absId)
                'myTem2.updateShipDate("s" & myQX.absId, pn_aaij, poAAIJ)
                'myTem2.insertDSCustom(myQX.absId, shipto)
                printBon(myQX.absId)
                My.Computer.Audio.Play(My.Resources.surat_jalan_oke, AudioPlayMode.Background)
                myCancel()
                'myCancelDS()
            End If
        Else
            myTem2.updateDS(idprint, "-", "", myQX.note_xml, myQX.keterangan_qx, myQX.xmlSend)
            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
            myCancel()
            Exit Sub
        End If
    End Sub

    Private Sub insertDSDet()
        Dim myTem As New queryDB_OES
        Dim myDT As New DataTable
        itemCard = txtItemCard.Text
        If txtItemCard.Text <> "FINISH DS" Then
            ''vwOrderDS(itemCard)
        End If

        myDT = myTem.cekDataDS(itemCard)
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            If itemCard <> "" Then
                trans_id = myDT.Rows(0).Item("id").ToString
                pn_aaij = myDT.Rows(0).Item("ms_item_nbr").ToString
                po = myDT.Rows(0).Item("ym_po").ToString
                snp = myDT.Rows(0).Item("ms_part_snp").ToString
                sum_status = myDT.Rows(0).Item("ym_stat_scan")
                order_qty = myDT.Rows(0).Item("ym_order_qty").ToString
                count_scan = myDT.Rows(0).Item("ym_qty_scan").ToString
                outstd_qty = myDT.Rows(0).Item("ym_order_qty") - myDT.Rows(0).Item("ym_qty_scan")
                cycle = myDT.Rows(0).Item("ym_cycle").ToString
                pn_cust = myDT.Rows(0).Item("ms_part_cust").ToString
                'qty_del = count_scan

                'If pn_cust.Substring(pn_cust.Length - 2, 1) = "J0" Then
                If pn_cust.Contains("-J0") Then
                    pn_aaij = pn_aaij & "-P"
                    cycle = "EXPORT"
                ElseIf pn_cust.Contains("-X0") Then
                    pn_aaij = pn_aaij & "-X"
                    cycle = "EXPORT"
                End If
                pn_desc = myDT.Rows(0).Item("ms_part_desc").ToString & "(" & pn_aaij & ")"

                kres = "#" & itemCard.Substring(37, 5)
                'userYim = itemCard.Substring(32, 4)
                userYim = "PRPC"
                poAAIJ = Format(cstTime, "MMyy") & userYim

                Dim query As String = "SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & poAAIJ & "' AND sod_part='" & pn_aaij & "' " &
                                            " AND sod_domain='AAIJ' WITH(nolock)"
                Try
                    elementXml = myTem.getwsa(query)
                    elemList = elementXml.GetElementsByTagName("f_val")
                    If elemList.Count > 0 Then
                        For z As Integer = 0 To elemList.Count - 1
                            rs = elementXml.GetElementsByTagName("f_val")(z).InnerText.Replace("||", "#")
                            poLine = rs.Split("#"c)(0)
                        Next
                    End If
                Catch ex As Exception
                    MsgBox("WSA problem " & ex.ToString)
                End Try

                'myDT = myTem.getLinePO(poAAIJ, pn_aaij)
                'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                '    poLine = myDT.Rows(0).Item("sod_line").ToString
                'Else
                '    My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                '    MsgBox("Data PO tidak ada!")
                '    myCancel()
                '    Exit Sub
                'End If
                shipto = "2" & userYim
                myDT = myTem.getTotScan(trans_id)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    qty_del = myDT.Rows(0).Item("jml").ToString
                    If qty_del <> order_qty Then
                        My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
                        frm_dialog.ShowDialog()
                        If frm_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
                            myCancel()
                            Exit Sub
                        End If
                    End If

                    If myTem.insertDSDet(trans_id, idprint, kres, poAAIJ, pn_desc, pn_cust, qty_del, txtItemCard.Text, cycle, pn_aaij) = True Then
                        My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                        txtItemCard.Text = ""
                        txtItemCard.Focus()
                        Exit Sub
                    Else
                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                        MsgBox("Gagal insert DS!")
                        txtItemCard.Text = ""
                        txtItemCard.Focus()
                        Exit Sub
                    End If
                Else
                    My.Computer.Audio.Play(My.Resources.DSsudahbarcode, AudioPlayMode.Background)
                    myCancel()
                End If
            End If
        Else
            My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
            myCancel()
        End If
    End Sub

    Private Sub printBon(idds)
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

        cryRpt1.SetDatabaseLogon("sa", "r3m04aaij", "HRD", "dotNET")
        Dim prd As New System.Drawing.Printing.PrintDocument

        cryRpt1.PrintOptions.PrinterName = prd.PrinterSettings.PrinterName

        'cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        'cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        cryRpt1.PrintToPrinter(1, True, 0, 0)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub frm_dsOES_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
    End Sub
End Class