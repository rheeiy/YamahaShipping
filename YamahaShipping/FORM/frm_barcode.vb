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
Imports System.Data.SqlClient
Imports System.Xml

Public Class frm_barcode
    Dim kanban As String = ""
    Dim itemCard As String = ""
    Dim trans_id As String = ""
    Dim pn_cust As String = ""
    Dim item_j0 As String = ""
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
    Public tagProduksi As String = ""
    Public kanbanSafety As String = ""
    Public item_no As String = ""
    Public kanban_qty As Decimal
    Public kanban_sft As String
    Public full_item As String
    Public snpSft As Integer
    Public usercode As String
    Public tglProd As Date

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
    Dim myQX As New maintainSalesOrderShipper()
    Dim myQX2 As New confirmDS()
    Public qty_del As Decimal = 0
    Public blnAwal As String = ""
    Public blnAkhir As String = ""
    Dim paletize As String = ""
    Public smallpart As Boolean = False
    Dim part_deskripsi As String = ""
    Dim tipeTag As String = ""
    Public lblPc As String = ""
    Public dtp_delivery As Date = cstTime.ToString("yyyy-MM-dd")
    Public pic_ds As String = ""
    Public printer_casemark As String = ""
    Public set_conform As String = ""
    Public effdate As String = ""
    Public effdateSafety As String = ""
    Dim Locfrom As String = "P2-FG"
    Dim Locto As String = "PPC-FG"
    Public idprint As Integer = 0
    Public sh_ds As String = ""
    Public bypassQad As String = ""
    Dim userTrace As String = ""
    Dim trigger As Integer = 1
    'WSA
    Dim elemList As XmlNodeList = Nothing
    Dim elementXml = New XmlDocument()
    Dim rs As String = ""
    'END WSA
    Private Sub vwOrder(kanban, itemCard)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        Try
            dgOrder.Rows.Clear()
            dgOrder.Columns.Clear()
            myDT = myTemp.getOrder(kanban, itemCard)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                dgOrder.Columns.Add("NO", "No. ")
                dgOrder.Columns.Add("po_no", "PO NO")
                dgOrder.Columns.Add("del_date", "DEL. DATE")
                dgOrder.Columns.Add("part_no", "PART NO")
                dgOrder.Columns.Add("part_desc", "PART DESC")
                dgOrder.Columns.Add("qty_order", "QTY ORD") '6
                dgOrder.Columns.Add("snp", "SNP")
                dgOrder.Columns.Add("qty_scan", "QTY SCAN")
                dgOrder.Columns.Add("qty_outstanding", "QTY OTST")
                dgOrder.Columns.Add("pp", "Postpone")
                dgOrder.Columns.Add("sf", "Safety")
                dgOrder.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns(0).Width = 30
                dgOrder.Columns(1).Width = 50
                dgOrder.Columns(2).Width = 70
                dgOrder.Columns(3).Width = 100
                dgOrder.Columns(4).Width = 110
                dgOrder.Columns(5).Width = 40
                dgOrder.Columns(6).Width = 40
                dgOrder.Columns(7).Width = 40
                dgOrder.Columns(8).Width = 40
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("snp").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("qty_scan").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("qty_outstanding").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Dim nomer As Integer = 1
                dgOrder.Rows.Add(myDT.Rows.Count)
                For i As Integer = 0 To myDT.Rows.Count - 1
                    dgOrder.Item(0, i).Value = nomer
                    dgOrder.Item(1, i).Value = myDT.Rows(i).Item("ym_po")
                    dgOrder.Item(2, i).Value = Format(myDT.Rows(i).Item("ym_del_duedate"), "yyyy-MM-dd")
                    dgOrder.Item(3, i).Value = myDT.Rows(i).Item("ym_item_nbr")
                    dgOrder.Item(4, i).Value = myDT.Rows(i).Item("ym_item_desc")
                    dgOrder.Item(5, i).Value = myDT.Rows(i).Item("ym_order_qty")
                    dgOrder.Item(6, i).Value = myDT.Rows(i).Item("ms_part_snp")
                    dgOrder.Item(7, i).Value = myDT.Rows(i).Item("ym_qty_scan")
                    dgOrder.Item(8, i).Value = myDT.Rows(i).Item("ym_order_qty") - myDT.Rows(i).Item("ym_qty_scan")
                    dgOrder.Item(9, i).Value = myDT.Rows(i).Item("ym_postpone").ToString
                    dgOrder.Item(10, i).Value = myDT.Rows(i).Item("ym_safety").ToString
                    nomer += 1
                Next
            End If
        Catch ex As Exception
            ''MsgBox(ex.ToString)
        End Try

    End Sub

    'FOR DS
    Private Sub vwOrderDS(itemCard)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        Try
            dgOrder.Rows.Clear()
            dgOrder.Columns.Clear()
            myDT = myTemp.getOrderDS("", itemCard)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                dgOrder.Columns.Add("NO", "No. ")
                dgOrder.Columns.Add("po_no", "PO NO")
                dgOrder.Columns.Add("del_date", "DEL. DATE")
                dgOrder.Columns.Add("part_no", "PART NO")
                dgOrder.Columns.Add("part_desc", "PART DESC")
                dgOrder.Columns.Add("qty_order", "QTY ORD") '6
                dgOrder.Columns.Add("snp", "SNP")
                dgOrder.Columns.Add("qty_scan", "QTY SCAN")
                dgOrder.Columns.Add("qty_outstanding", "QTY OTST")
                dgOrder.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns(0).Width = 30
                dgOrder.Columns(1).Width = 40
                dgOrder.Columns(2).Width = 60
                dgOrder.Columns(3).Width = 100
                dgOrder.Columns(4).Width = 120
                dgOrder.Columns(5).Width = 50
                dgOrder.Columns(6).Width = 50
                dgOrder.Columns(7).Width = 50
                dgOrder.Columns(8).Width = 50
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("snp").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("qty_scan").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns("qty_outstanding").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Dim nomer As Integer = 1
                dgOrder.Rows.Add(myDT.Rows.Count)
                For i As Integer = 0 To myDT.Rows.Count - 1
                    dgOrder.Item(0, i).Value = nomer
                    dgOrder.Item(1, i).Value = myDT.Rows(i).Item("ym_po")
                    dgOrder.Item(2, i).Value = Format(myDT.Rows(i).Item("ym_del_duedate"), "yyyy-MM-dd")
                    dgOrder.Item(3, i).Value = myDT.Rows(i).Item("ym_item_nbr")
                    dgOrder.Item(4, i).Value = myDT.Rows(i).Item("ym_item_desc")
                    dgOrder.Item(5, i).Value = myDT.Rows(i).Item("ym_order_qty")
                    dgOrder.Item(6, i).Value = myDT.Rows(i).Item("ms_part_snp")
                    dgOrder.Item(7, i).Value = myDT.Rows(i).Item("ym_qty_scan")
                    dgOrder.Item(8, i).Value = myDT.Rows(i).Item("ym_order_qty") - myDT.Rows(i).Item("ym_qty_scan")
                    nomer += 1
                Next
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub vwLot(trans_id)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        dgLot.Rows.Clear()
        dgLot.Columns.Clear()
        myDT = myTemp.getLOTTr(trans_id)
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            dgLot.Columns.Add("NO", "No. ")
            dgLot.Columns.Add("po_no", "PO NO")
            dgLot.Columns.Add("lot", "LOT")
            dgLot.Columns.Add("prod_date", "PROD DATE")
            dgLot.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgLot.Columns("prod_date").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgLot.Columns(0).Width = 40
            dgLot.Columns(1).Width = 60
            dgLot.Columns(2).Width = 120
            dgLot.Columns(3).Width = 120
            'dgLot.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Dim nomer As Integer = 1
            dgLot.Rows.Add(myDT.Rows.Count)
            For i As Integer = 0 To myDT.Rows.Count - 1
                dgLot.Item(0, i).Value = nomer
                dgLot.Item(1, i).Value = myDT.Rows(i).Item("ym_po")
                dgLot.Item(2, i).Value = myDT.Rows(i).Item("tr_lot")
                dgLot.Item(3, i).Value = Format(myDT.Rows(i).Item("tr_lot_time"), "yyyy-MM-dd")
                nomer += 1
            Next
        End If
    End Sub

    Private Sub vwSafety()
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        dgLot.Rows.Clear()
        dgLot.Columns.Clear()
        myDT = myTemp.getLotSafety()
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            dgLot.Columns.Add("NO", "No. ")
            dgLot.Columns.Add("lot", "LOT")
            dgLot.Columns.Add("creadate", "PROD DATE")
            dgLot.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgLot.Columns("creadate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgLot.Columns(0).Width = 40
            dgLot.Columns(1).Width = 150
            dgLot.Columns(2).Width = 150
            'dgLot.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Dim nomer As Integer = 1
            dgLot.Rows.Add(myDT.Rows.Count)
            For i As Integer = 0 To myDT.Rows.Count - 1
                dgLot.Item(0, i).Value = nomer
                dgLot.Item(1, i).Value = myDT.Rows(i).Item("lot")
                dgLot.Item(2, i).Value = Format(myDT.Rows(i).Item("creadate"), "yyyy-MM-dd")
                nomer += 1
            Next
        End If
    End Sub

    Private Sub frm_barcode_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
        lblStatus.Text = "Ready To Scan!"
        lblPc = GetIPv4Address()
        dtp_del.Value = cstTime.ToString("yyyy-MM-dd")
        dtp_delivery = Format(dtp_del.Value, "yyyy-MM-dd")
        vwOrder(kanban, itemCard)
        printer_casemark = GetSetting("printer_casemark")
        pic_ds = GetSetting("pic")
        set_conform = GetSetting("conform")
        sh_ds = GetSetting("pic")
        bypassQad = GetSetting("qad")

        If bypassQad.ToString = "True" Then
            lblQAD.Text = "ON"
        Else
            lblQAD.Text = "OFF"
        End If
        'BackgroundWorker1.RunWorkerAsync()
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Public Function GetSetting(ByVal sName As String) As String
        For Each s As String In IO.File.ReadAllLines("Settings.txt")
            If s.StartsWith(sName) Then
                Return s.Substring(s.IndexOf("=") + 1)
            End If
        Next
        Return String.Empty
    End Function

    Private Sub myCancel()
        lblLogin.Text = frm_Login.login_name
        dgLot.Visible = False

        'dgOrder.Refresh()
        txtKanban.Text = ""
        txtItemCard.Text = ""
        txtPartDesc.Text = ""
        txtQtyScan.Text = ""
        txtQtyOrder.Text = ""
        txtQtyOutsd.Text = ""
        txtCycle.Focus()
        kanban = ""
        itemCard = ""
        trans_id = ""
        po = ""
        snp = ""
        snp_cetak = ""
        count_scan = 0
        sum_status = False
        order_qty = 0
        num_to = 0
        vwOrder(kanban, itemCard)
        jmlScan = 0
        jmlScanTotal = 0
        txtNoScan.Text = ""
        txtCycle.Text = ""
        pn_his = ""
        pn_cust = ""
        label = ""
        pn_aaij = ""
        cycle = ""
        lot = ""
        lotTime = ""
        lotDB = ""
        pnAaijTag = ""
        pn_cust2 = ""
        tmrSwitch.Enabled = False
        smallpart = False
        part_deskripsi = ""
        tipeTag = ""
    End Sub


    Private Sub cekData()
        If reset() = False Then
            Dim myTem As New queryDB
            Dim myDT As New DataTable

            itemCard = txtItemCard.Text
            label = txtCycle.Text
            'vwOrder(kanban, itemCard)
            myDT = myTem.cekDataItemCard(itemCard)
            If myDT IsNot Nothing AndAlso myDT.Rows.Count > 0 Then
                sum_status = myDT.Rows(0).Item("ym_stat_scan")
                If sum_status.Equals(True) Then
                    My.Computer.Audio.Play(My.Resources.item_card_lunas, AudioPlayMode.Background)
                    lblStatus.Text = "Item Card Sudah Lunas"
                    tmrWarning2.Enabled = True
                    myCancel()
                    Exit Sub
                End If
            End If

            myDT = myTem.cekData(itemCard, label)
            If myDT IsNot Nothing AndAlso myDT.Rows.Count > 0 Then
                If kanban = "" And itemCard <> "" And label <> "" Then
                    trans_id = myDT.Rows(0).Item("id").ToString
                    pn_aaij = myDT.Rows(0).Item("ms_item_nbr").ToString
                    kanban = pn_aaij
                    po = myDT.Rows(0).Item("ym_po").ToString
                    usercode = myDT.Rows(0).Item("ym_user_code").ToString
                    snp = myDT.Rows(0).Item("ms_part_snp").ToString
                    sum_status = myDT.Rows(0).Item("ym_stat_scan")
                    order_qty = myDT.Rows(0).Item("ym_order_qty").ToString
                    count_scan = myDT.Rows(0).Item("ym_qty_scan").ToString
                    outstd_qty = myDT.Rows(0).Item("ym_order_qty") - myDT.Rows(0).Item("ym_qty_scan")
                    txtPartDesc.Text = myDT.Rows(0).Item("ym_item_desc").ToString
                    txtQtyOrder.Text = myDT.Rows(0).Item("ym_order_qty").ToString
                    pn_cust = myDT.Rows(0).Item("ms_part_cust").ToString
                    pn_cust2 = pn_cust.Replace("-J0", "").Replace("-80", "").Replace("-X0", "").Replace("-", "")
                    txtQtyOutsd.Text = outstd_qty
                    txtQtyScan.Text = count_scan
                    SummaryScan("1")
                    smallpart = myDT.Rows(0).Item("ms_smallpart").ToString
                    part_deskripsi = myDT.Rows(0).Item("ms_part_desc").ToString
                    tipeTag = myDT.Rows(0).Item("ms_part_desc1").ToString

                    If sum_status = False And Not pn_cust.Contains("J0") Then
                        If count_scan >= order_qty Then
                            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                            lblStatus.Text = "Scan Melebihi Order !"
                            tmrWarning2.Enabled = True
                            myCancel()
                        Else
                            'myTem.nyalainLampu(pn_cust2)
                            ''dgLot.Visible = True
                            ''PictureBox1.ImageLocation = "http://delivery.akebono-astra.co.id/asn/SPISJ0/" & pn_aaij & ".jpg"
                            If outstd_qty >= snp Then
                                snp = snp
                                If (smallpart = "True" Or smallpart = True) And usercode = "7618" Then
                                    lotDB = "OES"
                                    saveTransaksi()
                                Else
                                    txtKanban.Text = ""
                                    txtKanban.Focus()
                                    'vwOrder(kanban, itemCard)
                                End If
                            Else
                                'My.Computer.Audio.Play("konfirmasi.wav")
                                snp = outstd_qty
                                If (smallpart = "True" Or smallpart = True) And usercode = "7618" Then
                                    lotDB = "OES"
                                    saveTransaksi()
                                Else
                                    txtKanban.Text = ""
                                    txtKanban.Focus()
                                    'vwOrder(kanban, itemCard)
                                End If
                            End If
                        End If
                    Else
                        My.Computer.Audio.Play(My.Resources.item_card_tidak_ditemukan, AudioPlayMode.Background)
                        lblStatus.Text = "Item card salah!"
                        myCancel()
                        tmrWarning2.Enabled = True
                    End If
                ElseIf kanban <> "" And itemCard <> "" And label <> "" Then
                    'If myTem.cekTagProd(txtKanban.Text).ToString <> "" And (txtKanban.Text.Length = 9 Or txtKanban.Text.Length = 11) Then
                    '    My.Computer.Audio.Play("sudah-di-scan.wav")
                    '    myCancel()
                    '    Exit Sub
                    'End If
                    'myDT = myTem.getLot(lot, pn_cust)
                    'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    '    lotDB = myDT.Rows(0).Item("db").ToString
                    '    lotTime = myDT.Rows(0).Item("creadate")
                    '    saveTransaksi()
                    '    Exit Sub
                    'Else

                    'End If
                    If lot = pn_aaij And itemCard <> "" And label <> "" Then
                        saveTransaksi()
                    ElseIf lot.Contains("(1)") And itemCard <> "" And label <> "" Then
                        lotDB = lot.Substring(lot.IndexOf("(3)") + 3, lot.IndexOf("(4)", lot.IndexOf("(3)") + 1) - lot.IndexOf("(3)") - 3)
                        lotTime = lot.Substring(lot.IndexOf("(4)") + 3, lot.IndexOf("(5)", lot.IndexOf("(4)") + 1) - lot.IndexOf("(4)") - 3)
                        lotTime = Replace(lotTime, ".", ":")
                        pnAaijTag = lot.Substring(lot.IndexOf("(2)") + 3, lot.IndexOf("(3)", lot.IndexOf("(2)") + 1) - lot.IndexOf("(2)") - 3)
                        lot = lot.Substring(lot.IndexOf("(1)") + 3, lot.IndexOf("(2)", lot.IndexOf("(1)") + 1) - lot.IndexOf("(1)") - 3)

                        If myTem.cekTagProd(lot, lotDB).ToString <> "" OrElse myTem.cekTagProdLog(lot).ToString <> "" Then
                            My.Computer.Audio.Play(My.Resources.tag_sudah_discan, AudioPlayMode.Background)
                            lblStatus.Text = "Tag Salah, Tag sudah di pulling!"
                            myCancel()
                            Exit Sub
                        End If

                        If pnAaijTag = pn_aaij Then
                            kanban_sft = myTem.cekTag(txtKanban.Text).ToString

                            saveTransaksi()
                            'vwOrder(kanban, itemCard)
                        Else
                            My.Computer.Audio.Play(My.Resources.item_card_dan_tag_beda, AudioPlayMode.Background)
                            lblStatus.Text = "Tag Salah, Tag dan Item Card Beda!"
                            myCancel()
                            tmrWarning2.Enabled = True
                        End If
                    ElseIf lot.Length = 9 Then
                        myDT = myTem.getLot(lot, pn_cust)
                        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                            lotDB = myDT.Rows(0).Item("db").ToString
                            lotTime = myDT.Rows(0).Item("creadate")
                            saveTransaksi()
                            Exit Sub
                        End If
                    Else
                        My.Computer.Audio.Play(My.Resources.tag_tidak_ditemukan, AudioPlayMode.Background)
                        lblStatus.Text = "Tag tidak ditemukan, periksa tag yang discan!"
                        myCancel()
                        tmrWarning2.Enabled = True
                    End If
                End If
            Else
                My.Computer.Audio.Play(My.Resources.item_card_tidak_ditemukan, AudioPlayMode.Background)
                lblStatus.Text = "Data Tidak Sesuai"
                tmrWarning2.Enabled = True
                If txtItemCard.Text = "DANDORI" Or txtKanban.Text = "DANDORI" Or txtCycle.Text = "DANDORI" Then
                    myCancel()
                ElseIf txtCycle.Text <> "" And txtItemCard.Text <> "" And txtKanban.Text = "" Then
                    txtItemCard.Text = ""
                    txtItemCard.Focus()
                    'vwOrder("", itemCard)
                ElseIf txtItemCard.Text <> "" And txtKanban.Text <> "" And txtCycle.Text <> "" Then
                    txtKanban.Text = ""
                    txtItemCard.Text = ""
                    txtItemCard.Focus()
                    'vwOrder("", itemCard)
                Else
                    myCancel()
                End If
            End If
        End If
    End Sub

    Private Sub saveTransaksi()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        Dim myQX As New QXtendInboundTransferSingleItem()
        'If tmrJeda.Enabled = False Then

        ''
        If myTem.updateCountScan(trans_id, itemCard, snp, "False", po, cycle) = True Then
            If myTem.insertTransaksi(trans_id, kanban, snp, cycle, lot, lotTime, lotDB, txtKanban.Text, paletize, txtItemCard.Text) = True Then
                ''
                'If myTem.insertTransaksi(trans_id, kanban, snp, cycle, lot, lotTime, lotDB, txtKanban.Text) = True Then

                If bypassQad.ToString = "True" Then
                    'start inbound
                    Dim idlast As String = myTem.getIDTR(trans_id)
                    Dim number As String = "SI" & cstTime.ToString("yyyy") & "-" & idlast
                    'Dim number As String = idlast
                    Dim statInbound As Boolean = False

                    If cekOrderQad(number) = False Then
                        If myQX.requestInbound(kanban, effdate, snp, number, Locfrom, Locto, trans_id) = False Then
                            statInbound = False
                        Else
                            statInbound = True
                        End If
                        If myTem.updateInbound(idlast, statInbound, myQX.xmlSend, myQX.note_xml, Locfrom, Locto, effdate, myQX.keterangan_qx) = False Then
                            MsgBox("Failed to UPDATE data QExtend to SQL Server !!")
                        End If
                    End If
                    'End inbound
                End If


                count_scan += snp
                lblStatus.Text = "Transaksi berhasil !"
                tmrWarning2.Enabled = True
                SummaryScan("2")
                Try

                    If count_scan = order_qty Then
                        My.Computer.Audio.Play(My.Resources.lunas, AudioPlayMode.Background)

                        'If kanban_sft = "" And lotDB <> "SAFETY" Then
                        'myTem.updateLog(pn_cust.Replace("-J0", "").Replace("-80", "").Replace("-X0", ""))
                        myTem.insertLog(lot, pn_cust, txtKanban.Text, "REG", "OUT", pnAaijTag)
                            myTem.updateLogStatus(lot, "")
                            'Try
                            '    Dim webClient As New System.Net.WebClient
                            '    Dim result As String = webClient.DownloadString("https://dashboard.akebono-astra.co.id/store/test.php?part=" & pnAaijTag & "&snp=" & snp & "&status=OUT")
                            'Catch ex As Exception

                            'End Try
                            'ElseIf lotDB.Contains("SAFETY") Then
                            '    myTem.updateStockSafetyKurang(pnAaijTag)
                            '    myTem.insertLog(lot, pn_cust, txtKanban.Text, "SFT", "OUT", pnAaijTag)
                            '    myTem.updateLogStatus(lot, "1")
                            'Else
                            '    myTem.updateStatusTag(txtKanban.Text)
                            '    myTem.updateStockSafetyKurang(pnAaijTag)
                            '    myTem.insertLog(lot, pn_cust, txtKanban.Text, "SFT", "OUT", pnAaijTag)
                            '    myTem.updateLogStatus(lot, "1")
                            'End If

                            myTem.updateCountScan(trans_id, itemCard, snp, "True", po, cycle)
                        myDT = myTem.getStatCase(trans_id)
                        If myDT.Rows.Count = 0 Then
                            'printBon()
                        End If
                        vwOrder(kanban, itemCard)
                        'myTem.matiinLampu(pn_cust2)
                        achievementShift()
                        myCancel()

                    Else
                        My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)

                        'If kanban_sft = "" And lotDB <> "SAFETY" Then
                        'myTem.updateLog(pn_cust.Replace("-J0", "").Replace("-80", "").Replace("-X0", ""))
                        myTem.insertLog(lot, pn_cust, txtKanban.Text, "REG", "OUT", pnAaijTag)
                            myTem.updateLogStatus(lot, "")
                        'Try
                        '    Dim webClient As New System.Net.WebClient
                        '    Dim result As String = webClient.DownloadString("https://dashboard.akebono-astra.co.id/store/test.php?part=" & pnAaijTag & "&snp=" & snp & "&status=OUT")
                        'Catch ex As Exception

                        'End Try
                        'ElseIf lotDB.Contains("SAFETY") Then
                        '    myTem.updateStockSafetyKurang(pnAaijTag)
                        '    myTem.insertLog(lot, pn_cust, txtKanban.Text, "SFT", "OUT", pnAaijTag)
                        '    myTem.updateLogStatus(lot, "1")
                        'Else
                        '    myTem.updateStatusTag(txtKanban.Text)
                        '    myTem.updateStockSafetyKurang(pnAaijTag)
                        '    myTem.insertLog(lot, pn_cust, txtKanban.Text, "SFT", "OUT", pnAaijTag)
                        '    myTem.updateLogStatus(lot, "1")
                        'End If

                        txtItemCard.Text = ""
                        txtItemCard.Focus()
                        txtKanban.Text = ""
                        lot = ""
                        kanban = ""
                        dgLot.Visible = True
                        dgLot.BringToFront()
                        vwLot(trans_id)
                    End If
                Catch ex As Exception
                    'MsgBox(ex.ToString)
                End Try

                'tmrJeda.Enabled = True
                ''
            Else
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                lblStatus.Text = "Gagal, coba scan kembali"
                txtKanban.Text = ""
                myTem.updateMinScan(trans_id, snp)
            End If
        Else
            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
            lblStatus.Text = "Gagal, coba scan kembali"
            txtKanban.Text = ""
        End If

        ''

        ''Else
        'My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
        'lblStatus.Text = "Gagal, coba scan kembali"
        'txtKanban.Text = ""
        'End If
    End Sub

    Private Sub SummaryScan(status)
        If status = "2" Then
            txtQtyOutsd.Text = order_qty - count_scan
            txtQtyScan.Text = count_scan
            jmlScan = (count_scan) / snp
            jmlScanTotal = Round(order_qty / snp, 0)
            txtNoScan.Text = jmlScan & " / " & jmlScanTotal
        Else
            txtQtyOutsd.Text = order_qty - count_scan
            txtQtyScan.Text = count_scan
            jmlScan = count_scan / snp
            jmlScanTotal = Round(order_qty / snp, 0)
            txtNoScan.Text = jmlScan & " / " & jmlScanTotal
        End If
    End Sub

    Private Sub txtNotStd_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            If txtCycle.Text <> "" Then
                saveTransaksi()
            End If
        End If
    End Sub

    Private Function reset()
        If txtKanban.Text = "RESET PALET" Or txtItemCard.Text = "RESET PALET" Then
            myCancel()
            lblStatus.Text = "ALREADY SCAN PALET"
            tmrWarning2.Enabled = False
            Return True
        ElseIf txtItemCard.Text = "RESET KANBAN" Then
            txtItemCard.Text = ""
            txtItemCard.Focus()
            txtItemCard.Text = ""
            lblStatus.Text = "ALREADY SCAN KANBAN"
            tmrWarning2.Enabled = False
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub tmrWarning_Tick(sender As System.Object, e As System.EventArgs) Handles tmrWarning.Tick
        lblStatus.Text = "Ready To Scan!"
        tmrWarning.Enabled = False
    End Sub

    Private Sub tmrWarning2_Tick(sender As System.Object, e As System.EventArgs) Handles tmrWarning2.Tick
        lblStatus.Text = "Ready To Scan!"
        tmrWarning2.Enabled = False
    End Sub


    Private Sub tmrJam_Tick(sender As System.Object, e As System.EventArgs) Handles tmrJam.Tick
        timeUtc = Date.UtcNow
        cstZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time")
        cstTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone)

        Dim currentTime As Integer = CInt(Format(cstTime, "HHmm"))

        If currentTime > 730 AndAlso currentTime < 2100 Then
            Shift.Text = "1"
            shiftA = Format(cstTime, "yyyy-MM-dd") & " 07:30"
            shiftB = Format(cstTime, "yyyy-MM-dd") & " 21:00"
            effdate = Format(cstTime, "yyyy-MM-dd")
            effdateSafety = Format(cstTime, "yyyyMMdd")
        Else
            Shift.Text = "3"
            If cstTime.Hour >= 21 Then
                shiftA = Format(cstTime, "yyyy-MM-dd") & " 21:00"
                shiftB = Format(cstTime.AddDays(1), "yyyy-MM-dd") & " 07:30"
                effdate = Format(cstTime, "yyyy-MM-dd")
                effdateSafety = Format(cstTime, "yyyyMMdd")
            Else
                shiftA = Format(cstTime.AddDays(-1), "yyyy-MM-dd") & " 21:00"
                shiftB = Format(cstTime, "yyyy-MM-dd") & " 07:30"
                effdate = Format(cstTime.AddDays(-1), "yyyy-MM-dd")
                effdateSafety = Format(cstTime.AddDays(-1), "yyyyMMdd")
            End If
        End If

        lblJam.Text = Format(cstTime, "HH:mm")
        blnAwal = cstTime.ToString("yyyy-MM-01")
        blnAkhir = cstTime.AddMonths(1).ToString("yyyy-MM-01")

        tglProd = If(currentTime > 730 AndAlso currentTime < 2400,
                     cstTime.ToString("yyyy-MM-dd"),
                     cstTime.AddDays(-1).ToString("yyyy-MM-dd"))

        If cstTime.Day >= 5 AndAlso ((cstTime.Second = 0 AndAlso (cstTime.Minute = 16 OrElse cstTime.Minute = 30 OrElse cstTime.Minute = 45)) OrElse trigger = 1) Then
            If Not BackgroundWorker2.IsBusy Then
                BackgroundWorker2.RunWorkerAsync()
                trigger += 1
            End If
        End If

        If cstTime.ToString("HHmm") = "2100" OrElse cstTime.ToString("HHmm") = "0730" Then
            Me.Dispose()
            Me.Close()
            frm_MAIN.Close()
            frm_Login.Show()
        End If
    End Sub

    Private Sub txtPalet_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs)
        'If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
        '    txtPalet.Text = txtPalet.Text.ToUpper
        '    If txtPalet.Text.Contains("YIM") Then
        '        paletize = txtPalet.Text
        '        txtCycle.Focus()
        '    ElseIf txtCycle.Text = "PRINT DS" Or txtPalet.Text = "PRINT DS" Then
        '        myCancelDS()   
        '        txtItemCard.Text = ""
        '        txtKanban.Text = ""
        '        txtItemCard.Focus()
        '    Else
        '        My.Computer.Audio.Play(My.Resources.salah, AudioPlayMode.Background)
        '        lblStatus.Text = "Palet Tidak Sesuai"
        '        myCancel()
        '    End If
        'End If
        '[)>069KY0BPBKAF583V0000J0V76183L5207K00199Q00008010KY0B761810405030Z
    End Sub

    Private Sub txtCycle_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCycle.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            If txtCycle.Text = "CYCLE 1" Or txtCycle.Text = "CYCLE 2" Or txtCycle.Text = "CYCLE 3" Or txtCycle.Text = "CYCLE 4" Or txtCycle.Text = "CYCLE 1 WJ" Or txtCycle.Text = "CYCLE 2 WJ" Then
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtItemCard.Focus()
                cycle = txtCycle.Text
            ElseIf txtCycle.Text = "PRINT DS" Then
                'myCancelDS()
                insertDS()
                'txtItemCard.Text = ""
                'txtKanban.Text = ""
                'txtItemCard.Focus()
                txtKanban.Text = ""
                txtKanban.Focus()
            ElseIf txtCycle.Text = "SHIPPING" Then
                myCancelDS()
                myCancel()
            ElseIf txtCycle.Text = "POSTPONE" Then
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtItemCard.Focus()
            ElseIf txtCycle.Text = "START SAFETY" Then
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtKanban.Focus()
            ElseIf txtCycle.Text = "RESET" Then
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtItemCard.Focus()
            Else
                My.Computer.Audio.Play(My.Resources.cycle_salah, AudioPlayMode.Background)
                lblStatus.Text = "Kanban Cycle Salah"
                txtCycle.Text = ""
                txtCycle.Focus()
            End If
        End If
    End Sub

    Private Sub txtItemCard_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCard.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            pn_his = txtItemCard.Text
            If pn_his.Length > 40 And txtCycle.Text <> "" And txtCycle.Text <> "PRINT DS" And txtCycle.Text <> "POSTPONE" And txtCycle.Text <> "SAFETY" Then
                cekData()
            ElseIf pn_his.Length > 40 And txtCycle.Text = "PRINT DS" Then
                'cekDataDS()
                insertDS()
                txtKanban.Text = ""
                txtKanban.Focus()
            ElseIf pn_his.Length > 40 And (txtCycle.Text = "POSTPONE") Then
                cekPostpone()
                'ElseIf pn_his.Length >= 10 And (txtCycle.Text = "SAFETY") Then
                '    cekSafety()
            ElseIf pn_his.Length >= 10 And (txtCycle.Text = "RESET") Then
                Dim myTem As New queryDB
                myTem.updateKanban(txtItemCard.Text)
                My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                txtItemCard.Text = ""
            ElseIf pn_his.Contains("CYCLE") And txtCycle.Text = "PRINT DS" Then
                printSummary(pn_his)
                Dim myTem As New queryDB
                myTem.updateSummary(pn_his)
                txtItemCard.Text = ""
            ElseIf pn_his = "SHIPPING" Then
                tmrSwitch.Enabled = True
                myCancelDS()
                myCancel()
            Else
                My.Computer.Audio.Play(My.Resources.salah, AudioPlayMode.Background)
                lblStatus.Text = "Data Tidak Sesuai"
                myCancel()
                tmrWarning2.Enabled = True
            End If
        End If
    End Sub

    Private Sub txtKanban_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtKanban.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            txtKanban.Text = txtKanban.Text.ToUpper
            If (txtItemCard.Text = "PRINT DS" Or txtCycle.Text = "PRINT DS") And txtKanban.Text <> "FINISH DS" Then
                insertDSDet()
            ElseIf txtKanban.Text = "FINISH DS" Then
                createDS()
            ElseIf txtCycle.Text = "START SAFETY" Then
                If txtKanban.Text = "FINISH SAFETY" Then
                    'safetyProcessMaster()
                    My.Computer.Audio.Play(My.Resources.lunas, AudioPlayMode.Background)
                    lblStatus.Text = "Berhasil"
                    tmrWarning2.Enabled = True
                    myCancel()
                    dgOrder.Visible = True

                    txtCycle.Text = ""
                    txtKanban.Text = ""
                    txtItemCard.Text = ""
                    txtCycle.Focus()

                    frm_summary_safety.Show()
                    frm_summary_safety.BringToFront()
                Else
                    'insertSafety()
                    safetyProcessDetail(txtKanban.Text)
                End If
            ElseIf txtItemCard.Text <> "" And txtCycle.Text <> "" Then
                lot = txtKanban.Text
                cekData()
            Else
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                lblStatus.Text = "Scan Item Card!"
                myCancel()
                tmrWarning2.Enabled = True
            End If
        End If
    End Sub

    Public Sub printSafetyCase(ByVal label As String)
        Dim cryRpt1 As New LabelCase()

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = label
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@caseID")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = lblLogin.Text
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@pic")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        cryRpt1.SetDatabaseLogon("usrdotnet", "r3m04tdw", "dbserver2.akebono-astra.co.id", "DELIVERY")
        Dim prd As New System.Drawing.Printing.PrintDocument

        'cryRpt1.PrintOptions.PrinterName = "AAVH"
        cryRpt1.PrintOptions.PrinterName = printer_casemark

        cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
        cryRpt1.PrintToPrinter(2, True, 0, 0)
    End Sub

    Public Sub safetyProcessMaster()
        Dim myTemp, myTemp2 As New queryDB
        Dim myDT, myDT2 As New DataTable

        Dim lastlot As String = myTemp.getLot(tglProd.Year, tglProd.Month, tglProd.Day)
        Dim labelID = effdateSafety + "/" + Shift.Text + "/" + lastlot.ToString.PadLeft(4, "0"c)

        myTemp.insertSafetyMaster(labelID, effdate, GetIPv4Address)
        qrcode(labelID)
        printSafetyCase(labelID)
    End Sub

    Private Sub safetyProcessDetail(ByVal tag As String)
        Dim myTemp, myTemp2 As New queryDB
        Dim myDT, myDT2 As New DataTable

        lot = tag.Substring(tag.IndexOf("(1)") + 3, tag.IndexOf("(2)", tag.IndexOf("(1)") + 1) - tag.IndexOf("(1)") - 3)

        myDT = myTemp.cekTagSudahOut(lot)
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
            lblStatus.Text = "Tag sudah dipulling!"
            txtKanban.Text = ""
            txtKanban.Focus()
            tmrWarning2.Enabled = True
        Else
            myDT2 = myTemp2.cekTagSudahSafety(lot)
            If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                lblStatus.Text = "Tag sudah discan!"
                txtKanban.Text = ""
                txtKanban.Focus()
                tmrWarning2.Enabled = True
            Else
                myTemp.insertSafetyDetail(lot, GetIPv4Address)
                My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                lblStatus.Text = "Berhasil"
                tmrWarning2.Enabled = True
                dgLot.Visible = True
                dgOrder.Visible = False
                vwSafety()

                txtKanban.Text = ""
                txtKanban.Focus()
            End If
        End If
    End Sub

    Public Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = strHostName.ToString
                'GetIPv4Address = "AAIJ-ITE-055"
            End If
        Next
    End Function

    Private Sub myCancelDS()
        lblLogin.Text = frm_Login.login_name
        'dgOrder.Refresh()
        txtItemCard.Text = ""
        txtQtyScan.Text = ""
        txtQtyOrder.Text = ""
        txtQtyOutsd.Text = ""

        SJ = ""
        itemCard = ""
        trans_id = ""
        po = ""
        snp = ""
        snp_cetak = ""
        count_scan = 0
        sum_status = False
        order_qty = 0
        num_to = 0
        vwOrderDS(itemCard)
        jmlScan = 0
        jmlScanTotal = 0
        txtNoScan.Text = ""
        pn_his = ""
        label = ""
        cycle = ""
        userYim = ""
        poLine = ""
        shipto = ""
        custpart = ""
        txtItemCard.Focus()
        dgLot.Visible = False

    End Sub

    Private Sub insertDS()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        myTem.insertDS("-", "", txtItemCard.Text)
        idprint = myTem.getNoDs

        userTrace = ""
    End Sub

    Private Sub insertDSDet()
        Dim myTem As New queryDB
        Dim myDT As New DataTable

        itemCard = txtKanban.Text
        If txtKanban.Text <> "FINISH DS" Then
            'vwOrderDS(itemCard)
        End If
        If userTrace = "" Then
            userTrace = itemCard.Substring(32, 4)
        End If

        Dim DS_Next As String = myTem.createDS3(userTrace)

        If DS_Next <> "Data tidak ditemukan" Then
            Dim numericPart As String = DS_Next.Substring(9, 5)
            Dim newNumericPart As Integer = Convert.ToInt32(numericPart) + 1

            Dim newDS_No As String = DS_Next.Substring(0, 9) & newNumericPart.ToString("D5")
            Console.WriteLine("DS No Baru: " & newDS_No)
            myQX.absId = newDS_No
        Else
            Console.WriteLine("Data tidak ditemukan")
        End If

        If itemCard.Substring(32, 4) <> userTrace Then
            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
            txtKanban.Text = ""
            txtKanban.Focus()
            Exit Sub
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
                txtQtyOrder.Text = myDT.Rows(0).Item("ym_order_qty").ToString
                pn_cust = myDT.Rows(0).Item("ms_part_cust").ToString
                Dim statLoading As String = myDT.Rows(0).Item("ym_stat_loading").ToString()
                Dim dateLoading As String = myDT.Rows(0).Item("ym_date_loading").ToString()

                ' Kondisi untuk hanya menampilkan pesan jika statLoading = "1" atau "True" dan dateLoading ada
                If (statLoading = "1" OrElse statLoading.ToLower() = "true") AndAlso Not String.IsNullOrEmpty(dateLoading) Then
                    My.Computer.Audio.Play(My.Resources.sj_sudah_dibuat, AudioPlayMode.Background)
                    lblStatus.Text = "Data DS Sudah Pernah Dibuat"
                    tmrWarning2.Enabled = True
                    myCancelDS()
                    txtKanban.Text = ""
                    txtKanban.Focus()
                    Exit Sub
                End If

                'qty_del = count_scan

                'If pn_cust.Substring(pn_cust.Length - 2, 1) = "J0" Then
                If pn_cust.Contains("-J0") Then
                    item_j0 = myDT.Rows(0).Item("item_jo").ToString
                    If item_j0 = "" Then
                        MsgBox("Item belum di mapping! Harap infokan ke bagian IT")
                        myCancelDS()
                        txtKanban.Text = ""
                        txtKanban.Focus()
                        Exit Sub
                    End If
                    pn_aaij = item_j0
                    cycle = "EXPORT"
                ElseIf pn_cust.Contains("-X0") Then
                    pn_aaij = pn_aaij & "-X"
                    cycle = "EXPORT"
                End If
                pn_desc = myDT.Rows(0).Item("ms_part_desc").ToString & "(" & pn_aaij & ")"

                kres = "#" & itemCard.Substring(37, 5)
                userYim = itemCard.Substring(32, 4)
                'userYim = "PRPC"
                poAAIJ = Format(cstTime, "MMyy") & userYim
                poLine = ""
                'Dim query As String = "SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & poAAIJ & "' AND sod_part='" & pn_aaij & "' " &
                '                            " AND sod_domain='AAIJ' WITH(nolock)"
                'Try
                '    elementXml = myTem.getwsa(query)
                '    elemList = elementXml.GetElementsByTagName("f_val")
                '    If elemList.Count > 0 Then
                '        For z As Integer = 0 To elemList.Count - 1
                '            rs = elementXml.GetElementsByTagName("f_val")(z).InnerText.Replace("||", "#")
                '            poLine = rs.Split("#"c)(0)
                '        Next
                '    End If
                'Catch ex As Exception
                '    MsgBox("WSA problem " & ex.ToString)
                'End Try

                'If poLine = "" Then
                '    MsgBox("Line tidak ada di QAD! Harap infokan ke bagian Sales")
                '    myCancelDS()
                '    txtKanban.Text = ""
                '    txtKanban.Focus()
                '    Exit Sub
                'End If

                'myDT = myTem.getLinePOSQL(poAAIJ, pn_aaij)
                'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                '    poLine = myDT.Rows(0).Item("sod_line").ToString
                'Else
                '    My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                '    MsgBox("Data PO tidak ada!")
                '    myCancelDS()
                '    Exit Sub
                'End If

                shipto = "2" & userYim
                txtQtyOutsd.Text = outstd_qty
                txtQtyScan.Text = count_scan
                SummaryScanDS("1")
                myDT = myTem.getTotScan(trans_id)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    qty_del = Decimal.Parse(myDT.Rows(0).Item("jml"))
                    If count_scan <> order_qty Then
                        My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
                        frm_dialog.ShowDialog()
                        If frm_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
                            myCancelDS()
                            txtKanban.Text = ""
                            txtKanban.Focus()
                            Exit Sub
                        End If
                    End If

                    If myTem.insertDSDet(trans_id, idprint, kres, poAAIJ, pn_desc, pn_cust, qty_del, txtKanban.Text, cycle, pn_aaij) = True Then
                        lblStatus.Text = "Transaksi berhasil !"
                        My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                        txtKanban.Text = ""
                        txtKanban.Focus()
                        'hitCreateDS += 1
                        'If hitCreateDS = 4 Then
                        '    createDS()
                        'End If
                        vwCreateDS(idprint)
                        Exit Sub
                    Else
                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                        MsgBox("Gagal insert DS!")
                        txtKanban.Text = ""
                        txtKanban.Focus()
                        Exit Sub
                    End If
                Else
                    My.Computer.Audio.Play(My.Resources.DSsudahbarcode, AudioPlayMode.Background)
                    lblStatus.Text = "Part belum di shipping!"
                    tmrWarning2.Enabled = True
                    myCancelDS()
                    txtKanban.Text = ""
                    txtKanban.Focus()
                End If
            End If
        Else
            My.Computer.Audio.Play(My.Resources.item_card_tidak_ditemukan, AudioPlayMode.Background)
            lblStatus.Text = "Data Item Card Tidak Ditemukan"
            tmrWarning2.Enabled = True
            'vwOrderDS("")
            myCancelDS()
        End If
    End Sub

    Private Sub createDS()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        Dim myTem2 As New queryDB
        Dim myDT2 As New DataTable
        'userYim = "PRPC"
        poAAIJ = Format(cstTime, "MMyy") & userYim
        'myQX.absId = "Test"
        shipto = userYim
        If userTrace = "" Then
            userTrace = "6111"
        End If
        Dim DS_Next As String = myTem.createDS3(userTrace)

        If DS_Next <> "Data tidak ditemukan" Then
            Dim numericPart As String = DS_Next.Substring(9, 5)
            Dim newNumericPart As Integer = Convert.ToInt32(numericPart) + 1

            Dim newDS_No As String = DS_Next.Substring(0, 9) & newNumericPart.ToString("D5")
            Console.WriteLine("cek no ds: " & newDS_No)
            myQX.absId = newDS_No

        Else
            Console.WriteLine("Data tidak ditemukan")
        End If

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
            qrcode(myQX.absId)
            myTem2.updateDS(idprint, myQX.absId, "PENDING", "", "", "")
            myTem2.updateDSDet(idprint, myQX.absId)
            'myTem2.updateShipDate("s" & myQX.absId, pn_aaij, poAAIJ)
            'myTem2.updatewsa("", myQX.absId, Format(cstTime, "yyyy-MM-dd"))
            'myTem2.insertDSCustom(myQX.absId, shipto)
            printBon(myQX.absId)
            lblStatus.Text = "Transaksi berhasil !"
            My.Computer.Audio.Play(My.Resources.surat_jalan_oke, AudioPlayMode.Background)
            myCancel()
            'myCancelDS()
        End If

    End Sub

    'FOR DS
    Private Sub vwCreateDS(itemCard)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        Try
            dgOrder.Rows.Clear()
            dgOrder.Columns.Clear()
            myDT = myTemp.getCreateDS(itemCard)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                dgOrder.Columns.Add("NO", "No. ")
                dgOrder.Columns.Add("po_no", "PO NO")
                dgOrder.Columns.Add("kres", "KRES")
                dgOrder.Columns.Add("part_no", "PART NO")
                dgOrder.Columns.Add("part_desc", "PART DESC")
                dgOrder.Columns.Add("qty_order", "QTY") '6
                dgOrder.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgOrder.Columns(0).Width = 30
                dgOrder.Columns(1).Width = 40
                dgOrder.Columns(2).Width = 60
                dgOrder.Columns(3).Width = 100
                dgOrder.Columns(4).Width = 120
                dgOrder.Columns(5).Width = 50
                dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Dim nomer As Integer = 1
                dgOrder.Rows.Add(myDT.Rows.Count)
                For i As Integer = 0 To myDT.Rows.Count - 1
                    dgOrder.Item(0, i).Value = nomer
                    dgOrder.Item(1, i).Value = myDT.Rows(i).Item("po").ToString
                    dgOrder.Item(2, i).Value = myDT.Rows(i).Item("kres").ToString
                    dgOrder.Item(3, i).Value = myDT.Rows(i).Item("model").ToString
                    dgOrder.Item(4, i).Value = myDT.Rows(i).Item("part_desc").ToString
                    dgOrder.Item(5, i).Value = myDT.Rows(i).Item("qty").ToString
                    nomer += 1
                Next
            End If
        Catch ex As Exception

        End Try

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

        crParameterDiscreteValue.Value = sh_ds
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

    ''CETAK DS VERSI LAMA. SATU DS SATU ITEM
    'Private Sub cekDataDS()
    '    Dim myTem As New queryDB
    '    Dim myDT As New DataTable
    '    tmrSwitch.Enabled = False
    '    itemCard = txtItemCard.Text
    '    vwOrderDS(itemCard)
    '    myDT = myTem.cekDataDS(itemCard)
    '    If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
    '        If itemCard <> "" Then
    '            trans_id = myDT.Rows(0).Item("id").ToString
    '            pn_aaij = myDT.Rows(0).Item("ms_item_nbr").ToString
    '            po = myDT.Rows(0).Item("ym_po").ToString
    '            snp = myDT.Rows(0).Item("ms_part_snp").ToString
    '            sum_status = myDT.Rows(0).Item("ym_stat_scan")
    '            order_qty = myDT.Rows(0).Item("ym_order_qty").ToString
    '            count_scan = myDT.Rows(0).Item("ym_qty_scan").ToString
    '            outstd_qty = myDT.Rows(0).Item("ym_order_qty") - myDT.Rows(0).Item("ym_qty_scan")
    '            cycle = myDT.Rows(0).Item("ym_cycle").ToString
    '            txtQtyOrder.Text = myDT.Rows(0).Item("ym_order_qty").ToString
    '            pn_cust = myDT.Rows(0).Item("ms_part_cust").ToString
    '            smallpart = myDT.Rows(0).Item("ms_smallpart").ToString
    '            If smallpart = "True" Or smallpart = True Then

    '            End If
    '            'qty_del = count_scan

    '            'If pn_cust.Substring(pn_cust.Length - 2, 1) = "J0" Then
    '            If pn_cust.Contains("-J0") Then
    '                pn_aaij = pn_aaij & "-P"
    '                cycle = "EXPORT"
    '            ElseIf pn_cust.Contains("-X0") Then
    '                pn_aaij = pn_aaij & "-X"
    '                cycle = "EXPORT"
    '            End If
    '            pn_desc = myDT.Rows(0).Item("ms_part_desc").ToString & "(" & pn_aaij & ")"

    '            kres = "#" & itemCard.Substring(37, 5)
    '            userYim = itemCard.Substring(32, 4)
    '            poAAIJ = Format(cstTime, "MMyy") & userYim
    '            myDT = myTem.getLinePO(poAAIJ, pn_aaij)
    '            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
    '                poLine = myDT.Rows(0).Item("sod_line").ToString
    '            Else
    '                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
    '                MsgBox("Data PO tidak ada!")
    '                myCancelDS()
    '                Exit Sub
    '            End If
    '            shipto = "2" & userYim
    '            txtQtyOutsd.Text = outstd_qty
    '            txtQtyScan.Text = count_scan
    '            SummaryScanDS("1")
    '            myDT = myTem.getTotScan(trans_id)
    '            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
    '                qty_del = Decimal.Parse(myDT.Rows(0).Item("jml"))
    '                If qty_del <> order_qty Then
    '                    My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '                    frm_dialog.ShowDialog()
    '                    If frm_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
    '                        myCancelDS()
    '                        Exit Sub
    '                    End If
    '                End If
    '                If myQX.requestInbound(pn_aaij, qty_del, poAAIJ, kres, poLine, shipto) = True Then
    '                    If myTem.insertDS(trans_id, myQX.absId, kres, poAAIJ, pn_desc, pn_cust, qty_del, myQX.note_xml, txtItemCard.Text, cycle) = True Then
    '                        If myTem.updateLoadingDet(trans_id) = False Then
    '                            myTem.updateNGDS(myQX.absId)
    '                            My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
    '                            lblStatus.Text = "Gagal update data!"
    '                            myCancelDS()
    '                            Exit Sub
    '                        Else
    '                            qrcode(myQX.absId)
    '                            If count_scan = order_qty Then
    '                                myTem.updateLoading(trans_id)
    '                            End If
    '                            myTem.updateShipDate("s" & myQX.absId, pn_aaij, poAAIJ)
    '                            myTem.insertDSCustom(myQX.absId, shipto)
    '                            printBon()
    '                            lblStatus.Text = "Transaksi berhasil !"
    '                            My.Computer.Audio.Play(My.Resources.surat_jalan_oke, AudioPlayMode.Background)
    '                            vwOrderDS(itemCard)
    '                        End If
    '                        myCancelDS()
    '                    Else
    '                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
    '                        MsgBox("Gagal insert DS!")
    '                        myCancelDS()
    '                        Exit Sub
    '                    End If
    '                Else
    '                    My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
    '                    MsgBox("Gagal membuat DS!")
    '                    myCancelDS()
    '                    Exit Sub
    '                End If
    '            Else
    '                My.Computer.Audio.Play(My.Resources.DSsudahbarcode, AudioPlayMode.Background)
    '                lblStatus.Text = "Part belum di shipping!"
    '                tmrWarning2.Enabled = True
    '                myCancelDS()
    '            End If
    '            tmrSwitch.Enabled = True
    '        End If
    '    Else
    '        My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
    '        lblStatus.Text = "Data Tidak Sesuai"
    '        tmrWarning2.Enabled = True
    '        vwOrderDS("")
    '        myCancelDS()
    '    End If
    'End Sub

    Private Sub cekPostpone()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        tmrSwitch.Enabled = False
        itemCard = txtItemCard.Text
        vwOrderDS(itemCard)
        myDT = myTem.cekDataDS(itemCard)
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            If itemCard <> "" Then
                trans_id = myDT.Rows(0).Item("id").ToString
                myTem.updatePostponeMs(trans_id, txtCycle.Text)
                myTem.updatePostponeTr(trans_id, txtCycle.Text)
                My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                lblStatus.Text = "Berhasil"
                tmrWarning2.Enabled = True
                myCancel()
            End If
        Else
            My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
            lblStatus.Text = "Data Tidak Sesuai"
            tmrWarning2.Enabled = True
            vwOrderDS("")
            myCancelDS()
        End If
    End Sub

    'Private Sub cekSafety()
    '    Dim myTem As New queryDB
    '    Dim myDT As New DataTable
    '    Dim total, itemPn As String
    '    'tmrSwitch.Enabled = False

    '    kanbanSafety = txtItemCard.Text
    '    'vwOrderDS(itemCard)
    '    myDT = myTem.cekKanbanSafety(kanbanSafety)
    '    If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
    '        If kanbanSafety <> "" Then
    '            item_no = myDT.Rows(0).Item("item_no").Replace("-J0", "").Replace("-80", "").Replace("-X0", "").ToString
    '            itemPn = myDT.Rows(0).Item("item_no").ToString
    '            kanban_qty = myDT.Rows(0).Item("kanban_qty").ToString
    '            full_item = myDT.Rows(0).Item("item_no").ToString
    '            snp = myTem.cekSnp(itemPn).ToString

    '            total = myTem.getTotalTag(kanbanSafety)
    '            If total < (kanban_qty / snp) Then
    '                txtKanban.Text = ""
    '                txtKanban.Focus()
    '            Else
    '                My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '                lblStatus.Text = "Data Tidak Sesuai"
    '                tmrWarning2.Enabled = True
    '                txtItemCard.Text = ""
    '                txtItemCard.Focus()
    '            End If

    '            'txtQtyOrder.Text = Integer.Parse(kanban_qty / 10).ToString
    '            'txtQtyScan.Text = total
    '            'txtQtyOutsd.Text = (Integer.Parse(txtQtyOrder.Text) - Integer.Parse(txtQtyScan.Text)).ToString


    '            'myTem.updatePostponeMs(trans_id, txtCycle.Text)
    '            'My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
    '            'lblStatus.Text = "Berhasil"
    '            'tmrWarning2.Enabled = True
    '            'myCancel()
    '        End If
    '    Else
    '        My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '        lblStatus.Text = "Data Tidak Sesuai"
    '        tmrWarning2.Enabled = True
    '        'vwOrderDS("")
    '        'myCancelDS()
    '    End If
    'End Sub

    'Private Sub insertSafety()
    '    Dim myTem As New queryDB
    '    Dim myDT As New DataTable
    '    'Dim total As String
    '    Dim cek As String
    '    Dim pn_tag As String
    '    'tmrSwitch.Enabled = False
    '    tagProduksi = txtKanban.Text
    '    'vwOrderDS(itemCard)
    '    If tagProduksi <> "" Then

    '        'pn_kanban = myTem.cekPnKanban(txtItemCard.Text).ToString
    '        lot = tagProduksi.Substring(tagProduksi.IndexOf("(1)") + 3, tagProduksi.IndexOf("(2)", tagProduksi.IndexOf("(1)") + 1) - tagProduksi.IndexOf("(1)") - 3)
    '        pn_tag = tagProduksi.Substring(tagProduksi.IndexOf("(2)") + 3, tagProduksi.IndexOf("(3)", tagProduksi.IndexOf("(2)") + 1) - tagProduksi.IndexOf("(2)") - 3)
    '        pnAaijTag = tagProduksi.Substring(tagProduksi.IndexOf("(2)") + 3, tagProduksi.IndexOf("(3)", tagProduksi.IndexOf("(2)") + 1) - tagProduksi.IndexOf("(2)") - 3)
    '        pn_tag = myTem.cekPnTag(pn_tag).ToString

    '        'If pn_kanban = pn_tag Then
    '        cek = myTem.cekTag(lot).ToString

    '        If cek <> "" Then
    '            My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '            lblStatus.Text = "Data Tidak Sesuai"
    '            tmrWarning2.Enabled = True
    '            txtKanban.Text = ""
    '            txtKanban.Focus()
    '        Else
    '            'total = myTem.getTotalTag(kanbanSafety)

    '            'If total < (kanban_qty / snp) Then
    '            myTem.insertTrSafety(lot, tagProduksi)
    '            myTem.insertLog(lot, pn_tag, tagProduksi, "SFT", "IN", pnAaijTag)
    '            myTem.updateStockSafetyTambah(pnAaijTag)
    '            myTem.updateLogStatusSafety(lot)
    '            'myTem.updateStockKurang(item_no)

    '            'Try
    '            '    Dim webClient As New System.Net.WebClient
    '            '    Dim result As String = webClient.DownloadString("https://dashboard.akebono-astra.co.id/store/test.php?part=" & pnAaijTag & "&snp=" & snp & "&status=OUT")
    '            'Catch ex As Exception

    '            'End Try

    '            'total = myTem.getTotalTag(kanbanSafety)
    '            'If total = (kanban_qty / snp) Then
    '            '    My.Computer.Audio.Play(My.Resources.lunas, AudioPlayMode.Background)
    '            '    lblStatus.Text = "Berhasil"
    '            '    tmrWarning2.Enabled = True

    '            '    'txtCycle.Text = ""
    '            '    txtKanban.Text = ""
    '            '    txtItemCard.Text = ""
    '            '    'txtCycle.Focus()
    '            '    txtItemCard.Focus()
    '            'Else
    '            My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
    '            lblStatus.Text = "Berhasil"
    '            tmrWarning2.Enabled = True

    '            txtKanban.Text = ""
    '            txtKanban.Focus()
    '            'End If
    '            'End If
    '        End If
    '        'Else
    '        '    My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '        '    lblStatus.Text = "Data Tidak Sesuai"
    '        '    tmrWarning2.Enabled = True
    '        '    txtKanban.Text = ""
    '        '    txtKanban.Focus()
    '        'End If
    '    Else
    '        My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
    '        lblStatus.Text = "Data Tidak Sesuai"
    '        tmrWarning2.Enabled = True
    '        'vwOrderDS("")
    '        'myCancelDS()
    '    End If
    'End Sub

    Private Sub SummaryScanDS(status)
        If status = "2" Then
            txtQtyOutsd.Text = order_qty - count_scan - snp
            txtQtyScan.Text = count_scan + snp
            jmlScan = (count_scan + snp) / snp
            jmlScanTotal = Round(order_qty / snp, 0)
            txtNoScan.Text = jmlScan & " / " & jmlScanTotal
        Else
            txtQtyOutsd.Text = order_qty - count_scan
            txtQtyScan.Text = count_scan
            jmlScan = count_scan / snp
            jmlScanTotal = Round(order_qty / snp, 0)
            txtNoScan.Text = jmlScan & " / " & jmlScanTotal
        End If
    End Sub

    Private Sub printBon()
        Dim cryRpt1 As New cetak_DS()

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue


        crParameterDiscreteValue.Value = myQX.absId
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

        crParameterDiscreteValue.Value = pic_ds
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
        'cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
        cryRpt1.PrintToPrinter(1, True, 0, 0)
    End Sub

    Private Sub printSummary(cyc)
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

        crParameterDiscreteValue.Value = Format(cstTime, "yyyy-MM-dd")
        crParameterFieldDefinitions = cryRpt1.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@tgl")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        cryRpt1.SetDatabaseLogon("usrdotnet", "r3m04tdw", "dbserver2.akebono-astra.co.id", "DELIVERY")
        Dim prd As New System.Drawing.Printing.PrintDocument

        cryRpt1.PrintOptions.PrinterName = "HP LaserJet M101-M106"

        cryRpt1.PrintOptions.PaperSize = prd.PrinterSettings.DefaultPageSettings.PaperSize.RawKind
        cryRpt1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
        cryRpt1.PrintToPrinter(1, True, 0, 0)
    End Sub

    'Private Sub cekSJ(sj, pn_cust, qty_del)
    '    Dim myTemp As New queryDB
    '    Dim myDT As New DataTable
    '    myDT = myTemp.getPartSJ(sj, pn_cust, qty_del)
    '    If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
    '        saveTransaksi()
    '    ElseIf myDT.Rows.Count = 0 Then
    '        My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
    '        lblStatus.Text = "Surat Jalan salah!"
    '    Else
    '        My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
    '        lblStatus.Text = "Surat Jalan salah!"
    '    End If
    'End Sub

    Private Function cekOrderQad(number)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable
        Dim tr_nbr As String
        Try

            Dim query As String = "SELECT tr_nbr FROM PUB.tr_hist WHERE tr_domain ='AAIJ' AND tr_site='AAIJ' AND tr_userid='MFG' AND tr_type='ISS-TR' and tr_NBR= '" & number & "' ORDER BY tr_nbr asc WITH(nolock)"
            Try
                elementXml = myTemp.getwsa(query)
                elemList = elementXml.GetElementsByTagName("f_val")
                If elemList.Count > 0 Then
                    For z As Integer = 0 To elemList.Count - 1
                        rs = elementXml.GetElementsByTagName("f_val")(z).InnerText.Replace("||", "#")
                        tr_nbr = rs.Split("#"c)(0)
                        If tr_nbr <> "" Then
                            Return True
                        ElseIf tr_nbr = "" Then
                            Return False
                        Else
                            Return True
                        End If
                        Return True
                    Next
                End If
            Catch ex As Exception
                MsgBox("WSA problem " & ex.ToString)
            End Try

            'myDT = myTemp.getTransOrder(number)
            'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            '    Return True
            'ElseIf myDT.Rows.Count = 0 Then
            '    Return False
            'Else
            '    Return True
            'End If
        Catch ex As Exception
            Return True
        End Try
    End Function

    Private Sub cekLine(ByVal poAAIJ, ByVal pn_aaij)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

        Dim query As String = "SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & poAAIJ & "' AND sod_part='" & pn_aaij & "' " &
                                            " AND sod_domain='AAIJ' WITH(nolock)"
        Try
            elementXml = myTemp.getwsa(query)
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

        'myDT = myTemp.getLinePO(poAAIJ, pn_aaij)
        'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
        '    poLine = myDT.Rows(0).Item("sod_line").ToString
        'Else
        '    MsgBox("Data PO tidak ada!")
        '    Exit Sub
        'End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            If set_conform = "true" Then
                'If (cstTime.Day = 5 And cstTime.Hour = 15) Then
                Dim myTem As New queryDB
                Dim myDT As New DataTable
                Dim myTem2 As New queryDB
                Dim myDT2 As New DataTable
                myDT = myTem.getDStoConfirm(blnAwal, blnAkhir)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        Dim idorder As String = myDT.Rows(i).Item("idorder")
                        Dim dsno As String = myDT.Rows(i).Item("ds_no")
                        Dim tglConfirm As Date = myDT.Rows(i).Item("tgl")
                        'Dim model As String = myDT.Rows(i).Item("model")
                        Dim po_conf As String = myDT.Rows(i).Item("po")
                        'myDT2 = myTem2.cekItemCust(po_conf, model)
                        'If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                        If myQX2.requestDS(dsno, Format("", Format(tglConfirm, "yyyy-M-dd")).ToString, Format("", Format(cstTime, "yyyy-M-dd")).ToString) = True Then
                            myTem2.updateConfirmDS(idorder, myQX2.note_xml, myQX2.hasil)
                        End If
                        'End If
                    Next
                End If
                'End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            If set_conform = "true" Then
                'If (cstTime.Day >= 5 And cstTime.Minute < 5) Or (cstTime.Day > 5 And cstTime.Minute <= 2) Then
                Dim myTem As New queryDB
                Dim myDT As New DataTable
                Dim myTem2 As New queryDB
                Dim myDT2 As New DataTable
                myDT = myTem.getDStoConfirm(blnAwal, blnAkhir)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        Dim idorder As String = myDT.Rows(i).Item("idprint").ToString
                        Dim dsno As String = myDT.Rows(i).Item("ds_no").ToString
                        Dim tglConfirm As Date = myDT.Rows(i).Item("tgl")
                        'Dim model As String = myDT.Rows(i).Item("model")
                        'Dim po_conf As String = myDT.Rows(i).Item("po")
                        'myDT2 = myTem2.cekItemCust(po_conf, model)
                        'If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                        If myQX2.requestDS(dsno, Format("", Format(tglConfirm, "yyyy-MM-dd")).ToString, Format("", Format(cstTime, "yyyy-MM-dd")).ToString) = True Then
                            myTem2.updateConfirmDS(idorder, myQX2.note_xml, myQX2.hasil)
                        End If
                        'Else

                        'End If
                    Next
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub tmrSwitch_Tick(sender As System.Object, e As System.EventArgs) Handles tmrSwitch.Tick
        myCancelDS()
        myCancel()
    End Sub

    Private Sub tmrJeda_Tick(sender As System.Object, e As System.EventArgs) Handles tmrJeda.Tick
        tmrJeda.Enabled = False
    End Sub

    Private Sub btnDSOES_Click(sender As System.Object, e As System.EventArgs) Handles btnDSOES.Click
        frm_cetak.Show()
        frm_cetak.BringToFront()
        'Dim frmalert As New frm_cetak()
        'If frmalert.ShowDialog(Me) = System.Windows.Forms.DialogResult.Cancel Then
        '    myCancel()
        'End If
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub achievementShift()
        Dim myTemp As New queryDB
        Dim myDT As New DataTable
        Try
            Dim aw As List(Of String) = New List(Of String)
            Dim ak As List(Of String) = New List(Of String)
            Dim valTipe As New Dictionary(Of String, String)
            valTipe.Clear()
            myDT = myTemp.getMsPart(dtp_delivery)
            Try
                If myDT IsNot Nothing AndAlso myDT.Rows.Count > 0 Then
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        Dim part_no As String = myDT.Rows(i).Item("ym_item_nbr")
                        Dim ms_tipe As String = myDT.Rows(i).Item("ms_part_desc1")
                        If (ms_tipe.Length < 2) Then
                            ms_tipe = part_no
                        End If
                        ak.Add(part_no)
                        valTipe.Add(part_no, ms_tipe)
                    Next
                End If
            Catch ex As Exception
                Exit Sub
            End Try

            Dim type_ar As New Dictionary(Of String, String)
            Dim jml_ar As New Dictionary(Of String, String)

            Dim frmList As List(Of String) = New List(Of String)

            Chart3.Series.Clear()
            Chart3.Legends.Clear()
            Chart3.Legends.Add("legend1").Alignment = StringAlignment.Center
            Chart3.Legends("legend1").Docking = Docking.Bottom
            ''Chart3.Series.Add("total")
            ''Chart3.Series.Add("target")
            frmList.Add("Lunas")
            frmList.Add("On Proses")
            frmList = frmList.Distinct().ToList()
            For Each formu In frmList
                Chart3.Series.Add(formu)
                Chart3.Series(formu).ChartType = SeriesChartType.StackedColumn
                Chart3.Legends.Add(formu)
            Next

            Dim valJumArray As List(Of String) = New List(Of String)
            Dim valFormula As New Dictionary(Of String, String)
            Dim valJam As New Dictionary(Of String, String)
            Dim valJum As New Dictionary(Of String, String)
            valFormula.Clear()
            valJam.Clear()
            valJum.Clear()

            Dim hit As Integer = 0
            For Each hsl In ak
                valFormula.Clear()
                valJum.Clear()
                Dim total As Decimal = 0
                myDT = myTemp.getValueShift(dtp_delivery, hsl)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        valJum.Add("Lunas", myDT.Rows(i).Item("qty_scan"))
                        valJum.Add("On Proses", myDT.Rows(i).Item("outs"))
                        valFormula.Add("Lunas", "Lunas")
                        valFormula.Add("On Proses", "On Proses")
                        total += myDT.Rows(i).Item("qty_scan") + myDT.Rows(i).Item("outs")
                    Next
                End If
                'Dim color As System.Drawing.Color = color.Red
                For Each frm In frmList
                    Dim jml As String = 0
                    If valFormula.ContainsKey(frm) Then
                        jml = valJum(frm)
                        If jml = "" Or jml = 0 Then
                            jml = 0
                        Else
                            Chart3.Series(frm).Label = ""
                        End If
                        Chart3.Series(frm).Points.AddXY(valTipe(hsl), jml)
                        Chart3.Series(frm).IsValueShownAsLabel = True
                        'Chart3.Series(frm).Font = New Font("Verdana", "6")
                        Chart3.Series(frm).ChartType = SeriesChartType.StackedColumn

                        If frm = "Lunas" Then
                            Chart3.Series(frm).Color = Color.GreenYellow
                        Else
                            Chart3.Series(frm).Color = Color.Red
                        End If
                    Else
                        'Chart3.Series(frm).Points.AddXY(valTipe(hsl), jml)
                        'Chart3.Series(frm).IsValueShownAsLabel = False
                        'Chart3.Series(frm).Font = New Font("Verdana", "6")
                        'Chart3.Series(frm).ChartType = SeriesChartType.StackedColumn
                        'Chart3.Series(frm).Color = color
                    End If

                    Chart3.Series(frm).Font = New Font("Verdana", 6, FontStyle.Bold)
                    Chart3.ChartAreas(0).AxisY.Interval = 500
                Next
                Chart3.ChartAreas(0).AxisX.LabelStyle.Font = New Font("Verdana", "6")
                Chart3.ChartAreas(0).AxisX.Interval = 1
                Chart3.ChartAreas(0).AxisX.MajorGrid.Enabled = False
                Chart3.ChartAreas(0).AxisY.MajorGrid.Enabled = True
                Chart3.ChartAreas(0).AxisX.LabelStyle.Angle = -45
                'Chart3.ChartAreas(0).AxisY.LabelStyle.Angle = -90
                ''Chart3.Series(0).Points.AddXY(valTipe(hsl), total)
                Chart3.Series(0).IsValueShownAsLabel = True
                Chart3.Series(0).Font = New Font("Verdana", "6")

                hit += 1
            Next
            hit += 1
            Chart3.Series("target").Points.AddXY("Target", 6)
            Chart3.Series("target").IsValueShownAsLabel = False
            Chart3.Series("target").Color = Color.Transparent

            For Each s As Series In Chart3.Series
                For Each dp As DataPoint In s.Points
                    If dp.YValues(0) = 0 Then
                        dp.Label = ""
                    End If
                Next
            Next
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub dtp_del_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dtp_del.ValueChanged
        dtp_delivery = Format(dtp_del.Value, "yyyy-MM-dd")
        achievementShift()
        myCancel()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim frmalert As New frm_loading()
        If frmalert.ShowDialog(Me) = System.Windows.Forms.DialogResult.Cancel Then
            myCancel()
        End If
    End Sub

    Public Sub qrcode(barcode)
        Dim objQRcode As QRCodeEncoder = New QRCodeEncoder
        Dim imgImage As Image
        Dim objbitmap As Bitmap

        objQRcode.QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE
        objQRcode.QRCodeScale = 2
        objQRcode.QRCodeVersion = 5
        'objQRcode.QRCodeErrorCorrect = ThoughtWorks.QRCode.Codec.QRCodeEncoder()
        imgImage = objQRcode.Encode(barcode)
        objbitmap = New Bitmap(imgImage)

        barcode = barcode.ToString.Replace("/", "")
        objbitmap.Save("C:\AAIJ\Image\" & barcode & ".jpg")
        'PictureBox1.ImageLocation = "C:\AAIJ\" & lot & ".jpg"
    End Sub

    Private Sub BackgroundWorker3_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        Dim pn, snptag As String

        myDT = myTem.getOrderOut()
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            pn = myDT.Rows(0).Item("tag").ToString
            snptag = myTem.cekSnp(pn)
            Try
                Dim webClient As New System.Net.WebClient
                Dim result As String = webClient.DownloadString("https://dashboard.akebono-astra.co.id/store/test.php?part=" & pn & "&snp=" & snptag & "&status=OUT")
            Catch ex As Exception

            End Try
        Else
        End If
    End Sub

    Private Sub tmrUpdt_Tick(sender As Object, e As EventArgs) Handles tmrUpdt.Tick
        If BackgroundWorker3.IsBusy = False Then
            BackgroundWorker3.RunWorkerAsync()
        End If
    End Sub

    Private Sub btnReprintSafety_Click(sender As Object, e As EventArgs) Handles btnReprintSafety.Click
        frm_cetakCasemark.Show()
        frm_cetakCasemark.BringToFront()
    End Sub

    Private Sub txtCycle_TextChanged(sender As Object, e As EventArgs) Handles txtCycle.TextChanged

    End Sub
End Class
