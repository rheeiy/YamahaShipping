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
Imports System.Xml

Public Class frm_barcodev2
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
    Dim myQX As New maintainSalesOrderShipperSatuan()
    Dim myQX2 As New confirmDS()
    Public qty_del As Decimal = 0
    Public blnAwal As String = ""
    Public blnAkhir As String = ""
    Public pic_ds As String = ""
    'WSA
    Dim elemList As XmlNodeList = Nothing
    Dim elementXml = New XmlDocument()
    Dim rs As String = ""
    'END WSA
    Private Sub vwOrder(kanban, itemCard)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

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
            dgOrder.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgOrder.Columns(0).Width = 40
            dgOrder.Columns(1).Width = 80
            dgOrder.Columns(2).Width = 90
            dgOrder.Columns(3).Width = 150
            dgOrder.Columns(4).Width = 180
            dgOrder.Columns(5).Width = 70
            dgOrder.Columns(6).Width = 70
            dgOrder.Columns(7).Width = 70
            dgOrder.Columns(8).Width = 70
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
            dgLot.Columns(1).Width = 80
            dgLot.Columns(2).Width = 90
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

    Private Sub frm_barcode_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        vwOrder(kanban, itemCard)
        myCancel()
        lblStatus.Text = "Ready To Scan!"
        'BackgroundWorker1.RunWorkerAsync()
        BackgroundWorker2.RunWorkerAsync()
        lblPC.Text = GetIPv4Address()
        pic_ds = GetSetting("pic")
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
        dgOrder.Refresh()
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
    End Sub


    Private Sub cekData()
        If reset() = False Then
            Dim myTem As New queryDB
            Dim myDT As New DataTable
            itemCard = txtItemCard.Text
            label = txtCycle.Text
            vwOrder(kanban, itemCard)
            myDT = myTem.cekData(itemCard, label)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                If kanban = "" And itemCard <> "" And label <> "" Then
                    trans_id = myDT.Rows(0).Item("id").ToString
                    pn_aaij = myDT.Rows(0).Item("ms_item_nbr").ToString
                    kanban = pn_aaij
                    po = myDT.Rows(0).Item("ym_po").ToString
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

                    If sum_status = False Then
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
                                txtKanban.Text = ""
                                txtKanban.Focus()
                                vwOrder(kanban, itemCard)
                            Else
                                'My.Computer.Audio.Play("konfirmasi.wav")
                                snp = outstd_qty
                                txtKanban.Text = ""
                                txtKanban.Focus()
                                vwOrder(kanban, itemCard)
                                'Exit Sub
                            End If
                        End If
                    Else
                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                        lblStatus.Text = "Part sudah selesai dibarcode!"
                        myCancel()
                        tmrWarning2.Enabled = True
                    End If
                ElseIf kanban <> "" And itemCard <> "" And label <> "" Then
                    If myTem.cekTagProd(txtKanban.Text, lotDB).ToString <> "" And (txtKanban.Text.Length = 9 Or txtKanban.Text.Length = 11) Then
                        My.Computer.Audio.Play("sudah-di-scan.wav")
                        myCancel()
                        Exit Sub
                    End If
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
                        If myTem.cekTagProd(lot, lotDB).ToString <> "" Then
                            My.Computer.Audio.Play("sudah-di-scan.wav")
                            myCancel()
                            Exit Sub
                        End If
                        If pnAaijTag = pn_aaij Then
                            saveTransaksi()
                        Else
                            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                            lblStatus.Text = "Tag Salah!"
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
                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                        lblStatus.Text = "Part sudah selesai dibarcode!"
                        myCancel()
                        tmrWarning2.Enabled = True
                    End If
                End If
            Else
                My.Computer.Audio.Play(My.Resources.salah, AudioPlayMode.Background)
                lblStatus.Text = "Data Tidak Sesuai"
                tmrWarning2.Enabled = True
                If txtItemCard.Text = "DANDORI" Or txtKanban.Text = "DANDORI" Or txtCycle.Text = "DANDORI" Then
                    myCancel()
                ElseIf txtCycle.Text <> "" And txtItemCard.Text <> "" And txtKanban.Text = "" Then
                    txtItemCard.Text = ""
                    txtItemCard.Focus()
                    vwOrder("", itemCard)
                ElseIf txtItemCard.Text <> "" And txtKanban.Text <> "" And txtCycle.Text <> "" Then
                    txtKanban.Text = ""
                    txtItemCard.Text = ""
                    txtItemCard.Focus()
                    vwOrder("", itemCard)
                Else
                    myCancel()
                End If
            End If
        End If
    End Sub

    Private Sub saveTransaksi()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        'If tmrJeda.Enabled = False Then
        If myTem.updateCountScan(trans_id, itemCard, snp, "False", po, cycle) = True Then
            If myTem.insertTransaksi(trans_id, kanban, snp, cycle, lot, lotTime, lotDB, txtKanban.Text, txtPalet.Text, txtItemCard.Text) = True Then
                'If myTem.insertTransaksi(trans_id, kanban, snp, cycle, lot, lotTime, lotDB, txtKanban.Text) = True Then
                count_scan += snp
                lblStatus.Text = "Transaksi berhasil !"
                tmrWarning2.Enabled = True
                SummaryScan("2")
                If count_scan = order_qty Then
                    My.Computer.Audio.Play(My.Resources.lunas, AudioPlayMode.Background)
                    myTem.updateCountScan(trans_id, itemCard, snp, "True", po, cycle)
                    myDT = myTem.getStatCase(trans_id)
                    If myDT.Rows.Count = 0 Then
                        'printBon()
                    End If
                    vwOrder(kanban, itemCard)
                    'myTem.matiinLampu(pn_cust2)
                    myCancel()
                Else
                    My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                    txtItemCard.Text = ""
                    txtItemCard.Focus()
                    txtKanban.Text = ""
                    lot = ""
                    kanban = ""
                    dgLot.Visible = True
                    vwLot(trans_id)
                End If
                'tmrJeda.Enabled = True
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
        ''Else
        'My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
        'lblStatus.Text = "Gagal, coba scan kembali"
        'txtKanban.Text = ""
        'End If
    End Sub

    Private Sub SummaryScan(status)
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
        Jam.Text = Format(cstTime, "dd/M/yyyy HH:mm:ss tt")
        If (Format(cstTime, "HHmm")) > 730 And (Format(cstTime, "HHmm")) < 2100 Then
            Shift.Text = 1
            shiftA = Format(cstTime, "yyyy-M-dd") & " 07:30"
            shiftB = Format(cstTime, "yyyy-M-dd") & " 21:00"
        Else
            Shift.Text = 3
            If cstTime.Hour >= 21 Then
                shiftA = Format(cstTime, "yyyy-M-dd") & " 21:00"
                shiftB = Format(cstTime.AddDays("+1"), "yyyy-M-dd") & " 07:30"
            Else
                shiftA = Format(cstTime.AddDays("-1"), "yyyy-M-dd") & " 21:00"
                shiftB = Format(cstTime, "yyyy-M-dd") & " 07:30"
            End If
        End If

        lblJam.Text = (Format(cstTime, "HH:mm"))
        blnAwal = cstTime.ToString("yyyy-M-01 07:30")
        blnAkhir = cstTime.AddMonths(1).ToString("yyyy-M-01 07:30")

        If (cstTime.Day = 5 Or cstTime.Day = 1) And cstTime.Hour = 10 And cstTime.Minute = 0 And cstTime.Second = 0 Then
            BackgroundWorker1.RunWorkerAsync()
        End If
        If cstTime.Day > 5 And cstTime.Second = 0 And (cstTime.Minute = 16 Or cstTime.Minute = 30 Or cstTime.Minute = 45) Then
            If BackgroundWorker2.IsBusy = False Then
                BackgroundWorker2.RunWorkerAsync()
            End If
        End If

        If cstTime.ToString("HHmm") = "2100" Or cstTime.ToString("HHmm") = "0730" Then
            Me.Dispose()
            Me.Close()
            frm_MAIN.Close()
            frm_Login.Show()
        End If
    End Sub

    Private Sub txtCycle_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCycle.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            If txtCycle.Text = "CYCLE 1" Or txtCycle.Text = "CYCLE 2" Or txtCycle.Text = "CYCLE 3" Or txtCycle.Text = "CYCLE 4" Or txtCycle.Text = "CYCLE 1 WJ" Or txtCycle.Text = "CYCLE 2 WJ" Then
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtItemCard.Focus()
                cycle = txtCycle.Text
            ElseIf txtCycle.Text = "PRINT DS" Then
                myCancelDS()
                txtItemCard.Text = ""
                txtKanban.Text = ""
                txtItemCard.Focus()
            ElseIf txtCycle.Text = "SHIPPING" Then
                myCancelDS()
                myCancel()
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
            If pn_his.Length > 40 And txtCycle.Text <> "" And txtCycle.Text <> "PRINT DS" Then
                cekData()
            ElseIf pn_his.Length > 40 And txtCycle.Text = "PRINT DS" Then
                cekDataDS()
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
            If txtItemCard.Text <> "" And txtCycle.Text <> "" Then
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
    Private Function GetIPv4Address() As String
        GetIPv4Address = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = strHostName.ToString
            End If
        Next
    End Function

    ''FOR DS
    Private Sub vwOrderDS(itemCard)
        Dim myTemp As New queryDB
        Dim myDT As New DataTable

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
            dgOrder.Columns(0).Width = 40
            dgOrder.Columns(1).Width = 80
            dgOrder.Columns(2).Width = 90
            dgOrder.Columns(3).Width = 120
            dgOrder.Columns(4).Width = 230
            dgOrder.Columns(5).Width = 70
            dgOrder.Columns(6).Width = 40
            dgOrder.Columns(7).Width = 70
            dgOrder.Columns(8).Width = 70
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
    End Sub

    Private Sub myCancelDS()
        lblLogin.Text = frm_Login.login_name
        dgOrder.Refresh()
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


    Private Sub cekDataDS()
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        tmrSwitch.Enabled = False
        itemCard = txtItemCard.Text
        vwOrderDS(itemCard)
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
                userYim = itemCard.Substring(32, 4)
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
                '    myCancelDS()
                '    Exit Sub
                'End If
                shipto = "2" & userYim
                txtQtyOutsd.Text = outstd_qty
                txtQtyScan.Text = count_scan
                SummaryScanDS("1")
                myDT = myTem.getTotScan(trans_id)
                If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                    qty_del = myDT.Rows(0).Item("jml")
                    If qty_del <> order_qty Then
                        My.Computer.Audio.Play(My.Resources.no, AudioPlayMode.Background)
                        frm_dialog.ShowDialog()
                        If frm_dialog.DialogResult = Windows.Forms.DialogResult.Cancel Then
                            myCancelDS()
                            Exit Sub
                        End If
                    End If
                    If myQX.requestInbound(pn_aaij, qty_del, poAAIJ, kres, poLine, shipto) = True Then
                        If myTem.insertDS(trans_id, myQX.absId, kres, poAAIJ, pn_desc, pn_cust, qty_del, myQX.note_xml, txtItemCard.Text, cycle) = True Then
                            If myTem.updateLoadingDet(trans_id) = False Then
                                myTem.updateNGDS(myQX.absId)
                                My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
                                lblStatus.Text = "Gagal update data!"
                                myCancelDS()
                                Exit Sub
                            Else
                                If count_scan = order_qty Then
                                    myTem.updateLoading(trans_id)
                                End If
                                'myTem.updateShipDate("s" & myQX.absId, pn_aaij, poAAIJ)
                                myTem.updatewsa("", myQX.absId, Format(cstTime, "yyyy-MM-dd"))
                                'myTem.insertDSCustom(myQX.absId, shipto)
                                printBon()
                                lblStatus.Text = "Transaksi berhasil !"
                                My.Computer.Audio.Play(My.Resources.surat_jalan_oke, AudioPlayMode.Background)
                                vwOrderDS(itemCard)
                            End If
                            myCancelDS()
                        Else
                            My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                            MsgBox("Gagal insert DS!")
                            myCancelDS()
                            Exit Sub
                        End If
                    Else
                        My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                        MsgBox("Gagal membuat DS!")
                        myCancelDS()
                        Exit Sub
                    End If
                Else
                    My.Computer.Audio.Play(My.Resources.DSsudahbarcode, AudioPlayMode.Background)
                    lblStatus.Text = "Part belum di shipping!"
                    tmrWarning2.Enabled = True
                    myCancelDS()
                End If
                tmrSwitch.Enabled = True
            End If
        Else
            My.Computer.Audio.Play(My.Resources.surat_jalan_salah, AudioPlayMode.Background)
            lblStatus.Text = "Data Tidak Sesuai"
            tmrWarning2.Enabled = True
            vwOrderDS("")
            myCancelDS()
        End If
    End Sub

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

    Public Sub printBon()
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

        cryRpt1.SetDatabaseLogon("sa", "r3m04aaij", "HRD", "dotNET")
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

        cryRpt1.SetDatabaseLogon("sa", "r3m04aaij", "HRD", "dotNET")
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
                    Dim model As String = myDT.Rows(i).Item("model")
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
        Catch ex As Exception

        End Try

    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            'If (cstTime.Day >= 5 And cstTime.Minute < 5) Or (cstTime.Day > 5 And cstTime.Minute <= 2) Then
            Dim myTem As New queryDB
            Dim myDT As New DataTable
            Dim myTem2 As New queryDB
            Dim myDT2 As New DataTable
            myDT = myTem.getDStoConfirm(blnAwal, blnAkhir)
            If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                For i As Integer = 0 To myDT.Rows.Count - 1
                    Dim idorder As String = myDT.Rows(i).Item("idorder").ToString
                    Dim dsno As String = myDT.Rows(i).Item("ds_no").ToString
                    Dim tglConfirm As Date = myDT.Rows(i).Item("tgl")
                    'Dim model As String = myDT.Rows(i).Item("model")
                    Dim po_conf As String = myDT.Rows(i).Item("po")
                    'myDT2 = myTem2.cekItemCust(po_conf, model)
                    'If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                    If myQX2.requestDS(dsno, Format("", Format(tglConfirm, "yyyy-MM-dd")).ToString, Format("", Format(cstTime, "yyyy-MM-dd")).ToString) = True Then
                        myTem2.updateConfirmDS(idorder, myQX2.note_xml, myQX2.hasil)
                    End If
                    'Else

                    'End If
                Next
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
End Class
