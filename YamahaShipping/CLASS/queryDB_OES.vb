Imports System.Data.SqlClient
Imports System.Xml

Public Class queryDB_OES
    Dim timeUtc As Date = Date.UtcNow
    Dim cstZone As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time")
    Dim cstTime As Date = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone)
    Public Function getOrder(ByVal kanban As String, ByVal itemCard As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim po As String = ""
            Dim pn_customer As String = ""
            Dim user As String = ""

            If itemCard <> "" Then
                pn_customer = itemCard.Substring(11, 14)
                po = itemCard.Substring(37, 5)
                If IsNumeric(po) = True Then
                    po = Convert.ToInt32(po)
                End If
                user = itemCard.Substring(32, 4)
            End If
            'kanban = kanban.Replace("-P", "")
            'kanban = kanban.Replace("-", "")
            Dim query As String = ""

            'executing the command and assigning it to connection
            If kanban = "" And itemCard = "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on a.ym_item_nbr=replace(ms_part_cust,'-','') where ym_stat_scan = 'False' and  YEAR(ym_del_duedate)='" & frm_barcode.cstTime.Year & "' and ym_user_code ='7618' and ym_del_duedate > '2019-07-25' order by id asc"
            ElseIf kanban = "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on a.ym_item_nbr=replace(ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "'  and ym_user_code ='" & user & "' and ym_user_code ='7618' and ym_del_duedate > '2019-07-25' order by id asc"
            ElseIf kanban <> "" And itemCard <> "" Then
                'query = "select  * from yamaha_order a inner join yamahaP3_ms_part b on a.ym_item_nbr=replace(substring(part,2,18),'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and part='" & kanban & "'  and ym_user_code ='" & user & "' and YEAR(ym_del_duedate)='" & frm_barcode.cstTime.Year & "' and ym_user_code ='7618' and ym_del_duedate > '2019-07-25' order by id asc"
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on a.ym_item_nbr=replace(ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "'  and ym_user_code ='" & user & "' and ym_user_code ='7618' and ym_del_duedate > '2019-07-25' order by id asc"
            End If

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLOTTr(ByVal idTr As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim query As String = ""
            query = "select  ym_po, tr_lot, tr_lot_time from yamaha_tr_hist a inner join yamaha_order b on a.tr_id=b.id where tr_id = '" & idTr & "' order by a.id desc"

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function


    Public Function cekData(ByVal itemCard As String, ByVal label As String) As DataTable
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim po As String = ""
            Dim pn_customer As String = ""
            Dim user As String = ""

            If itemCard <> "" Then
                pn_customer = itemCard.Substring(11, 14)
                po = itemCard.Substring(37, 5)
                If IsNumeric(po) = True Then
                    po = Convert.ToInt32(po)
                End If
                user = itemCard.Substring(32, 4)
            End If
            'kanban = kanban.Replace("-P", "")
            'kanban = kanban.Replace("-", "")
            Dim query As String = ""

            'executing the command and assigning it to connection
            If itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_user_code ='" & user & "' and ms_part_cust not like '%J0' order by id asc"
            End If
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertTransaksi(ByVal transid As String, ByVal sidd_part As String, ByVal snp As Decimal, ByVal cycle As String, ByVal lot As String, lotDB As String, ByVal fulltag As String, ByVal palet As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_tr_hist (tr_id, tr_part, tr_snp, tr_creadate, tr_creaby, tr_shift, tr_cycle, tr_pc, tr_lot,  tr_db, tr_fulltag, tr_palet) values (" & transid & ", '" & sidd_part & "', " & snp & ", getdate(),  '" & frm_barcode.lblLogin.Text & "'," & frm_barcode.Shift.Text & ",'" & cycle & "','" & frm_barcode.lblPc & "','" & lot & "','" & lotDB & "','" & fulltag & "','" & palet & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateCountScan(ByVal transid As String, ByVal sidd_part As String, ByVal snp As Decimal, ByVal status As String, ByVal po As String, ByVal cycle As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            If status = False Then
                cmd = "update yamaha_order set ym_qty_scan=ym_qty_scan+ " & snp & " where id='" & transid & "'"
            Else
                cmd = "update yamaha_order set ym_stat_scan='" & status & "', ym_cycle='" & cycle & "' where id='" & transid & "' "
            End If
            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateStock(ByVal part As String, ByVal snp As Decimal) As Boolean
        Try
            connectDB.connectSQL()
            Dim sqlQuery As String = String.Format("update yamahaP3_ms_part set part_stock=part_stock - '" & snp & "' where part='Y" & part & "'")

            connectDB.executeDBSQL(sqlQuery)
            Return True
        Catch ex As Exception
            'MsgBox("SOMETHING PROBLEM")
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateMinScan(ByVal transid As String, ByVal snp As Decimal) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_order set ym_qty_scan=ym_qty_scan- " & snp & " where id='" & transid & "'"
            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertUser(ByVal npk As String, ByVal nama As String, ByVal password As String, ByVal admin As String, ByVal aktif As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = "insert into aavh_login(login_npk,login_nama,login_password,login_admin,login_aktif,creadate,creaby)values('" & npk & "','" & nama & "', '" & password & "','" & admin & "','" & aktif & "', getdate() ,'" & frm_barcode.lblLogin.Text & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLabel(ByVal tag As String, ByVal label As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""
            query = "select  * from yamaha_ms_part where ms_item_nbr='" & tag & "'"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getStatCase(ByVal id As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select  * from yamaha_order where ym_stat_scan='False' and id='" & id & "'"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function cekTagProd(ByVal lot As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select tr_lot from yamaha_tr_hist where tr_fulltag='" & lot & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function


    'FOR DS
    Public Function getOrderDS(ByVal SJ As String, ByVal itemCard As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim po As String = ""
            Dim pn_customer As String = ""
            Dim user As String = ""

            If itemCard <> "" Then
                pn_customer = itemCard.Substring(11, 14)
                po = itemCard.Substring(37, 5)
                If IsNumeric(po) = True Then
                    po = Convert.ToInt32(po)
                End If
                user = itemCard.Substring(32, 4)
            End If
            Dim query As String = ""

            'executing the command and assigning it to connection
            If SJ = "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_part_cust like '%80' and ym_stat_loading='False' and ym_user_code ='" & user & "' and year(ym_del_duedate)='" & frm_barcode.cstTime.Year & "' order by id asc"
            ElseIf SJ <> "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_item_nbr='" & SJ & "' and ms_part_cust like '%80' and ym_stat_loading='False' and ym_user_code ='" & user & "' and year(ym_del_duedate)='" & frm_barcode.cstTime.Year & "' order by id asc"
            Else
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ms_part_cust like '%80' and ym_stat_loading='False' and ym_del_duedate > '2018-05-01' order by id asc"
            End If

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function cekDataDS(ByVal itemCard As String) As DataTable
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim po As String = ""
            Dim pn_customer As String = ""
            Dim user As String = ""

            If itemCard <> "" Then
                pn_customer = itemCard.Substring(11, 14)
                po = itemCard.Substring(37, 5)
                If IsNumeric(po) = True Then
                    po = Convert.ToInt32(po)
                End If
                user = itemCard.Substring(32, 4)
            End If
            Dim query As String = ""

            'executing the command and assigning it to connection
            'query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_stat_loading='False' and ym_user_code ='" & user & "' order by id asc"
            'query = "select  top 8 * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_loading='False' and ym_user_code ='7618' and ym_del_duedate > '2019-04-01' AND ym_item_nbr not like '%J0%' AND ym_qty_scan <>'0' order by id asc"
            'query = "select tr_id, sum(tr_snp) as jml, tr_part, ym_po, ym_item_desc,ym_order_qty, ym_stat_scan, ym_po,ym_qty_scan " &
            '        " from yamaha_tr_hist a inner join yamaha_order b on a.tr_id=b.id " &
            '        " where tr_loading_st='False' and tr_creadate > '2019-04-01' " &
            '        " AND tr_db='OES' AND tr_snp IS NOT NULL group by tr_id, tr_part, " &
            '        " ym_po, ym_item_desc, ym_order_qty, ym_stat_scan, ym_po, ym_qty_scan"
            query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_stat_loading='False' and ym_user_code ='" & user & "' order by id asc"

            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function createDS(ByVal id As String) As DataTable
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            'executing the command and assigning it to connection
            query = "select SUM(qty) as qty, model, pn_aaij FROM yamaha_print_dsdet_oes where ds_no = '" & id & "' group by model, pn_aaij"


            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function createDS2(ByVal id As String) As DataTable
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            'executing the command and assigning it to connection
            query = "select * FROM yamaha_print_dsdet_oes where ds_no = '" & id & "'"


            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    'Public Function getLinePO(ByVal po As String, ByVal part As String) As DataTable
    '    Dim item_nbr As String = ""
    '    Try
    '        connectDB.connectQAD()
    '        Dim myDA As New Odbc.OdbcDataAdapter
    '        Dim myDT As New DataTable
    '        myDA = New Odbc.OdbcDataAdapter("SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & po & "' AND sod_part='" & part & "' " &
    '                                        " AND sod_domain='AAIJ' WITH(nolock)", connectDB.myConnectionQAD)
    '        myDA.Fill(myDT)
    '        Return myDT
    '    Catch ex As Exception
    '        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    Public Function getwsa(ByVal query As String) As XmlDocument
        Try
            query = query.Replace("#", "aiueo")
            Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri("https://wsa.akebono-astra.co.id/wsa_aaij.php?wsa_func=wsa_exec&query=" & query & "&conn=Q21_Dev|VB"))
            request.Method = System.Net.WebRequestMethods.Http.Get
            Dim response As System.Net.HttpWebResponse = request.GetResponse()
            ' process response here
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())

            Dim datastream As String = sr.ReadToEnd
            datastream = datastream.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "'").Replace(vbCrLf, "").Replace(vbLf, "")

            Dim xmlRequest As String = datastream.Split("^"c)(0)
            Dim xmlRespon As String = datastream.Split("^"c)(1)

            Dim elementXml = New XmlDocument()
            elementXml.LoadXml(xmlRespon)

            Return elementXml
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function getTotScan(ByVal transid As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select sum(tr_snp) as jml from yamaha_tr_hist where tr_id='" & transid & "' and tr_loading_st='False' AND tr_snp IS NOT NULL "
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getTotScan2(ByVal transid As String) As String
        Dim id As String = 0
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select sum(tr_snp) as jml from yamaha_tr_hist where tr_id='" & transid & "' and tr_loading_st='False' AND tr_snp IS NOT NULL ", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
            id = 0
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function insertDS(ByVal ds_no As String, ByVal noteXml As String, ByVal dstag As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_print_ds_OES (ds_no,  noteXml, ds_tag, shift) values ('" & ds_no & "',  '" & noteXml & "','" & dstag & "', '" & frm_barcode.Shift.Text & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertDSDet(ByVal idorder As Integer, ByVal ds_no As String, ByVal kres As String, ByVal po As String, ByVal part_desc As String, ByVal model As String, ByVal qty As Decimal, ByVal dstag As String, ByVal cycle As String, ByVal pnaaij As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_print_dsdet_OES (idorder, ds_no, kres, po, part_desc, model, qty, ds_tag, cycle, pn_aaij) values ('" & idorder & "','" & ds_no & "', '" & kres & "', '" & po & "','" & part_desc & "',  '" & model & "','" & qty & "','" & dstag & "', '" & cycle & "', '" & pnaaij & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateNGDS(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds set stat_ds='NG' WHERE ds_no='" & transid & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLoadingDet(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_tr_hist set tr_loading_st='True', tr_loading_date='" & frm_barcode.cstTime & "' where tr_id='" & transid & "' and tr_loading_st='False'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLoading(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_order set ym_stat_loading='True', ym_date_loading='" & frm_barcode.cstTime & "' where id='" & transid & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateConfirmDS(ByVal transid As String, ByVal notexml As String, ByVal statDS As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds set stat_ds='OK',noteXmlDS='" & notexml & "', creadate_ds='" & frm_barcode.cstTime & "' WHERE idorder='" & transid & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateSummary(ByVal cycle As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds set stat_print_sum='True', print_date_sum='" & frm_barcode.cstTime & "' where stat_print_sum='False' and cycle='" & cycle & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getDS() As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select  top 300 * from yamaha_print_ds_oes order by idprint desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getSummary() As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select top 100 convert(varchar(10),tgl,120) as tgl, cycle from yamaha_print_ds group by tgl, cycle order by tgl desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function



    Public Function getNoDs() As String
        Dim id As String = 0
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select top 1 idprint from yamaha_print_ds_oes order by idprint desc", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
            id = 0
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function updateDS(ByVal transid As String, ByVal dsno As String, ByVal statDs As String, ByVal noteDS As String, ByVal noteError As String, ByVal xmlSend As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            If statDs <> "-" Then
                cmd = "update yamaha_print_ds_OES set ds_no='" & dsno & "',stat_ds='" & statDs & "',noteXml='" & noteDS & "',creadate_ds=GETDATE(), noteError='" & noteError & "', xmlSend='" & xmlSend & "' where idprint='" & transid & "'"
            Else
                cmd = "update yamaha_print_ds_OES set noteXml='" & noteDS & "',creadate_ds=GETDATE(), noteError='" & noteError & "', xmlSend='" & xmlSend & "' where idprint='" & transid & "'"
            End If

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateDSDet(ByVal transid As String, ByVal dsno As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_print_dsdet_OES set ds_no='" & dsno & "' where ds_no='" & transid & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertLoading(ByVal truckid As String, ByVal trid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "insert into yamaha_load_trms (id_truck, id_trans) values ('" & truckid & "','" & trid & "')"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLoadId() As String
        Dim id As String = 0
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select substring(id_trans,7,2) as id from yamaha_load_trms where year(creadate)='" & cstTime.Year & "' and month(creadate)='" & cstTime.Month & "' and day(creadate)='" & cstTime.Day & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
            id = 0
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function getValidPalet(ByVal paletid As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select tr_palet from yamaha_tr_hist where tr_palet='" & paletid & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            'MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function insertLoadingDet(ByVal paletid As String, ByVal trid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "insert into yamaha_load_trdet (idms, palet_id) values ('" & trid & "','" & paletid & "')"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

End Class
