Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Xml
Imports CrystalDecisions.CrystalReports.Engine
Imports MySql.Data.MySqlClient

Public Class queryDB
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
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_scan = 'False' and ms_part_cust like '%80' and YEAR(ym_del_duedate)='" & cstTime.Year & "' and month(ym_del_duedate)='" & cstTime.Month & "' order by id asc"
            ElseIf kanban = "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_part_cust like '%80' and ym_user_code ='" & user & "' and YEAR(ym_del_duedate)='" & cstTime.Year & "' and month(ym_del_duedate)='" & cstTime.Month & "' order by id asc"
            ElseIf kanban <> "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_item_nbr='" & kanban & "' and ms_part_cust like '%80' and ym_user_code ='" & user & "' and YEAR(ym_del_duedate)='" & cstTime.Year & "' and month(ym_del_duedate)='" & cstTime.Month & "' order by id asc"
            End If

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
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
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLotSafety() As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim query As String = ""
            query = "select lot, creadate from yamaha_pallet_dtl_sft where palet_code = '-' order by id"

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function


    Public Function cekData(ByVal itemCard As String, ByVal label As String) As DataTable
        Try
            connectDB.connectSQL()
            '[)>069KY0EP1FDF580U0100J0V76183L5201KM1267Q00000510K                Z

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim po As String = ""
            Dim pn_customer As String = ""
            Dim user As String = ""

            If itemCard <> "" Then
                pn_customer = itemCard.Substring(11, 14) '[)>069KY0EPB6HF580U0000J0V76183L5208KM0005Q00000510K                Z
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
            If itemCard <> "" And label <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ym_stat_scan = 'False' and replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_user_code ='" & user & "' order by id asc"
            End If
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function cekDataItemCard(ByVal itemCard As String) As DataTable
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

            query = "select  * from yamaha_order where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_user_code ='" & user & "'"

            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertTransaksi(ByVal transid As String, ByVal sidd_part As String, ByVal snp As Decimal, ByVal cycle As String, ByVal lot As String, ByVal lotTime As Date, lotDB As String, ByVal fulltag As String, ByVal tr_palet As String, ByVal tr_idcard As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_tr_hist (tr_id, tr_part, tr_snp, tr_creadate, tr_creaby, tr_shift, tr_cycle, tr_pc, tr_lot, tr_lot_time, tr_db, tr_fulltag, tr_palet,tr_idcard) values (" & transid & ", '" & sidd_part & "', " & snp & ", getdate(),  '" & frm_barcode.lblLogin.Text & "'," & frm_barcode.Shift.Text & ",'" & cycle & "','" & frm_barcode.lblPc & "','" & lot & "','" & lotTime & "','" & lotDB & "','" & fulltag & "','" & tr_palet & "','" & tr_idcard & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateInbound(ByVal transid As String, ByVal status As String, ByVal xmlRequest As String, ByVal xmlRespond As String, ByVal locFrom As String, ByVal locTo As String, ByVal effdate As Date, ByVal inbound_note As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_tr_hist set tr_inbound_st='" & status & "', tr_inbound_note='" & inbound_note & "', tr_effdate='" & effdate & "', tr_locfrom='" & locFrom & "', tr_locto='" & locTo & "', tr_xmlRequest='" & xmlRequest & "', tr_xmlRespond='" & xmlRespond & "',tr_inbounddate=getdate() where id='" & transid & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertLog(ByVal lot As String, ByVal item As String, ByVal fulltag As String, ByVal asal As String, ByVal status As String, ByVal tag As String)
        Try
            connectDB.connectSQL()
            Dim cmd As String
            If asal = "SFT" And status = "OUT" Then
                cmd = "INSERT INTO [dbo].[yamaha_tr_log_fg] ([lot],[tr_ms_prt],[creadate],[tr_status],[tr_fulltag],[line_asal],[tag],[cek]) VALUES('" + lot + "', '" + item + "',GETDATE(),'" + status + "','" + fulltag + "','" + asal + "','" + tag + "', 1)"
            Else
                cmd = "INSERT INTO [dbo].[yamaha_tr_log_fg] ([lot],[tr_ms_prt],[creadate],[tr_status],[tr_fulltag],[line_asal],[tag]) VALUES('" + lot + "', '" + item + "',GETDATE(),'" + status + "','" + fulltag + "','" + asal + "','" + tag + "')"
            End If

            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLog(ByVal item As String)
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "UPDATE yamaha_ms_part SET stock_fg = stock_fg - ms_part_snp WHERE ms_part_cust LIKE '%" + item + "%'"
            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
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
            ''MsgBox(ex.Message)
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
            ''MsgBox(ex.Message)
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
            'MsgBox(ex.Message)
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
            'MsgBox(ex.Message)
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
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLot(ByVal lot As String, ByVal pn_cus As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "SELECT * from assy_ms_tr_tag where lot ='" & lot & "'"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function cekTagProd(ByVal lot As String, ByVal db As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select tr_lot from yamaha_tr_hist where tr_lot='" & lot & "' and tr_db = '" + db + "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function
    Public Function cekTagProdLog(ByVal lot As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select lot from yamaha_tr_log_fg where lot ='" & lot & "' and tr_status = 'OUT'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
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
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_part_cust like '%80' and ym_stat_loading='False' and ym_user_code ='" & user & "' and year(ym_del_duedate)='" & cstTime.Year & "' order by id asc"
            ElseIf SJ <> "" And itemCard <> "" Then
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ms_item_nbr='" & SJ & "' and ms_part_cust like '%80' and ym_stat_loading='False' and ym_user_code ='" & user & "' and year(ym_del_duedate)='" & cstTime.Year & "' order by id asc"
            Else
                query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where ms_part_cust like '%80' and ym_stat_loading='False' and ym_del_duedate > '2018-05-01' order by id asc"
            End If

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
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
            'query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and (ym_stat_loading = 0 OR ym_date_loading IS NULL) and ym_user_code ='" & user & "' order by id asc"
            query = "select  * from yamaha_order a inner join yamaha_ms_part b on replace(a.ym_item_nbr,'-','')=replace(b.ms_part_cust,'-','') where replace(ym_item_nbr,'-','')='" & pn_customer & "' and ym_po='" & po & "' and ym_user_code ='" & user & "' order by id asc"

            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    'Public Function cekKanbanSafety(ByVal kanban As String) As DataTable
    '    Try
    '        connectDB.connectSQLKanban()

    '        Dim myDA As New SqlDataAdapter
    '        Dim myDT As New DataTable
    '        Dim query As String = ""

    '        'executing the command and assigning it to connection
    '        query = "select item_no,item_name,item_back_no,store_back_no,kanban_qty from kanban_print a " &
    '        " inner join kanban_tr_kanban_sft b on a.kanban_no=b.kanban_no" &
    '        " inner join kanban_tr_store c on b.kanban_back_no_a=c.no" &
    '        " inner join kanban_ms_item d on c.store_back_no=d.item_back_no" &
    '        " where a.print_no = '" & kanban & "'"

    '        myDA = New SqlDataAdapter(query, connectDB.sqlConnectionKanban)
    '        myDA.Fill(myDT)
    '        Return myDT
    '    Catch ex As Exception
    '        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectSQLKanban()
    '    End Try
    'End Function

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
    '        '        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    Public Function getLinePOSQL(ByVal po As String, ByVal part As String) As DataTable
        Dim item_nbr As String = ""
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & po & "' AND sod_part='" & part & "' " &
                                            " AND sod_domain='AAIJ' WITH(nolock)", connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            '        'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getLot(ByVal thn As String, ByVal bln As String, ByVal tgl As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select ISNULL(max(substring(palet_code,12,4)) + 1,0) as palet_code from yamaha_pallet_mstr_sft where year(effdate)='" & thn & "' AND month(effdate)='" & bln & "' and day(effdate)='" & tgl & "' AND len(palet_code)='15'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch exa As Exception

        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function getwsa(ByVal query As String) As XmlDocument
        Try
            query = query.Replace("#", "aiueo")
            Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri("https://wsa.akebono-astra.co.id/wsa_aaij.php?wsa_func=wsa_exec&query=" & query & "&conn=Q21_Live|VB"))
            'Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri("https://wsa.akebono-astra.co.id/wsa_aaij.php?wsa_func=wsa_exec&query=" & query & "&conn=Q21|VB"))
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

    Public Function updatewsa(ByVal query As String, ByVal nomerDs As String, ByVal tgl As String) As XmlDocument
        Try
            query = query.Replace("#", "aiueo")
            query = "date=" & tgl & "&ds=" & nomerDs
            Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri("https://wsa.akebono-astra.co.id/wsa_aaij.php?wsa_func=aaij_update_ds&" & query & "&conn=Q21_Live|VB"))
            'Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(New System.Uri("https://wsa.akebono-astra.co.id/wsa_aaij.php?wsa_func=aaij_update_ds&" & query & "&conn=Q21|VB"))
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

    Public Function getOrderOut() As DataTable
        Try
            connectDB.connectSQLBGP()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("SELECT TOP 1 tag FROM yamaha_tr_log_fg WHERE tr_status = 'OUT' and tag IS NOT NULL and creadate > '2021-09-11' and cek = 0 GROUP BY tag", connectDB.sqlConnectionBGP)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQLBGP()
        End Try
    End Function

    Public Function cekTagSudahOut(ByVal tag As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("SELECT TOP 1 tag FROM yamaha_tr_log_fg WHERE tr_status = 'OUT' and lot = '" + tag + "'", connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function cekTagSudahSafety(ByVal tag As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("SELECT TOP 1 lot FROM yamaha_pallet_dtl_sft WHERE lot = '" + tag + "'", connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getTotScan(ByVal transid As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select ISNULL(sum(tr_snp) ,0)as jml from yamaha_tr_hist where tr_id='" & transid & "' and tr_loading_st='False'"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertDS(ByVal idorder As Integer, ByVal ds_no As String, ByVal kres As String, ByVal po As String, ByVal part_desc As String, ByVal model As String, ByVal qty As Decimal, ByVal noteXml As String, ByVal dstag As String, ByVal cycle As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_print_ds (idorder, ds_no, kres, po, part_desc, model, qty, noteXml, ds_tag, cycle, shift) values ('" & idorder & "','" & ds_no & "', '" & kres & "', '" & po & "','" & part_desc & "',  '" & model & "','" & qty & "', '" & noteXml & "','" & dstag & "', '" & cycle & "','" & frm_barcode.Shift.Text & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateNGDS(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds_OES set stat_ds='NG' WHERE ds_no='" & transid & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLoadingDet(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_tr_hist set tr_loading_st='True', tr_loading_date='" & cstTime & "',tr_postpone='False',tr_safety='False' where tr_id='" & transid & "' and tr_loading_st='False'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLoading(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_order set ym_stat_loading='True', ym_date_loading='" & cstTime & "',ym_postpone='False',ym_safety='False' where id='" & transid & "' and ym_stat_scan = 'True'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    'Public Function updateShipDate(ByVal ds As String, ByVal item As String, ByVal order As String) As Boolean
    '    Try
    '        connectDB.connectQAD()
    '        Dim cmd As String = ""
    '        cmd = "UPDATE PUB.abs_mstr SET abs_shp_date ='" & Format(cstTime, "yyyy-M-dd") & "' WHERE abs_par_id='" & ds & "' AND abs_domain='AAIJ' AND abs_shipfrom='AAIJ' AND  abs_item='" & item & "' AND abs_order='" & order & "'"
    '        'UPDATE PUB.abs_mstr SET abs_shp_date = '2018-03-23' WHERE abs_par_id='s18/2590B/01567'
    '        connectDB.executeDBQAD(cmd)
    '        Return True
    '    Catch ex As Exception
    ''        'MsgBox(ex.Message)
    '        Return False
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    'Public Function insertDSCustom(ByVal ds As String, ByVal shipto As String) As Boolean
    '    Try
    '        connectDB.connectQADCustom()
    '        Dim cmd As String = ""
    '        cmd = "insert into PUB.xxds_data (xxds_doctype, xxds_no, xxds_shipto, xxds_status, xxds_domain, xxds_year) values ('Shipper', '" & ds & "','" & shipto & "','OPEN','AAIJ','" & cstTime.Year & "')"
    '        connectDB.executeDBQADCustom(cmd)
    '        Return True
    '    Catch ex As Exception
    ''        'MsgBox(ex.Message)
    '        Return False
    '    Finally
    '        connectDB.disconnectQADCustom()
    '    End Try
    'End Function

    'Public Function getPartSJ(ByVal sj As String, ByVal part As String, ByVal qty As Decimal) As DataTable
    '    Dim item_nbr As String = ""
    '    Try
    '        connectDB.connectQAD()
    '        Dim myDA As New Odbc.OdbcDataAdapter
    '        Dim myDT As New DataTable
    '        sj = sj.Replace("(1)", "")
    '        myDA = New Odbc.OdbcDataAdapter("SELECT sod_custpart  FROM PUB.abs_mstr a, PUB.sod_det b " &
    '        " WHERE a.abs_order = b.sod_nbr And a.abs_item = b.sod_part " &
    '        " AND a.abs_domain='AAIJ' AND a.abs_domain=b.sod_domain " &
    '        " AND substring(A.abs_par_id,2)='" & sj & "' and sod_custpart ='" & part & "' and abs_qty='" & qty & "' WITH(nolock)", connectDB.myConnectionQAD)
    '        myDA.Fill(myDT)
    '        Return myDT
    '    Catch ex As Exception
    ''        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    Public Function getDStoConfirm(ByVal blnAwal As String, ByVal blnAkhir As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            'query = "select  * from yamaha_print_ds where stat_ds is NULL and po not like '%PRPC' order by idprint asc"
            query = "select  * from yamaha_print_ds_OES where confirmStatus is NULL and creadate between '" & blnAwal & "' and '" & blnAkhir & "' and ds_no <> '-' order by idprint asc"
            'query = "select  * from yamaha_print_ds_OES where confirmStatus is NULL and creadate between '2022-04-01 00:00' and '2022-04-19 07:30' and ds_no <> '-' order by idprint desc"
            'query = "select  * from yamaha_print_ds where ds_no='18/2690D/00746' order by idprint desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    'Public Function cekItemCust(ByVal po As String, ByVal part As String) As DataTable
    '    Dim item_nbr As String = ""
    '    Try
    '        connectDB.connectQAD()
    '        Dim myDA As New Odbc.OdbcDataAdapter
    '        Dim myDT As New DataTable
    '        myDA = New Odbc.OdbcDataAdapter("SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & po & "' AND sod_custpart='" & part & "' " &
    '                                        " AND sod_domain='AAIJ' WITH(nolock)", connectDB.myConnectionQAD)
    '        myDA.Fill(myDT)
    '        Return myDT
    '    Catch ex As Exception
    ''        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    Public Function updateConfirmDS(ByVal transid As String, ByVal notexml As String, ByVal statDS As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds_OES set confirmStatus='True',confirmXml='" & notexml & "', confirmDate=getdate() WHERE idprint='" & transid & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateSummary(ByVal cycle As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_print_ds_OES set stat_print_sum='True', print_date_sum='" & cstTime & "' where stat_print_sum='False' and cycle='" & cycle & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateKanban(ByVal kanban As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update yamaha_tr_safety set status_pull = 1 where kanban ='" & kanban & "'"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
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

            query = "select  top 300 * from yamaha_print_ds_OES where line_name='OEM' order by idprint desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getCasemark() As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select top 300 * from yamaha_pallet_mstr_sft order by creadate desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getSafetybyPC(ByVal pc As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "select * from vw_yamaha_safety_summary WHERE pc = '" + pc + "'"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
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

            query = "select top 100 convert(varchar(10),tgl,120) as tgl, cycle from yamaha_print_ds_OES group by tgl, cycle order by tgl desc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getMsPart(ByVal tgl As Date) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("select distinct ym_item_nbr, ms_part_desc1 from yamaha_order a inner join yamaha_ms_part b on a.ym_item_nbr=replace(b.ms_part_cust,'-','') where ym_del_duedate='" & tgl & "' and ym_user_code <> '7618' group by ym_item_nbr, ms_part_desc1 order by ms_part_desc1 asc", connectDB.sqlConnection)

            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getValueShift(ByVal tgl As Date, ByVal pn As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            myDA = New SqlDataAdapter("select sum(ym_order_qty-ym_qty_scan) as outs, sum(ym_qty_scan) qty_scan, ym_item_nbr from yamaha_order " &
                   " where ym_del_duedate='" & tgl & "' and ym_item_nbr='" & pn & "' and ym_user_code <>'7618' AND ym_item_nbr not like '%J0' group by ym_item_nbr", connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return Nothing
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
            ''MsgBox(ex.Message)
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
            ''MsgBox(ex.Message)
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
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function getValidDS(ByVal DS As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select ds_no from yamaha_print_ds_OES where ds_no='" & DS & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function getValidDSOES(ByVal DS As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select ds_no from yamaha_print_ds_OES where ds_no='" & DS & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
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
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getTruck(ByVal nopol As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select nopol from yamaha_ms_truck where nopol='" & nopol & "'", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
            id = ""
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function updateStatusTruck(ByVal truck As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "update yamaha_load_trms set status_kirm='True' where id_trans='" & truck & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    'Public Function getTransOrder(ByVal order As String) As DataTable
    '    Dim item_nbr As String = ""
    '    Try
    '        connectDB.connectQAD()
    '        Dim myDA As New Odbc.OdbcDataAdapter
    '        Dim myDT As New DataTable
    '        'Dim a As String = "SELECT tr_nbr FROM PUB.tr_hist WHERE tr_domain ='AAIJ' AND tr_site='AAIJ' AND tr_userid='MFG' AND tr_type='ISS-TR' and substring(tr_NBR,1,11)= '" & order & "' ORDER BY tr_nbr DESC WITH(nolock)"
    '        myDA = New Odbc.OdbcDataAdapter("SELECT tr_nbr FROM PUB.tr_hist WHERE tr_domain ='AAIJ' AND tr_site='AAIJ' AND tr_userid='MFG' AND tr_type='ISS-TR' and tr_NBR= '" & order & "' ORDER BY tr_nbr asc WITH(nolock)", connectDB.myConnectionQAD)
    '        myDA.Fill(myDT)
    '        Return myDT
    '    Catch ex As Exception
    ''        'MsgBox(ex.Message)
    '        Return Nothing
    '    Finally
    '        connectDB.disconnectQAD()
    '    End Try
    'End Function

    Public Function getIDTR(ByVal idtr As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select top 1 id from yamaha_tr_hist where tr_id='" & idtr & "' order by id desc", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function updatePostponeMs(ByVal transid As String, ByVal category As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            If category = "POSTPONE" Then
                cmd = "update yamaha_order set ym_postpone='True' where id='" & transid & "'"
            Else
                cmd = "update yamaha_order set ym_safety='True' where id='" & transid & "'"
            End If

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updatePostponeTr(ByVal transid As String, ByVal category As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            If category = "POSTPONE" Then
                cmd = "update yamaha_tr_hist set tr_postpone='True',tr_postpone_date=getdate() where tr_id='" & transid & "' and tr_loading_st='False'"
            Else
                cmd = "update yamaha_tr_hist set tr_safety='True',tr_safety_date=getdate() where tr_id='" & transid & "' and tr_loading_st='False'"
            End If

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updatePostponeDS(ByVal transid As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd = "update yamaha_order set ym_postpone='False',ym_safety='False' where id='" & transid & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function


    Public Function insertTrSafety(ByVal kanban As String, ByVal tag As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "insert into yamaha_tr_safety (kanban, tag_produksi) values ('" & kanban & "','" & tag & "')"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertSafetyDetail(ByVal tag As String, ByVal pc As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "insert into yamaha_pallet_dtl_sft (palet_code, lot, pc) values ('-','" & tag & "','" & pc & "')"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertSafetyMaster(ByVal palet As String, ByVal effdate As String, ByVal pc As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            cmd = "insert into yamaha_pallet_mstr_sft (palet_code, effdate) values ('" & palet & "','" & effdate & "');"
            cmd += "update yamaha_pallet_dtl_sft set palet_code = '" & palet & "' where palet_code = '-' and pc = '" & pc & "';"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function deleteSafety(ByVal pc As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd += "DELETE yamaha_pallet_dtl_sft where palet_code = '-' and pc = '" & pc & "';"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateStockSafetyTambah(ByVal itemno As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd = "update yamaha_ms_part set stock_safety += ms_part_snp where ms_item_nbr = '" & itemno & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateStockSafetyKurang(ByVal itemno As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd = "update yamaha_ms_part set stock_safety = stock_safety - ms_part_snp where ms_item_nbr = '" & itemno & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateStockKurang(ByVal itemno As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd = "update yamaha_ms_part set stock_fg = stock_fg - ms_part_snp where ms_part_cust like '%" & itemno & "%'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getTotalTag(ByVal kanban As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select count(id) as total FROM [yamaha_tr_safety] where kanban = '" & kanban & "' and status_pull = 0", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function cekTag(ByVal tag As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select kanban FROM [yamaha_tr_safety] where kanban = '" & tag & "' and status_pull = 0", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    'Public Function cekPnKanban(ByVal tag As String) As String
    '    Dim id As String = ""
    '    Try
    '        connectDB.connectSQLKanban()
    '        Dim myCommand As New SqlCommand("select top 1 item_no from kanban_print a " &
    '        "inner join kanban_tr_kanban b on a.kanban_no=b.kanban_no " &
    '        "inner join kanban_tr_store c on b.kanban_back_no_a=c.no " &
    '        "inner join kanban_ms_item d on c.store_back_no=d.item_back_no where a.print_no = '" & tag & "'", connectDB.sqlConnectionKanban)
    '        Dim myDR As SqlDataReader = myCommand.ExecuteReader()

    '        While myDR.Read
    '            id = myDR(0)
    '        End While

    '        myDR.Close()
    '    Catch ex As Exception
    '        ''MsgBox(ex.Message)
    '    Finally
    '        connectDB.disconnectSQLKanban()
    '    End Try

    '    Return id
    'End Function

    Public Function cekPnTag(ByVal tag As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select top 1 ms_part_cust from yamaha_ms_part where ms_item_nbr = '" & tag & "' and row_name is not null", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function cekSnp(ByVal tag As String) As String
        Dim id As String = ""
        Try
            connectDB.connectSQLBGP()
            Dim myCommand As New SqlCommand("select top 1 ms_part_snp from yamaha_ms_part where (ms_item_nbr like '%" + tag + "%' or ms_part_cust like '%" + tag + "%') and row_name is not null", connectDB.sqlConnectionBGP)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
        Finally
            connectDB.disconnectSQLBGP()
        End Try

        Return id
    End Function

    Public Function updateStatusTag(ByVal tag As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""

            cmd = "update yamaha_tr_safety set status_pull = 1 where tag_produksi = '" & tag & "'"

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLogStatus(ByVal fulltag As String, ByVal stt As String)
        Try
            connectDB.connectSQL()
            Dim cmd As String
            If stt = "1" Then
                cmd = "update [dbo].[yamaha_tr_log_fg] set tr_status_out='True', tr_out_date=getdate(),tr_safety='False',tr_safety_out_date=getdate() where lot='" & fulltag & "' and tr_status='IN'"
            Else
                cmd = "update [dbo].[yamaha_tr_log_fg] set tr_status_out='True', tr_out_date=getdate(),tr_safety='False' where lot='" & fulltag & "' and tr_status='IN'"
            End If

            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateLogStatusSafety(ByVal fulltag As String)
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "update [dbo].[yamaha_tr_log_fg] set tr_safety='True', tr_safety_in_date=getdate() where lot='" & fulltag & "' and tr_status='IN'"
            connectDB.executeDBSQL(cmd)

            Return True
        Catch ex As Exception
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function insertDS(ByVal ds_no As String, ByVal noteXml As String, ByVal dstag As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_print_ds_OES (ds_no,  noteXml, ds_tag, shift, line_name) values ('" & ds_no & "',  '" & noteXml & "','" & dstag & "', '" & frm_barcode.Shift.Text & "','OEM')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getNoDs() As String
        Dim id As String = 0
        Try
            connectDB.connectSQL()
            Dim myCommand As New SqlCommand("select top 1 idprint from yamaha_print_ds_OES order by idprint desc", connectDB.sqlConnection)
            Dim myDR As SqlDataReader = myCommand.ExecuteReader()

            While myDR.Read
                id = myDR(0)
            End While

            myDR.Close()
        Catch ex As Exception
            ''MsgBox(ex.Message)
            id = 0
        Finally
            connectDB.disconnectSQL()
        End Try

        Return id
    End Function

    Public Function insertDSDet(ByVal idorder As Integer, ByVal ds_no As String, ByVal kres As String, ByVal po As String, ByVal part_desc As String, ByVal model As String, ByVal qty As Decimal, ByVal dstag As String, ByVal cycle As String, ByVal pnaaij As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String
            cmd = "insert into yamaha_print_dsdet_OES (idorder, ds_no, kres, po, part_desc, model, qty, ds_tag, cycle, pn_aaij) values ('" & idorder & "','" & ds_no & "', '" & kres & "', '" & po & "','" & part_desc & "',  '" & model & "','" & qty & "','" & dstag & "', '" & cycle & "', '" & pnaaij & "')"
            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Return False
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
            query = "select * FROM yamaha_print_dsdet_OES where ds_no = '" & id & "'"


            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function
    Public Function createDS3(ByVal id As String) As String
        Try
            connectDB.connectSQL()

            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable

            Dim query As String = ""

            query = "SELECT TOP 1 * 
                From yamaha_print_ds_OES
                Where SUBSTRING(ds_no, 5, 4) = '" & id & "'
                ORDER BY TRY_CAST(SUBSTRING(ds_no, 10, 5) AS INT) DESC;"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            If myDT.Rows.Count > 0 Then
                Return myDT.Rows(0)("ds_no").ToString()
            Else
                Return "Data tidak ditemukan"
            End If
        Catch ex As Exception
            ' Jika terjadi error, kembalikan pesan error
            Return "Error: " & ex.Message
        Finally


            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function updateDS(ByVal transid As String, ByVal dsno As String, ByVal statDs As String, ByVal noteDS As String, ByVal noteError As String, ByVal xmlSend As String) As Boolean
        Try
            connectDB.connectSQL()
            Dim cmd As String = ""
            If statDs <> "-" Then
                cmd = "update yamaha_print_ds_OES Set ds_no='" & dsno & "',stat_ds='" & statDs & "',noteXml='" & noteDS & "',creadate_ds=GETDATE(), noteError='" & noteError & "', xmlSend='" & xmlSend & "' where idprint='" & transid & "'"
            Else
                cmd = "update yamaha_print_ds_OES set noteXml='" & noteDS & "',creadate_ds=GETDATE(), noteError='" & noteError & "', xmlSend='" & xmlSend & "' where idprint='" & transid & "'"
            End If

            connectDB.executeDBSQL(cmd)
            Return True
        Catch ex As Exception
            ''MsgBox(ex.Message)
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
            ''MsgBox(ex.Message)
            Return False
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
            query = "select SUM(qty) as qty, model, pn_aaij FROM yamaha_print_dsdet_OES where ds_no = '" & id & "' group by model, pn_aaij"


            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)
            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function

    Public Function getCreateDS(ByVal ds As String) As DataTable
        Try
            connectDB.connectSQL()
            Dim myDA As New SqlDataAdapter
            Dim myDT As New DataTable
            Dim query As String = ""


            'executing the command and assigning it to connection
            query = "select  * from yamaha_print_dsdet_OES where ds_no ='" & ds & "' order by idid asc"

            'query = query & " order by del_date asc"
            myDA = New SqlDataAdapter(query, connectDB.sqlConnection)
            myDA.Fill(myDT)

            Return myDT
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        Finally
            connectDB.disconnectSQL()
        End Try
    End Function
End Class
