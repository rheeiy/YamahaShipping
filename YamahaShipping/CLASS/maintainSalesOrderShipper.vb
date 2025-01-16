Imports System.Xml

Public Class maintainSalesOrderShipper
    Public keterangan_qx As String
    Public note_xml As String
    Public absId As String = ""
    Public xmlSend As String = ""

    Public Function requestInbound(ByVal idprint As String, ByVal shipto As String, ByVal poAAIJ As String) As Boolean
        Dim entryPoint As String
        Dim qDoc As String
        Dim header As String
        Dim data As String
        Dim dataDetail As String
        Dim dataDetailAll As String = ""
        Dim footer As String
        Dim rs_error As String
        Dim myTem As New queryDB
        Dim myDT As New DataTable
        Dim myTem2 As New queryDB
        Dim myDT2 As New DataTable

        Dim po As String = ""
        Dim po_line As String = ""
        Dim qty As String = ""
        'WSA
        Dim elemList As XmlNodeList = Nothing
        Dim elementXml = New XmlDocument()
        Dim rs As String = ""
        Dim query As String = ""
        'END WSA

        ''Dim poAAIJ As String = Format(frm_barcode.cstTime, "MMyy") & "PRPC"
        ''poAAIJ = "1119PRPC"
        'entryPoint = "https://ssi0fc6xdb01.odqad.com:8143/qxiqonddb/services/QdocWebService"
        entryPoint = "https://ssi0fc6xdb03.odqad.com:8143/qxiqonddb/services/QdocWebService"
        'entryPoint = "http://qadee2014.akebono-astra.co.id:8989/qxitdw/services/QdocWebService"
        '/entryPoint = "http://qadee2014.akebono-astra.co.id:8989/qxisim/services/QdocWebService"

        header = "<?xml version=""1.0"" encoding=""UTF-8""?>" &
                  "<soapenv:Envelope xmlns=""urn:schemas-qad-com:xml-services"" xmlns:qcom=""urn:schemas-qad-com:xml-services:common"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" &
                     "<soapenv:Header>" &
                         "<wsa:Action/>" &
                         "<wsa:To>urn:services-qad-com:qad2021dlv</wsa:To>" &
                         "<wsa:MessageID>urn:services-qad-com::qad2021dlv</wsa:MessageID>" &
                         "<wsa:ReferenceParameters>" &
                             "<qcom:suppressResponseDetail>false</qcom:suppressResponseDetail>" &
                         "</wsa:ReferenceParameters>" &
                         "<wsa:ReplyTo>" &
                             "<wsa:Address>urn:services-qad-com:</wsa:Address>" &
                         "</wsa:ReplyTo>" &
                     "</soapenv:Header>" &
                     "<soapenv:Body>" &
                         "<maintainSalesOrderShipper>" &
                             "<qcom:dsSessionContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>domain</qcom:propertyName>" &
                                     "<qcom:propertyValue>AAIJ</qcom:propertyValue>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>scopeTransaction</qcom:propertyName>" &
                                     "<qcom:propertyValue>true</qcom:propertyValue>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>version</qcom:propertyName>" &
                                     "<qcom:propertyValue>ERP3_1</qcom:propertyValue>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>mnemonicsRaw</qcom:propertyName>" &
                                     "<qcom:propertyValue>false</qcom:propertyValue>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>action</qcom:propertyName>" &
                                     "<qcom:propertyValue/>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>entity</qcom:propertyName>" &
                                     "<qcom:propertyValue/>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>email</qcom:propertyName>" &
                                     "<qcom:propertyValue/>" &
                                 "</qcom:ttContext>" &
                                 "<qcom:ttContext>" &
                                     "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                     "<qcom:propertyName>emailLevel</qcom:propertyName>" &
                                     "<qcom:propertyValue/>" &
                                 "</qcom:ttContext>" &
                             "</qcom:dsSessionContext>"

        data = "<dsSalesOrderShipper>" &
                    "<salesOrderShipper>" &
                        "<absShipfrom>AAIJ</absShipfrom>" &
                        "<absId></absId>" &
                        "<absShipto>" & shipto & "</absShipto>" &
                        "<vInvmov>SHIPMENT</vInvmov>" &
                        "<vCont>true</vCont>" &
                        "<vCont1>true</vCont1>" &
                        "<vCmmts>false</vCmmts>" &
                        "<vOk>true</vOk>" &
                        "<absShipvia>" & poAAIJ & "</absShipvia>" &
                        "<absTransMode>DI GUDANG SAUDARA</absTransMode>" &
                        "<absFormat>01</absFormat>" &
                        "<consShip>optional</consShip>" &
                        "<absLang>us</absLang>" &
                        "<cmmts>false</cmmts>" &
                        "<absTotalpallet>0</absTotalpallet>" &
                        "<vCmmts>false</vCmmts>"

        myDT = myTem.createDS(idprint)
        Dim poLine As String = ""
        Dim trans_id As String = ""
        Dim qty_del As Decimal = 0
        Dim pn_cust As String = ""
        Dim pn_aaij As String = ""
        Dim loc_new As String = "PPC-FG"

        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            For i As Integer = 0 To myDT.Rows.Count - 1
                'trans_id = myDT.Rows(i).Item("tr_id").ToString
                qty_del = myDT.Rows(i).Item("qty").ToString
                pn_cust = myDT.Rows(i).Item("model").ToString
                pn_aaij = myDT.Rows(i).Item("pn_aaij").ToString

                query = "SELECT sod_line FROM PUB.sod_det WHERE sod_nbr ='" & poAAIJ & "' AND sod_part='" & pn_aaij & "' " &
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

                'myDT = myTem.getLinePOSQL(poAAIJ, pn_aaij)
                'If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
                '    poLine = myDT.Rows(0).Item("sod_line").ToString
                'Else
                '    'My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                '    'MsgBox("Data PO tidak ada!")
                'End If

                'qty_del = myTem2.getTotScan2(trans_id)
                'myDT2 = myTem2.getLinePO(poAAIJ, pn_cust)
                'If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                '    poLine = myDT2.Rows(0).Item("sod_line").ToString
                'Else
                '    myDT2 = myTem2.getLinePO(poAAIJ, pn_aaij)
                '    If myDT2 IsNot Nothing And myDT2.Rows.Count > 0 Then
                '        poLine = myDT2.Rows(0).Item("sod_line").ToString
                '    End If
                'End If
                Try
                    query = "SELECT lnd_chr01 FROM PUB.lnd_det WHERE lnd_domain='AAIJ' AND lnd_part ='" & pn_aaij & "' GROUP BY lnd_part,lnd_chr01 WITH(nolock)"
                    elementXml = myTem.getwsa(query)
                    elemList = elementXml.GetElementsByTagName("f_val")
                    If elemList.Count > 0 Then
                        For z As Integer = 0 To elemList.Count - 1
                            rs = elementXml.GetElementsByTagName("f_val")(z).InnerText.Replace("||", "#")
                            loc_new = rs.Split("#"c)(0)
                        Next
                    End If
                    If loc_new = "" Then
                        loc_new = "PPC-FG"
                    End If
                Catch ex As Exception
                    MsgBox("WSA Problem " & ex.ToString)
                End Try

                dataDetail =
                       "<schedOrderItemDetail>" &
                           "<scxPart></scxPart>" &
                           "<scxPo></scxPo>" &
                           "<scxCustref></scxCustref>" &
                           "<scxOrder>" & poAAIJ & "</scxOrder>" &
                           "<scxLine>" & poLine & "</scxLine>" &
                           "<srQty>" & qty_del & "</srQty>" &
                           "<srLoc>" & loc_new & "</srLoc>" &
                           "<transUm>PC</transUm>" &
                           "<transConv>1</transConv>" &
                           "<srSite>AAIJ</srSite>" &
                           "<multiple>false</multiple>" &
                           "<vCmmts>false</vCmmts>" &
                       "</schedOrderItemDetail>"
                dataDetailAll = dataDetailAll & dataDetail
            Next
        Else
            Return False
            Exit Function
        End If
        footer = "</salesOrderShipper>" &
     "</dsSalesOrderShipper>" &
 "</maintainSalesOrderShipper>" &
 "</soapenv:Body> </soapenv:Envelope>"

        qDoc = header & data & dataDetailAll & footer
        xmlSend = qDoc
        Dim enc As New System.Text.UTF8Encoding()
        Dim req As Byte() = enc.GetBytes(qDoc)
        Dim web As Net.HttpWebRequest = CType(Net.WebRequest.Create(entryPoint), Net.HttpWebRequest)

        web.Headers.Add("SOAPAction", "")
        web.ContentLength = req.Length
        web.ContentType = "text/xml"
        web.Method = "POST"
        web.Timeout = Threading.Timeout.Infinite    '--> unlimited timeout
        'web.Timeout = 10000    '--> limited timeout in milisecond

        Dim strm As IO.Stream
        Dim resp As Net.HttpWebResponse
        Dim qDocResp = New XmlDocument()

        Try
            ' send qdoc melalui streaming
            strm = web.GetRequestStream
            strm.Write(req, 0, req.Length)    ' --> synchronus job
            strm.Close()

            ' tangkap response melalui streaming
            resp = CType(web.GetResponse(), Net.HttpWebResponse)
            strm = resp.GetResponseStream()
            qDocResp.Load(strm)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
            Exit Function
        End Try
        Dim hasil As String = ""
        hasil = qDocResp.GetElementsByTagName("ns1:result")(0).InnerText
        keterangan_qx = ""
        rs_error = ""
        note_xml = ""

        'Dim rs As String = ""
        elemList = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")
        For i As Integer = 0 To elemList.Count - 1
            rs = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")(i).InnerText
            keterangan_qx = keterangan_qx & i & ". " & rs & Chr(13) & Chr(10)
        Next

        Try
            note_xml = qDocResp.InnerXml
            absId = qDocResp.GetElementsByTagName("ns1:absId")(0).InnerText
        Catch ex As Exception
            MsgBox("Pembuatan DS gagal! " & ex.Message)
            Return False
        End Try

        If (hasil.ToUpper = "ERROR") Then
            Return False
            Exit Function
        ElseIf (hasil.ToUpper = "SUCCESS" Or hasil.ToUpper = "WARNING") Then
            'keterangan_qx = hasil
            Return True
            Exit Function
        ElseIf web.Timeout > 25000 Then
            'rs_error = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")(0).InnerText
            keterangan_qx = keterangan_qx & " Exception Else Timeout : " & hasil & web.Timeout
            Return False
            Exit Function
        Else
            keterangan_qx = keterangan_qx & " Exception Else : " & hasil
            Return True
            Exit Function
        End If
    End Function
End Class
