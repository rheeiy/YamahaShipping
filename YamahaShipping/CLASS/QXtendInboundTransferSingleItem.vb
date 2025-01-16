Imports System.Xml

Public Class QXtendInboundTransferSingleItem
    Public keterangan_qx As String
    Public note_xml As String
    Public xmlSend As String = ""

    Public Function requestInbound(ByVal itm As String, ByVal effdate As String, qty As Decimal, ByVal num As String, ByVal loc_from As String, ByVal loc_to As String, ByVal boxno As String) As Boolean
        Dim entryPoint As String
        Dim qDoc As String
        Dim header As String
        Dim data As String
        Dim footer As String
        Dim rs_error As String

        'entryPoint = "https://ssi0fc6xdb01.odqad.com:8143/qxiqonddb/services/QdocWebService"
        entryPoint = "https://ssi0fc6xdb03.odqad.com:8143/qxiqonddb/services/QdocWebService"
        'entryPoint = "http://192.168.0.42:8080/qxilive/services/QdocWebService"
        'entryPoint = "http://qadee2014.akebono-astra.co.id:8989/qxitdw/services/QdocWebService"
        'entryPoint = "http://qadee2014.akebono-astra.co.id:8989/qxisim/services/QdocWebService"

        header = "<?xml version=""1.0"" encoding=""UTF-8""?>" &
                 "<soapenv:Envelope xmlns=""urn:schemas-qad-com:xml-services"" xmlns:qcom=""urn:schemas-qad-com:xml-services:common"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:wsa=""http://www.w3.org/2005/08/addressing"">" &
                    "<soapenv:Header>" &
                        "<wsa:Action/>" &
                        "<wsa:To>urn:services-qad-com:qad2014ee</wsa:To>" &
                        "<wsa:MessageID>urn:services-qad-com::qad2014ee</wsa:MessageID>" &
                        "<wsa:ReferenceParameters>" &
                            "<qcom:suppressResponseDetail>true</qcom:suppressResponseDetail>" &
                        "</wsa:ReferenceParameters>" &
                        "<wsa:ReplyTo>" &
                            "<wsa:Address>urn:services-qad-com:</wsa:Address>" &
                        "</wsa:ReplyTo>" &
                    "</soapenv:Header>" &
                    "<soapenv:Body>" &
                        "<transferInvSingleItem>" &
                            "<qcom:dsSessionContext>" &
                                "<qcom:ttContext>" &
                                    "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                    "<qcom:propertyName>domain</qcom:propertyName>" &
                                    "<qcom:propertyValue>AAIJ</qcom:propertyValue>" &
                                "</qcom:ttContext>" &
                                "<qcom:ttContext>" &
                                    "<qcom:propertyQualifier>QAD</qcom:propertyQualifier>" &
                                    "<qcom:propertyName>scopeTransaction</qcom:propertyName>" &
                                    "<qcom:propertyValue>false</qcom:propertyValue>" &
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

        data = "<dsItem>" &
                "<item>" &
                    "<operation>A</operation>" &
                    "<part>" & itm & "</part>" &
                    "<itemDetail>" &
                        "<operation>A</operation>" &
                        "<lotserialQty>" & qty & "</lotserialQty>" &
                        "<nbr>" & num & "</nbr>" &
                        "<rmks>" & boxno & "</rmks>" &
                        "<siteFrom>AAIJ</siteFrom>" &
                        "<locFrom>" & loc_from & "</locFrom>" &
                        "<locTo>" & loc_to & "</locTo>" &
                        "<siteTo>AAIJ</siteTo>" &
                        "<effDate>" & effdate & "</effDate>" &
                        "<yn>true</yn>" &
                    "</itemDetail>" &
                "</item>" &
            "</dsItem>" &
        "</transferInvSingleItem>"

        footer = "</soapenv:Body> </soapenv:Envelope>"

        qDoc = header & data & footer
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
            Return False
            Exit Function
        End Try
        Dim hasil As String = ""
        hasil = qDocResp.GetElementsByTagName("ns1:result")(0).InnerText
        keterangan_qx = ""
        rs_error = ""
        note_xml = ""

        Dim rs As String = ""
        Dim elemList As XmlNodeList = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")
        For i As Integer = 0 To elemList.Count - 1
            rs = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")(i).InnerText
            keterangan_qx = keterangan_qx & i & ". " & rs & Chr(13) & Chr(10)
        Next
        note_xml = qDocResp.InnerXml
        If (hasil.ToUpper = "ERROR") Then
            'rs_error = qDocResp.GetElementsByTagName("ns3:tt_msg_desc")(0).InnerText
            'keterangan_qx = rs_error

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
