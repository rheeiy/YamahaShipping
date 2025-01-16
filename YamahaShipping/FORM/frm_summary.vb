Public Class frm_summary
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

        dgOrder.Rows.Clear()
        dgOrder.Columns.Clear()
        myDT = myTemp.getLOTTr(trans_id)
        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            dgOrder.Columns.Add("NO", "No. ")
            dgOrder.Columns.Add("po_no", "PO NO")
            dgOrder.Columns.Add("lot", "LOT")
            'dgOrder.Columns.Add("prod_date", "PROD DATE")
            dgOrder.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgOrder.Columns("prod_date").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgOrder.Columns(0).Width = 40
            dgOrder.Columns(1).Width = 80
            dgOrder.Columns(2).Width = 90
            'dgOrder.Columns(3).Width = 120
            'dgOrder.Columns("qty_order").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Dim nomer As Integer = 1
            dgOrder.Rows.Add(myDT.Rows.Count)
            For i As Integer = 0 To myDT.Rows.Count - 1
                dgOrder.Item(0, i).Value = nomer
                dgOrder.Item(1, i).Value = myDT.Rows(i).Item("ym_po")
                dgOrder.Item(2, i).Value = myDT.Rows(i).Item("tr_lot")
                'dgOrder.Item(3, i).Value = Format(myDT.Rows(i).Item("tr_lot_time"), "yyyy-MM-dd")
                nomer += 1
            Next
        End If
    End Sub
End Class