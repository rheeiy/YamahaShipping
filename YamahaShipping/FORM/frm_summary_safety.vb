Public Class frm_summary_safety
    Private Sub loadListDS()
        dgLOV.Columns.Clear()
        dgLOV.Rows.Clear()

        Dim myTemp As New queryDB
        Dim myDT As New DataTable
        myDT = myTemp.getSafetybyPC(frm_barcode.GetIPv4Address)

        If myDT IsNot Nothing AndAlso myDT.Rows.Count > 0 Then
            Dim no As Integer = 1
            dgLOV.Columns.Add("No", "No. ")
            dgLOV.Columns.Add("col1", "Item Number")
            dgLOV.Columns.Add("col2", "Type")
            dgLOV.Columns.Add("col3", "Total Safety")
            'dgLOV.Columns.Add("col4", "Qty")
            'dgLOV.Columns.Add("col5", "PO")
            'dgLOV.Columns.Add("col6", "DESC")
            dgLOV.Rows.Add(myDT.Rows.Count)
            dgLOV.AutoResizeColumns()

            For i As Integer = 0 To myDT.Rows.Count - 1
                dgLOV.Item(0, i).Value = no.ToString
                dgLOV.Item(1, i).Value = myDT.Rows(i).Item("ms_part_cust").ToString
                dgLOV.Item(2, i).Value = myDT.Rows(i).Item("ms_part_desc1").ToString
                dgLOV.Item(3, i).Value = myDT.Rows(i).Item("total_sft").ToString
                'dgLOV.Item(4, i).Value = myDT.Rows(i).Item("qty").ToString
                'dgLOV.Item(5, i).Value = myDT.Rows(i).Item("po").ToString
                'dgLOV.Item(6, i).Value = myDT.Rows(i).Item("part_desc").ToString
                no += 1
            Next

            dgLOV.AutoResizeColumns()
            dgLOV.CurrentCell = dgLOV.Item(0, 0)
        Else
            MessageBox.Show("Belum ada safety yang di scan", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            Me.Dispose()
        End If
    End Sub

    Private Sub frm_summary_safety_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadListDS()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frm_barcode.safetyProcessMaster()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Menampilkan kotak dialog untuk konfirmasi
        Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data ini?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        ' Jika pengguna memilih Yes, lanjutkan proses delete
        If result = DialogResult.Yes Then
            Dim myTemp As New queryDB
            myTemp.deleteSafety(frm_barcode.GetIPv4Address)
            MessageBox.Show("Data berhasil dihapus.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            Me.Dispose()
        Else
            MessageBox.Show("Proses hapus dibatalkan.", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class