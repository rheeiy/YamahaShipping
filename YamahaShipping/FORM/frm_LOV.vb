Public Class frm_LOV

    Public namaForm As String = ""
    Public parameterID As String
    Public parameterID2 As String

    Private Sub txtCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frm_LOV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadListDS()
    End Sub

    Private Sub loadListDS()
        dgLOV.Columns.Clear()
        dgLOV.Rows.Clear()

        Dim myTemp As New queryDB
        Dim myDT As New DataTable
        myDT = myTemp.getDS()

        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            Dim no As Integer = 1
            dgLOV.Columns.Add("No", "No. ")
            dgLOV.Columns.Add("col1", "No DS")
            'dgLOV.Columns.Add("col2", "Kres")
            'dgLOV.Columns.Add("col3", "Model")
            'dgLOV.Columns.Add("col4", "Qty")
            'dgLOV.Columns.Add("col5", "PO")
            'dgLOV.Columns.Add("col6", "DESC")
            dgLOV.Rows.Add(myDT.Rows.Count)
            dgLOV.AutoResizeColumns()

            For i As Integer = 0 To myDT.Rows.Count - 1
                dgLOV.Item(0, i).Value = no.ToString
                dgLOV.Item(1, i).Value = myDT.Rows(i).Item("ds_no").ToString
                'dgLOV.Item(2, i).Value = myDT.Rows(i).Item("kres").ToString
                'dgLOV.Item(3, i).Value = myDT.Rows(i).Item("model").ToString
                'dgLOV.Item(4, i).Value = myDT.Rows(i).Item("qty").ToString
                'dgLOV.Item(5, i).Value = myDT.Rows(i).Item("po").ToString
                'dgLOV.Item(6, i).Value = myDT.Rows(i).Item("part_desc").ToString
                no += 1
            Next

            dgLOV.AutoResizeColumns()
            dgLOV.CurrentCell = dgLOV.Item(0, 0)
        Else
            MessageBox.Show("You have no project to display !!", "Project", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub txtOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOK.Click
        frm_cetak.txtLabel.Text = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        frm_cetak.absId = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        'frm_cetak.kres = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        'frm_cetak.poAAIJ = dgLOV.Item(5, dgLOV.CurrentRow.Index).Value
        'frm_cetak.pn_desc = dgLOV.Item(6, dgLOV.CurrentRow.Index).Value
        'frm_cetak.pn_cust = dgLOV.Item(3, dgLOV.CurrentRow.Index).Value
        'frm_cetak.qty_del = dgLOV.Item(4, dgLOV.CurrentRow.Index).Value
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        If isValid() Then
            loadListDS()
        End If
    End Sub

    Private Function isValid() As Boolean
        If (txtFind.Text = "") Then
            epFind.SetError(btnFind, "Masukkan Keyword Pencarian !!")
            Return False
        Else
            epFind.SetError(btnFind, "")
        End If

        Return True
    End Function

    Private Sub dgLOV_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgLOV.DoubleClick
        frm_cetak.txtLabel.Text = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        frm_cetak.absId = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        'frm_cetak.kres = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        'frm_cetak.poAAIJ = dgLOV.Item(5, dgLOV.CurrentRow.Index).Value
        'frm_cetak.pn_desc = dgLOV.Item(6, dgLOV.CurrentRow.Index).Value
        'frm_cetak.pn_cust = dgLOV.Item(3, dgLOV.CurrentRow.Index).Value
        'frm_cetak.qty_del = dgLOV.Item(4, dgLOV.CurrentRow.Index).Value
        Me.Close()
        Me.Dispose()
    End Sub
End Class