Public Class frm_LOV_summary

    Public namaForm As String = ""
    Public parameterID As String
    Public parameterID2 As String

    Private Sub txtCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frm_LOV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadListSummary()
    End Sub

    Private Sub loadListSummary()
        dgLOV.Columns.Clear()
        dgLOV.Rows.Clear()

        Dim myTemp As New queryDB
        Dim myDT As New DataTable
        myDT = myTemp.getSummary()

        If myDT IsNot Nothing And myDT.Rows.Count > 0 Then
            Dim no As Integer = 1
            dgLOV.Columns.Add("No", "No. ")
            dgLOV.Columns.Add("col1", "Tanggal")
            dgLOV.Columns.Add("col2", "Cycle")
            dgLOV.Rows.Add(myDT.Rows.Count)
            dgLOV.AutoResizeColumns()

            For i As Integer = 0 To myDT.Rows.Count - 1
                dgLOV.Item(0, i).Value = no.ToString
                dgLOV.Item(1, i).Value = myDT.Rows(i).Item("tgl").ToString
                dgLOV.Item(2, i).Value = myDT.Rows(i).Item("cycle").ToString
                no += 1
            Next

            dgLOV.AutoResizeColumns()
            dgLOV.CurrentCell = dgLOV.Item(0, 0)
        Else
            MessageBox.Show("You have no project to display !!", "Project", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub txtOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOK.Click
        frm_cetak_summary.txtLabel.Text = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        frm_cetak_summary.tgl = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        frm_cetak_summary.cyc = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        If isValid() Then
            loadListSummary()
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
        frm_cetak_summary.txtLabel.Text = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        frm_cetak_summary.tgl = dgLOV.Item(1, dgLOV.CurrentRow.Index).Value
        frm_cetak_summary.cyc = dgLOV.Item(2, dgLOV.CurrentRow.Index).Value
        Me.Close()
        Me.Dispose()
    End Sub
End Class