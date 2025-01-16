Public Class frm_Safety
    Private Sub btnCasemark_Click(sender As Object, e As EventArgs) Handles btnCasemark.Click
        frm_cetakCasemark.Show()
        frm_cetakCasemark.BringToFront()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnTag_Click(sender As Object, e As EventArgs) Handles btnTag.Click
        frm_cetakTag.Show()
        frm_cetakTag.BringToFront()
        Me.Close()
        Me.Dispose()
    End Sub
End Class