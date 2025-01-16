Public Class frm_MAIN

    Private Sub Label7_Click(sender As System.Object, e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub btnScanner_Click(sender As System.Object, e As System.EventArgs) Handles btnScanner.Click
        'frm_barcode.MdiParent = Me
        frm_barcode.ShowDialog()
        frm_barcode.BringToFront()
    End Sub

    Private Sub frm_MAIN_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'frm_barcode.MdiParent = Me
        frm_barcode.ShowDialog()
        frm_barcode.BringToFront()
    End Sub

    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnSetting_Click(sender As System.Object, e As System.EventArgs) Handles btnSetting.Click
        frm_cetak.MdiParent = Me
        frm_cetak.Show()
        frm_cetak.BringToFront()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If frm_Login.login_admin = "T" Then
            frm_User.Show()
            frm_User.BringToFront()
        Else
            MsgBox("Administrator Only!")
        End If
    End Sub

    Private Sub btnLogOff_Click(sender As System.Object, e As System.EventArgs) Handles btnLogOff.Click
        Me.Close()
        Me.Dispose()
        frm_Login.ShowDialog()
    End Sub
End Class