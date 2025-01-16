Imports System.Data
Imports System.Data.Sql
Imports System
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Math

Public Class frm_loading

    Dim truckid As String = ""
    Dim paletid As String = ""
    Dim trid As String = ""

    Private Sub myCancel()
        txtTruck.Text = ""
        txtTruck.Focus()
        txtPalet.Text = ""
        truckid = ""
        paletid = ""
        trid = ""
    End Sub

    Private Sub frm_cetak_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        myCancel()
    End Sub

    Private Sub txtTruck_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtTruck.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Dim myTemp As New queryDB
            Dim myDT As New DataTable
            txtTruck.Text = txtTruck.Text.ToUpper
            txtPalet.Focus()
            truckid = txtTruck.Text
            If myTemp.getTruck(truckid) <> "" Then
                trid = myTemp.getLoadId() + 1
                trid = frm_barcode.cstTime.ToString("yyMMdd") & trid.PadLeft(2, "0"c)
                myTemp.insertLoading(truckid, trid)
            Else
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                myCancel()
            End If
        End If
    End Sub

    Private Sub txtPalet_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPalet.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Dim myTemp As New queryDB
            Dim myDT As New DataTable
            txtPalet.Text = txtPalet.Text.ToUpper
            paletid = txtPalet.Text
            If myTemp.getValidPalet(paletid) <> "" Or myTemp.getValidDS(paletid) <> "" Or myTemp.getValidDSOES(paletid) <> "" Then
                If myTemp.insertLoadingDet(paletid, trid) = True Then
                    My.Computer.Audio.Play(My.Resources.okay, AudioPlayMode.Background)
                    txtPalet.Text = ""
                    txtPalet.Focus()
                Else
                    My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                    txtPalet.Text = ""
                    txtPalet.Focus()
                End If
            ElseIf paletid.Contains("FINISH") Then
                myTemp.updateStatusTruck(trid)
                My.Computer.Audio.Play(My.Resources.okay , AudioPlayMode.Background)
                myCancel()
                Me.DialogResult = DialogResult.Cancel
            Else
                My.Computer.Audio.Play(My.Resources.ALARME, AudioPlayMode.Background)
                txtPalet.Text = ""
                txtPalet.Focus()
            End If
        End If
    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        myCancel()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub
End Class