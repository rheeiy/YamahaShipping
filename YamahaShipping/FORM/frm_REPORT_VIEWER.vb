Imports CrystalDecisions.Shared

Public Class frm_REPORT_VIEWER
    Public Sub configureCrystalReports()
        Dim myconnectioninfo As ConnectionInfo = New ConnectionInfo()
        myconnectioninfo.DatabaseName = "dotNET"
        myconnectioninfo.UserID = "sa"
        myconnectioninfo.Password = "r3m04aaij"
        SetDBLogOnForReport(myconnectioninfo)
    End Sub

    Private Sub SetDBLogOnForReport(ByVal myConnectionInfo As ConnectionInfo)
        Dim myTableLogonInfos As TableLogOnInfos = crRpt.LogOnInfo
        For Each myTabeLogoninfo As TableLogOnInfo In myTableLogonInfos
            myTabeLogoninfo.ConnectionInfo = myConnectionInfo
        Next
    End Sub

    Private Sub frm_REPORT_VIEWER_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        configureCrystalReports()
    End Sub
End Class