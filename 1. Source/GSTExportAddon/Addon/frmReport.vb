Imports System.Drawing.Printing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO


Public Class frmReport
    Dim rptDocument As ReportDocument

    Public Sub showForm(ByRef myReport As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        CrystalReportViewer1.ReportSource = myReport
        rptDocument = myReport
        CrystalReportViewer1.Visible = True
        CrystalReportViewer1.Show()
        Me.ShowDialog()
    End Sub


    Private Sub frmReport_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Application.DoEvents()
        Me.BringToFront()
        Me.TopMost = True
    End Sub
End Class