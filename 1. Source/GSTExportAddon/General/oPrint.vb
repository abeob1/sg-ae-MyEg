Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class oPrint
    Public Function PrintCrystalReport1(ReportFile As String, PrinterName As String, ConStr As String, _
                                   ParaName As String, ParaValue As String, OutputType As Integer, PDFFile As String) As String
        'OutputType: 0: show, 1: pdf, 2: to printer

        Dim fReportViewer As New frmReport
        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
        Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim MyArr As Array = ConStr.Split(";")
        ' Create a report document instance to hold the report

        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Try
            ' Load the report 
            Dim rptReportDoc As New ReportDocument
            rptReportDoc.Load(ReportFile)

            'Set DB con
            With crConnectionInfo
                .ServerName = MyArr(3)
                .DatabaseName = MyArr(0)
                .UserID = MyArr(4)
                .Password = MyArr(5)
            End With

            'Apply DB con
            CrTables = rptReportDoc.Database.Tables

            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            'add parameter and value
            Dim MyArr1 As Array = ParaName.Split(";")
            Dim MyArr2 As Array = ParaValue.Split(";")
            For i As Integer = 0 To MyArr1.Length - 1
                Para.Value = MyArr2(i)
                pvCollection.Add(Para)
                rptReportDoc.DataDefinition.ParameterFields(MyArr1(i)).ApplyCurrentValues(pvCollection)
            Next


            Dim doctoprint As New System.Drawing.Printing.PrintDocument()
            doctoprint.PrinterSettings.PrinterName = PrinterName

            fReportViewer.CrystalReportViewer1.ReportSource = rptReportDoc
            rptReportDoc.PrintOptions.PrinterName = PrinterName
            rptReportDoc.PrintOptions.PaperSize = PaperSize.DefaultPaperSize

            Select Case OutputType
                Case 0
                    fReportViewer.showForm(rptReportDoc)
                Case 1
                    rptReportDoc.ExportToDisk(ExportFormatType.PortableDocFormat, PDFFile)
                Case 2
                    rptReportDoc.PrintToPrinter(1, False, 0, 0)
            End Select

            rptReportDoc.Dispose()
            fReportViewer.CrystalReportViewer1.Dispose()
            Return ""
        Catch Exp As Exception
            Return Exp.ToString
        End Try
    End Function
    Public Sub PrintCrystalReport_(ParaName As String, ParaValue As String, ReportFile As String)
        'Dim frm As New frmReportX
        'frm.LoadReport(ParaName, ParaValue, ReportFile)
    End Sub
    Public Function PrintCrystalReport(ParaName As String, ParaValue As String, ReportFile As String) As String
        'OutputType: 0: show, 1: pdf, 2: to printer

        Dim fReportViewer As New frmReport
        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
        Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue

        Dim MyArr As Array = System.Configuration.ConfigurationSettings.AppSettings.Get("LocalConnection").ToString.Split(";")
        ' Create a report document instance to hold the report

        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Try
            ' Load the report 
            Dim rptReportDoc As New ReportDocument
            rptReportDoc.Load(Application.StartupPath + "\" + ReportFile)

            'Set DB con
            With crConnectionInfo
                .ServerName = MyArr(3)
                .DatabaseName = MyArr(0)
                .UserID = MyArr(4)
                .Password = MyArr(5)
            End With

            'Apply DB con
            CrTables = rptReportDoc.Database.Tables

            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            'add parameter and value
            Dim MyArr1 As Array = ParaName.Split(";")
            Dim MyArr2 As Array = ParaValue.Split(";")
            For i As Integer = 0 To MyArr1.Length - 1
                If MyArr1(i) <> "" Then
                    Para.Value = MyArr2(i)
                    pvCollection.Add(Para)
                    rptReportDoc.DataDefinition.ParameterFields(MyArr1(i)).ApplyCurrentValues(pvCollection)
                End If

            Next


            Dim doctoprint As New System.Drawing.Printing.PrintDocument()

            fReportViewer.CrystalReportViewer1.ReportSource = rptReportDoc
            rptReportDoc.PrintOptions.PaperSize = PaperSize.DefaultPaperSize

            fReportViewer.showForm(rptReportDoc)

            rptReportDoc.Dispose()
            fReportViewer.CrystalReportViewer1.Dispose()
            Return ""
        Catch Exp As Exception
            Return Exp.ToString
        End Try
    End Function
End Class
