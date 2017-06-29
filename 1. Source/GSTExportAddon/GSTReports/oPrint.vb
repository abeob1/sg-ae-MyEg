Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class oPrint


    Public Function SAPPrintCrystalReport_HANA(SQLPass As String, ParaName As String, ParaValue As String, ReportFile As String) As Boolean
        Try

            Dim fReportViewer As New frmReport
            Dim rptReportDoc As New ReportDocument()
            For i As Integer = 1 To 100000
                Application.DoEvents()
            Next
            rptReportDoc.Load(Application.StartupPath + "\" + ReportFile)



           

            Dim strConnection As String = Convert.ToString("DRIVER= {B1CRHPROXY32};UID=SYSTEM")
            strConnection += Convert.ToString(";PWD=" + SQLPass + ";SERVERNODE=") & PublicVariable.oCompany.Server
            strConnection += (Convert.ToString(";DATABASE=") & PublicVariable.oCompany.CompanyDB) + ";"

            Dim logonProps2 As NameValuePairs2 = rptReportDoc.DataSourceConnections(0).LogonProperties
            logonProps2.[Set]("Provider", "B1CRHPROXY32")
            logonProps2.[Set]("Server Type", "B1CRHPROXY32")
            logonProps2.[Set]("Connection String", strConnection)
            logonProps2.[Set]("Locale Identifier", "1033")
            logonProps2.[Set]("QE_DatabaseType", "ODBC (RDO)")

            rptReportDoc.DataSourceConnections(0).SetLogonProperties(logonProps2)
            rptReportDoc.DataSourceConnections(0).SetConnection(PublicVariable.oCompany.Server, PublicVariable.oCompany.CompanyDB, "SYSTEM", SQLPass)

            'add parameter and value
            Dim MyArr1 As Array = ParaName.Split(";")
            Dim MyArr2 As Array = ParaValue.Split(";")
            For i As Integer = 0 To MyArr1.Length - 1
                If MyArr1(i) <> "" Then
                    rptReportDoc.SetParameterValue(i, MyArr2(i).ToString)
                End If

            Next

            fReportViewer.CrystalReportViewer1.ReportSource = rptReportDoc
            rptReportDoc.PrintOptions.PaperSize = PaperSize.DefaultPaperSize

            fReportViewer.showForm(rptReportDoc)
            'Application.DoEvents()
            'rptReportDoc.Dispose()
            'fReportViewer.CrystalReportViewer1.Dispose()
            'For i As Integer = 1 To 100000
            '    Application.DoEvents()
            'Next
            Return ""

        Catch er As Exception
            Functions.WriteLog(er.ToString)
            Return er.ToString
        End Try
    End Function


    Public Function SAPPrintCrystalReport(SQLPass As String, ParaName As String, ParaValue As String, ReportFile As String, Optional PrinterName As String = "") As String
        'OutputType: 0: show, 1: pdf, 2: to printer

        Dim fReportViewer As New frmReport
        Dim pvCollection As New CrystalDecisions.Shared.ParameterValues
        Dim Para As New CrystalDecisions.Shared.ParameterDiscreteValue

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
                .ServerName = PublicVariable.oCompany.Server
                .DatabaseName = PublicVariable.oCompany.CompanyDB
                .UserID = PublicVariable.oCompany.DbUserName
                .Password = SQLPass
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


            'Dim doctoprint As New System.Drawing.Printing.PrintDocument()

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
