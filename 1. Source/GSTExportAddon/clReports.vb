Imports SAPbobsCOM
Imports System.IO

Public Class clReports
    Public Function GenerateReport() As String
        Try
            Dim oLayoutService As ReportLayoutsService = PublicVariable.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
            Dim oReport As ReportLayout = oLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
            Dim oNewReportParams As ReportLayoutParams

            'oReport = oLayoutService.GetReportLayout(oNewReportParams)
            '---------------------ADD GST 03 Report----------------------------

            oReport.Name = "GST 03 Report"
            oReport.TypeCode = "RCRI"
            oReport.Author = oCompany.UserName
            oReport.Category = ReportLayoutCategoryEnum.rlcCrystal
            oReport.Remarks = PublicVariable.Version
            Dim newReportCode As String
            oNewReportParams = oLayoutService.AddReportLayoutToMenu(oReport, "9728")
            newReportCode = oNewReportParams.LayoutCode


            Dim rptFilePath As String = "C:\Malaysia\GST Addon\GSTExportAddon\GSTScripts\GST.rpt"

            Dim oCompanyService As CompanyService = PublicVariable.oCompany.GetCompanyService()
            Dim oBlobParams As BlobParams = oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            Dim oKeySegment As BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = newReportCode

            Dim oBlob As Blob = oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob)

            Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)

            Dim fileSize As Integer = CInt(oFile.Length)
            Dim buf As Byte() = New Byte(fileSize - 1) {}

            oFile.Read(buf, 0, fileSize)
            oFile.Close()

            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
            oCompanyService.SetBlob(oBlobParams, oBlob)

            '---------------------ADD GST 03 DETAIL Report----------------------------
            oReport.Name = "GST 03 Detail Report"
            oReport.TypeCode = "RCRI"
            oReport.Author = oCompany.UserName
            oReport.Category = ReportLayoutCategoryEnum.rlcCrystal
            oReport.Remarks = PublicVariable.Version

            oNewReportParams = oLayoutService.oLayoutService.AddReportLayoutToMenu(oReport, 9728)
            newReportCode = oNewReportParams.LayoutCode


            rptFilePath = "C:\Malaysia\GST Addon\GSTExportAddon\GSTScripts\GSTDetail.rpt"

            oCompanyService = PublicVariable.oCompany.GetCompanyService()
            oBlobParams = oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"
            oKeySegment.Value = newReportCode

            oBlob = oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob)

            oFile = New FileStream(rptFilePath, System.IO.FileMode.Open)

            fileSize = CInt(oFile.Length)
            buf = New Byte(fileSize - 1) {}

            oFile.Read(buf, 0, fileSize)
            oFile.Close()

            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
            oCompanyService.SetBlob(oBlobParams, oBlob)

            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
End Class
