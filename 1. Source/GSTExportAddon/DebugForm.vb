Public Class DebugForm

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            PublicVariable.oCompany.Server = "10.0.20.105:30015"
            PublicVariable.oCompany.LicenseServer = "10.0.20.105:40000"
            PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            PublicVariable.oCompany.DbPassword = "Sapb1hana"
            PublicVariable.oCompany.DbUserName = "SYSTEM"

            PublicVariable.oCompany.UserName = "manager"
            PublicVariable.oCompany.Password = "1234"
           
            PublicVariable.oCompany.CompanyDB = "SBOMYEG_SERVICESTRAINING1"

            Dim erc As Integer = PublicVariable.oCompany.Connect()
            If erc <> 0 Then
                MessageBox.Show(PublicVariable.oCompany.GetLastErrorDescription)
            Else
                MessageBox.Show("Connected!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim opr As New oPrint
        Dim str As String
        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            str = opr.SAPPrintCrystalReport_HANA(PublicVariable.SAPPass, "@pointat;@fromdate;@todate;@duedate", "All;2016-12-12;2016-12-12;2016-12-12", "GST03-HANA.rpt") ' rptFile)
        Else
            str = opr.SAPPrintCrystalReport(PublicVariable.SAPPass, "@pointat;@fromdate;@todate;@duedate", "", "GST03.rpt")
        End If

        If str <> "" Then
            MessageBox.Show(str)
        End If
    End Sub

    Private Sub DebugForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
        PublicVariable.SAPPass = "Sapb1hana"
    End Sub
End Class