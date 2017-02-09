Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Data
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Net.Mail
Imports System.Configuration
Imports System.Data.Odbc

Module modCommon

    Public Function GetCompanyInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long
        Dim sFunctName As String = String.Empty
        Dim sConnection As String = String.Empty

        Try
            sFunctName = "Get Company Initialization"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company Initialization", sFunctName)

            oCompDef.sPGDatabase = String.Empty
            oCompDef.sPGSQLServer = String.Empty
            oCompDef.sPGUserId = String.Empty
            oCompDef.sPGPassword = String.Empty
            oCompDef.sDummyCust = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PGSqlServer")) Then
                oCompDef.sPGSQLServer = ConfigurationManager.AppSettings("PGSqlServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PGSqlDatabase")) Then
                oCompDef.sPGDatabase = ConfigurationManager.AppSettings("PGSqlDatabase")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PGUserId")) Then
                oCompDef.sPGUserId = ConfigurationManager.AppSettings("PGUserId")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("PGPassword")) Then
                oCompDef.sPGPassword = ConfigurationManager.AppSettings("PGPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DummyCustomer")) Then
                oCompDef.sDummyCust = ConfigurationManager.AppSettings("DummyCustomer")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFunctName)
            GetCompanyInfo = RTN_SUCCESS

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFunctName)
            GetCompanyInfo = RTN_ERROR
        End Try

    End Function

    Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    ConnectDICompSSO()
        '   Purpose    :    Connect To DI Company Object
        '
        '   Parameters :    ByRef objCompany As SAPbobsCOM.Company
        '                       objCompany = set the SAP Company Object
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Sri
        '   Date       :    29 April 2013
        '   Change     :
        ' ***********************************************************************************
        Dim sCookie As String = String.Empty
        Dim sConnStr As String = String.Empty
        Dim sFuncName As String = String.Empty
        Dim lRetval As Long
        Dim iErrCode As Int32
        Try
            sFuncName = "ConnectDICompSSO()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            objCompany = New SAPbobsCOM.Company

            sCookie = objCompany.GetContextCookie
            sConnStr = p_oUICompany.GetConnectionContext(sCookie)
            'sConnStr = p_oSBOApplication.Company.GetConnectionContext(sCookie)
            lRetval = objCompany.SetSboLoginContext(sConnStr)

            If Not lRetval = 0 Then
                Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
            End If
            p_oSBOApplication.StatusBar.SetText("Please Wait While Company Connecting... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            lRetval = objCompany.Connect
            If lRetval <> 0 Then
                objCompany.GetLastError(iErrCode, sErrDesc)
                Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
            Else
                p_oSBOApplication.StatusBar.SetText("Company Connection Has Established with the " & objCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If
            ConnectDICompSSO = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectDICompSSO = RTN_ERROR
        End Try
    End Function

    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Try
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            ''sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            sPath = Application.StartupPath.ToString
            'oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
            oXmlDoc.Load(sPath & "\" & FileName)
            ' MsgBox(Application.StartupPath)

            Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Public Sub ShowErr(ByVal sErrMsg As String)
        ' ***********************************************************************************
        '   Function   :    ShowErr()
        '   Purpose    :    Show Error Message
        '   Parameters :  
        '                   ByVal sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Dev
        '   Date       :    23 Jan 2007
        '   Change     :
        ' ***********************************************************************************
        Try
            If sErrMsg <> "" Then
                If Not p_oSBOApplication Is Nothing Then
                    If p_iErrDispMethod = ERR_DISPLAY_STATUS Then

                        p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                    ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                        p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                    End If
                End If
            End If
        Catch exc As Exception
            WriteToLogFile(exc.Message, "ShowErr()")
        End Try
    End Sub

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Function ExecuteSQLQueryDataset(ByVal sQuery As String, ByRef sErrDesc As String) As DataSet
        Dim oPostgreODBC As OdbcConnection = New OdbcConnection
        Dim sConnection As String = "DRIVER={PostgreSQL ANSI};SERVER=" & p_oCompDef.sPGSQLServer & ";UID=" & p_oCompDef.sPGUserId & ";PWD=" & p_oCompDef.sPGPassword & ";DATABASE=" & p_oCompDef.sPGDatabase & ";"
        Dim oDbcCmd As New OdbcCommand
        Dim oDs As New DataSet

        Dim sFuncName As String = "ExecuteQuery()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oPostgreODBC = New OdbcConnection(sConnection)
            oPostgreODBC.Open()

            ''''MyCon.Open()
            oDbcCmd.CommandType = CommandType.Text
            oDbcCmd.CommandText = sQuery
            oDbcCmd.Connection = oPostgreODBC
            oDbcCmd.CommandTimeout = 0
            Dim da As New OdbcDataAdapter(oDbcCmd)
            da.Fill(oDs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)

        Finally
            oPostgreODBC.Dispose()
        End Try

        Return oDs
    End Function

    Public Function ExecuteSQLNonQuery(ByVal sQuery As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty
        Dim oPostgreODBC As OdbcConnection = New OdbcConnection
        Dim sConnection As String = "DRIVER={PostgreSQL ANSI};SERVER=" & p_oCompDef.sPGSQLServer & ";UID=" & p_oCompDef.sPGUserId & ";PWD=" & p_oCompDef.sPGPassword & ";DATABASE=" & p_oCompDef.sPGDatabase & ";"

        Dim oCon As New OdbcConnection(sConnection)
        Dim oDbcCmd As New OdbcCommand
        Dim oDs As New DataSet

        Try
            sFuncName = "ExecuteSQLNonQuery()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Fucntion...", sFuncName)
            oDbcCmd.CommandType = CommandType.Text
            oDbcCmd.CommandText = sQuery
            oDbcCmd.Connection = oCon
            If oCon.State = ConnectionState.Closed Then
                oCon.Open()
            End If
            oDbcCmd.CommandTimeout = 0
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            oDbcCmd.ExecuteNonQuery()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
            ExecuteSQLNonQuery = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            ExecuteSQLNonQuery = RTN_ERROR
        Finally
            If Not oCon Is Nothing Then
                oCon.Close()
                oCon.Dispose()
            End If
        End Try

    End Function

    Public Function ExecuteSQLQueryDataTable(ByVal sQuery As String, ByRef sErrDesc As String) As DataTable
        Dim sFuncName As String = String.Empty
        Dim oPostgreODBC As OdbcConnection = New OdbcConnection
        Dim sConnection As String = "DRIVER={PostgreSQL ANSI};SERVER=" & p_oCompDef.sPGSQLServer & ";UID=" & p_oCompDef.sPGUserId & ";PWD=" & p_oCompDef.sPGPassword & ";DATABASE=" & p_oCompDef.sPGDatabase & ";"

        Dim oCon As New OdbcConnection(sConnection)
        Dim oDbcCmd As New OdbcCommand
        Dim oDs As New DataSet

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oPostgreODBC = New OdbcConnection(sConnection)
            oPostgreODBC.Open()

            oDbcCmd.CommandType = CommandType.Text
            oDbcCmd.CommandText = sQuery
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            oDbcCmd.Connection = oPostgreODBC
            oDbcCmd.CommandTimeout = 0
            Dim da As New OdbcDataAdapter(oDbcCmd)
            da.Fill(oDs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)

        Finally
            oCon.Dispose()
        End Try
        Return oDs.Tables(0)
    End Function

    Public Function GetDataSetFromExcel(ByVal CurrFileToUpload As String, ByRef sErrDesc As String) As DataTable

        Dim MyConnection As System.Data.OleDb.OleDbConnection = Nothing
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter = Nothing
        Dim workbook As String = String.Empty
        Dim sw As StreamWriter = Nothing
        Dim oDs As New DataSet
        Dim xl As New Microsoft.Office.Interop.Excel.Application

        Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlwbook As Microsoft.Office.Interop.Excel.Workbook
        '  Dim oDT As DataTable
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "MyFunction()"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; " & _
            "data source=" & CurrFileToUpload & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1""")

            xlwbook = xl.Workbooks.Open(CurrFileToUpload)
            xlsheet = xlwbook.Sheets.Item(1)

            ''For Each sht In xlwbook.Worksheets
            '' If sht.Visible = True Then
            workbook = xlsheet.Name
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM [" & workbook & "$]", MyConnection)
            MyCommand.TableMappings.Add("Table", workbook)
            MyCommand.Fill(oDs)

            '' End If
            ''  Next


            GetDataSetFromExcel = oDs.Tables(0)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while reading content of  " & ex.Message, sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            Return Nothing
        Finally

            xl.ActiveWorkbook.Close(False)
            xl.Quit()
            xlwbook = Nothing
            xl = Nothing
            MyCommand.Dispose()
            MyConnection.Dispose()
        End Try

    End Function

End Module
