Module modMain

    Public Structure CompanyDefault
        
        Public sPGSQLServer As String
        Public sPGDatabase As String
        Public sPGUserId As String
        Public sPGPassword As String
        Public sDummyCust As String
    End Structure

    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public p_oEventHandler As clsEventHandler
    Public WithEvents p_oSBOApplication As SAPbouiCOM.Application
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oUICompany As SAPbouiCOM.Company
    Public p_oCompDef As CompanyDefault
    Public sFuncName As String
    Public sErrDesc As String

    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

    <STAThread()>
    Sub Main(ByVal args() As String)
        sFuncName = "Main()"
        Try
            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            p_oApps = New SAPbouiCOM.SboGuiApi
            p_oApps.Connect(args(0))

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", sFuncName)
            p_oSBOApplication = p_oApps.GetApplication

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = p_oSBOApplication.Company

            p_oDICompany = New SAPbobsCOM.Company
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting Company", sFuncName)
            If Not p_oDICompany.Connect() Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", sFuncName)
            p_oEventHandler = New clsEventHandler(p_oSBOApplication, p_oDICompany)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddMenuItems()", sFuncName)
            p_oEventHandler.AddMenuItems()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("System Initialization", sFuncName)
            If GetCompanyInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

            p_oSBOApplication.StatusBar.SetText("Addon Started Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Sub
End Module
