Module clAppEvent
    Public WithEvents oApp As SAPbouiCOM.Application
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApp.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                oCompany.Disconnect()
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                oCompany.Disconnect()
                System.Windows.Forms.Application.Exit()
            Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                oCompany.Disconnect()
                System.Windows.Forms.Application.Exit()

        End Select
        'SBO_Application.AppEvent = Nothing
    End Sub
End Module
