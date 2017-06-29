Imports System.Threading

Module clMenuEvent
    Public WithEvents oApp4MenuEvent As SAPbouiCOM.Application = Nothing
    Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApp4MenuEvent.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "mnGSTGAF"
                        'Dim othr As ThreadStart, myThread As Thread
                        'othr = New ThreadStart(AddressOf CallForm_frmExportGAF)
                        'myThread = New Thread(othr)
                        'myThread.SetApartmentState(ApartmentState.STA)
                        'myThread.Start()
                        Dim fr As New fExportFile(SBO_Application)
                    Case "mnGSTBadDeptSetup"
                        Dim fr As New fBadDeptSetup(SBO_Application)
                    Case "mnGSTBadDeptRelief"
                        Dim fr As New fBadDeptRelief(SBO_Application)
                End Select
            End If

        Catch ex As Exception
            MsgBoxWrapper(ex.Message, "", True)
        End Try
    End Sub

#Region "Functions"
    Private Sub CallForm_frmExportGAF()
        Dim frm As New frmExportGAF
        frm.Show()
        frm.Activate()
        System.Windows.Forms.Application.Run()
    End Sub
#End Region
End Module
