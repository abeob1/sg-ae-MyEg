Option Explicit On
Option Strict Off
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Data
Imports System.Threading
Imports System.Data.Odbc
Imports Sap.Data.Hana
Imports SAPbobsCOM

Public Class Add_on
    Private WithEvents SBO_Application As SAPbouiCOM.Application
#Region "Initial"
    Public Sub New()
        MyBase.New()
        Class_Init()
        AddMenuItems()


        
    End Sub
    Public Sub SetApplication()
        Dim sbogui As SAPbouiCOM.SboGuiApi
        Dim oconnection As String
        sbogui = New SAPbouiCOM.SboGuiApi
        If Environment.GetCommandLineArgs().Length = 1 Then
            oconnection = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        Else
            oconnection = Environment.GetCommandLineArgs.GetValue(1)
        End If
        Try
            sbogui.Connect(oconnection)
        Catch ex As Exception
            MsgBox("No SAP Application Running")
            End
        End Try
        SBO_Application = sbogui.GetApplication
        Dim f1 As New fAREvents(SBO_Application)
        Dim f2 As New fOutgoingPaymentEvents(SBO_Application)
        Dim f3 As New fIncomingPaymentEvents(SBO_Application)
        Dim f4 As New fReturnsEvents(SBO_Application)
        'oApp4MenuEvent = sbogui.GetApplication
        'oApp = sbogui.GetApplication
    End Sub

    Private Function SetConnectionContext() As Integer
        Dim sCookie As String
        Dim sConnectionContext As String
        Try
            PublicVariable.oCompany = New SAPbobsCOM.Company
            sCookie = PublicVariable.oCompany.GetContextCookie
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
            If PublicVariable.oCompany.Connected = True Then
                PublicVariable.oCompany.Disconnect()
            End If
            SetConnectionContext = PublicVariable.oCompany.SetSboLoginContext(sConnectionContext)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Private Function ConnectToCompany() As Integer
        ConnectToCompany = PublicVariable.oCompany.Connect

    End Function
    Private Sub Class_Init()
        SetApplication()
        If Not SetConnectionContext() = 0 Then
            SBO_Application.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If
        If Not ConnectToCompany() = 0 Then
            SBO_Application.MessageBox("Failed connecting to the company's Database")
            End ' Terminating the Add-On Application
        Else
            SBO_Application.SetStatusBarMessage("Please wait, addon is loading.......", , False)
            Dim oUserTable As SAPbobsCOM.UserTable = Nothing
            Try
                oUserTable = PublicVariable.oCompany.UserTables.Item("GSTSETUP")
                If oUserTable.GetByKey("SQLPass") Then
                    PublicVariable.SAPPass = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
                End If
            Catch ex As Exception
                SBO_Application.SetStatusBarMessage(ex.Message, , True)
            Finally
               
            End Try

            'LoadPublicParameter()
            'select * from RDOC where TypeCode='QUT1' and Template is not null'
            'Dim a As New InitData
            'a.GetCrystalReportFile("QUT10002", "D:\Report.rpt")
        End If

        SBO_Application.SetStatusBarMessage("Add-on is loaded", , False)
    End Sub
    Private Sub AddMenuItems()

        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)



        'Dim sPath As String

        'sPath = Application.StartupPath
        'sPath = sPath.Remove(sPath.Length - 3, 3)

        oMenuItem = SBO_Application.Menus.Item("43526") 'Administration - Setup - Financial
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnGSTBadDeptSetup"
        oCreationPackage.String = "GST Setup"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("43526") 'Administration - Setup - Financial
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnGSTCheckSetup"
        oCreationPackage.String = "GST Heath Check"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("1536") 'Financial module
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "mnGST"
        oCreationPackage.String = "GST"
        oCreationPackage.Enabled = True
        oCreationPackage.Position = 15
        oMenus = oMenuItem.SubMenus
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGST")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnGSTGAF"
        oCreationPackage.String = "GST File Export"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGST")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnDO"
        oCreationPackage.String = "Outstanding Delivery Order"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGST")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnPayment"
        oCreationPackage.String = "Payment Contra"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception
        End Try

        'oMenuItem = SBO_Application.Menus.Item("mnGST")
        'oMenus = oMenuItem.SubMenus
        'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        'oCreationPackage.UniqueID = "mnReverseRpt"
        'oCreationPackage.String = "Reverse Mechanism Report"
        'Try
        '    oMenus.AddEx(oCreationPackage)
        'Catch er As Exception
        'End Try

        oMenuItem = SBO_Application.Menus.Item("mnGST")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "mnGSTR"
        oCreationPackage.String = "GST Reports"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGSTR")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnGST03"
        oCreationPackage.String = "GST-03 Report"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGSTR")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnGST3Det"
        oCreationPackage.String = "GST Detail Report"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGST")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "mnGSTBadDeb"
        oCreationPackage.String = "GST Bad Debt"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGSTBadDeb")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnARGSTBadDeptRelief"
        oCreationPackage.String = "GST AR Bad Debt Relief"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try

        oMenuItem = SBO_Application.Menus.Item("mnGSTBadDeb")
        oMenus = oMenuItem.SubMenus
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = "mnAPGSTBadDeptRelief"
        oCreationPackage.String = "GST AP Bad Debt Relief"
        Try
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
        End Try
    End Sub
#End Region
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
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
    Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "mnDO"
                        Dim fr As New fOutstandingDO(SBO_Application)
                    Case "mnPayment"
                        'Dim dt As DataTable
                        'If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '    dt = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" Where ""Code""='ContraAct';")
                        'Else
                        '    dt = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] Where Code='ContraAct'")
                        'End If
                        'Dim ContraAct As String = ""
                        'If Not IsNothing(dt) Then
                        '    If dt.Rows.Count > 0 Then
                        '        ContraAct = dt.Rows(0).Item("U_Value").ToString()
                        '    End If
                        'End If
                        'If ContraAct = "" Then
                        '    SBO_Application.MessageBox("Please setup contra account in GST Setup!")
                        'Else
                        '    Dim fr As New fPaymentContra(SBO_Application)
                        'End If
                        Dim fr As New fPaymentContra(SBO_Application)

                    Case "mnReverseRpt"
                        Dim dt As DataTable = Functions.DoQueryReturnDT("Select * from [@GSTSetup] where Code='PaymentRpt'")
                        If Not IsNothing(dt) Then
                            If dt.Rows.Count > 0 Then
                                SBO_Application.ActivateMenuItem(dt.Rows(0).Item("U_Value").ToString) '"f952e13a8d2d4db3be9ad33a7b108542"
                            End If
                        End If
                    Case "mnGSTGAF"
                        Dim fr As New fExportFile(SBO_Application)
                    Case "mnGSTCheckSetup"
                        Dim fr As New fGSTHeathCheck(SBO_Application)
                    Case "mnGSTBadDeptSetup"
                        Dim fr As New fBadDeptSetup(SBO_Application)
                    Case "mnARGSTBadDeptRelief"
                        'Dim dt As DataTable
                        'If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        '    dt = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" Where ""Code""='ARInTaxCode';")
                        'Else
                        '    dt = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] Where Code='ARInTaxCode'")
                        'End If
                        'Dim ARInTaxCode As String = ""
                        'If Not IsNothing(dt) Then
                        '    If dt.Rows.Count > 0 Then
                        '        ARInTaxCode = dt.Rows(0).Item("U_Value").ToString()
                        '    End If
                        'End If
                        'If ARInTaxCode = "" Then
                        '    SBO_Application.MessageBox("Please setup AR Input Tax Code in GST Setup!")
                        'Else
                        '    Dim fr As New fARBadDeptRelief(SBO_Application)
                        'End If

                        Dim fr As New fARBadDeptRelief(SBO_Application)

                    Case "mnAPGSTBadDeptRelief"
                        Dim fr As New fAPBadDeptRelief(SBO_Application)
                    Case "mnGST3Det"
                        Dim fr As New fReportPara(SBO_Application, "GST03DETAIL.rpt", "@pointat;@fromdate;@todate;@duedate", "Point At;From Date; To Date; Due Date", "String;Date;Date;Date")
                        'Dim dt As DataTable = Functions.DoQueryReturnDT("Select * from [@GSTSetup] where Code='GSTRptDet'")
                        'If Not IsNothing(dt) Then
                        '    If dt.Rows.Count > 0 Then
                        '        SBO_Application.ActivateMenuItem(dt.Rows(0).Item("U_Value").ToString) '"f952e13a8d2d4db3be9ad33a7b108542"
                        '    End If
                        'End If

                    Case "mnGST03"
                        Dim fr As New fReportPara(SBO_Application, "GST03.rpt", "@pointat;@fromdate;@todate;@duedate", "Point At;From Date; To Date; Due Date", "String;Date;Date;Date")

                        'Dim dt As DataTable = Functions.DoQueryReturnDT("Select * from [@GSTSetup] where Code='GSTRpt'")
                        'If Not IsNothing(dt) Then
                        '    If dt.Rows.Count > 0 Then
                        '        SBO_Application.ActivateMenuItem(dt.Rows(0).Item("U_Value").ToString) '"f952e13a8d2d4db3be9ad33a7b108542"
                        '    End If
                        'End If


                End Select
            End If

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("MenuEventHandler: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub CallForm_frmReport()
        Dim frm As New frmReport
        frm.Show()
        frm.Activate()
        System.Windows.Forms.Application.Run()

        Dim ParaName As String = "Code"
        Dim ParaValue As String = "" 'txtCode.Text
        Dim opr As New oPrint
        'opr.SAPPrintCrystalReport(ParaName, ParaValue, "PrintQCForm.rpt")
    End Sub

    Private Sub LoadPublicParameter()
        Try
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" Where ""Code""='SQLPass'")
            Else
                dt = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] Where Code='SQLPass'")
            End If

            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    PublicVariable.SAPPass = dt.Rows(0).Item("U_Value").ToString
                End If
            End If

            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" Where ""Code""='IsDebug'")
            Else
                dt = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] Where Code='IsDebug'")
            End If

            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    PublicVariable.IsDebug = dt.Rows(0).Item("U_Value").ToString
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, , True)
        End Try
    End Sub
End Class
