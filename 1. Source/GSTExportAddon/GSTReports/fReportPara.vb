Imports System.Threading

Public Class fReportPara
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Dim f As SAPbouiCOM.Form
    Dim rptFile As String, ArrayParaName As String, ArrayParaCaption As Array, ArrayParaType As Array
    Dim ArrayParaValue As String = ""
    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application, rptFile1 As String, ArrayParaName1 As String, ArrayParaCaption1 As String, ArrayParaType1 As String)
        SBO_Application = sbo_application1
        rptFile = rptFile1
        ArrayParaName = ArrayParaName1 'ArrayParaName1.Split(";")
        ArrayParaCaption = ArrayParaCaption1.Split(";")
        ArrayParaType = ArrayParaType1.Split(";")
        DrawForm()
    End Sub
    Private Sub DrawForm()
        Try

            Dim oItem As SAPbouiCOM.Item
            Dim oLabel As SAPbouiCOM.StaticText
            Dim cp As SAPbouiCOM.FormCreationParams
            Dim oEdit As SAPbouiCOM.EditText
            Dim obt As SAPbouiCOM.Button
            Dim oCombo As SAPbouiCOM.ComboBox
            cp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            cp.FormType = "fReportPara"
            f = SBO_Application.Forms.AddEx(cp)
            'f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 290
            f.ClientHeight = 110

            ' set the form title
            f.Title = "Report Parameters"

            f.DataSources.UserDataSources.Add("frdateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("todateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("duedateds", SAPbouiCOM.BoDataType.dt_DATE)
            ' ------------------------------Point At---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oLabel = oItem.Specific
            oLabel.Caption = "Point At"

           
            If rptFile = "GST03.rpt" Then
                oItem = f.Items.Add("txtPoint", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                oItem.Left = 95
                oItem.Top = 10
                oItem.Width = 100
                oItem.DisplayDesc = True
                oItem.Enabled = False
                oEdit = oItem.Specific
                oEdit.Value = "All"

            Else
                oItem = f.Items.Add("cbPoint", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                oItem.Left = 95
                oItem.Top = 10
                oItem.Width = 100
                oItem.DisplayDesc = True
                oCombo = oItem.Specific
                oCombo.ValidValues.Add("5a", "Total Value of Standard Rated Supply")
                oCombo.ValidValues.Add("5b", "Total Output Tax")
                oCombo.ValidValues.Add("6a", "Total Value of Standard Rate and Flat Rate Acquisition")
                oCombo.ValidValues.Add("6b", "Total Input Tax")
                oCombo.ValidValues.Add("A", "All")
            End If

            ' ------------------------------From Date---------------------------------
            oItem = f.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oLabel = oItem.Specific
            oLabel.Caption = "From Date"
            oItem = f.Items.Add("txtFrDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 30
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "frdateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            ' ------------------------------To Date---------------------------------
            oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 50
            oLabel = oItem.Specific
            oLabel.Caption = "To Date"
            oItem = f.Items.Add("txtToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 50
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "todateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            ' ------------------------------Due Date---------------------------------
            oItem = f.Items.Add("lbl3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 70
            oLabel = oItem.Specific
            oLabel.Caption = "Due Date"
            oItem = f.Items.Add("txtDueDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 70
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "duedateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")

            ' -------------------Add the OK button----------------------
            oItem = f.Items.Add("btnPrint", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 90
            obt = oItem.Specific
            obt.Caption = "View"
            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 90

            f.Visible = True
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DrawForm Event: " + ex.ToString, , True)
        End Try
    End Sub

    Private Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction = False Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                End If
                Select Case pVal.FormTypeEx
                    Case "fReportPara"
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                SBO_Application = Nothing
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If pVal.ItemUID = "btnPrint" Then
                                    Dim oEdit As SAPbouiCOM.EditText
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                    Dim format As String = "yyyyMMdd"
                                    Dim oCombo As SAPbouiCOM.ComboBox
                                    ArrayParaValue = ""
                                    If rptFile = "GST03.rpt" Then
                                        'oEdit = oForm.Items.Item("txtPoint").Specific
                                        'ArrayParaValue = ArrayParaValue + oEdit.Value.ToString + ";"
                                        ArrayParaValue = "All;"
                                    Else
                                        oCombo = oForm.Items.Item("cbPoint").Specific
                                        ArrayParaValue = ArrayParaValue + oCombo.Value.ToString + ";"
                                    End If

                                    oEdit = oForm.Items.Item("txtFrDate").Specific
                                    ArrayParaValue = ArrayParaValue + DateTime.ParseExact(oEdit.Value.ToString, format, provider).ToString("yyyy-MM-dd") + ";"
                                    oEdit = oForm.Items.Item("txtToDate").Specific
                                    ArrayParaValue = ArrayParaValue + DateTime.ParseExact(oEdit.Value.ToString, format, provider).ToString("yyyy-MM-dd") + ";"
                                    oEdit = oForm.Items.Item("txtDueDate").Specific
                                    ArrayParaValue = ArrayParaValue + DateTime.ParseExact(oEdit.Value.ToString, format, provider).ToString("yyyy-MM-dd")

                                    Dim str As String = ""
                                    Dim dt As DataTable
                                    If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        str = "Call sp_SAPB1Addon_21DO ('" + DateTime.ParseExact(oEdit.Value.ToString, format, provider).ToString("yyyyMMdd") + "','N')"
                                        dt = Functions.Hana_RunQuery(str)
                                    Else
                                        str = "exec sp_SAPB1Addon_21DO '" + DateTime.ParseExact(oEdit.Value.ToString, format, provider).ToString("yyyy-MM-dd") + "','N'"
                                        dt = Functions.DoQueryReturnDT(str)
                                    End If

                                    If IsNothing(dt) Then
                                        Dim othr As ThreadStart, myThread As Thread
                                        othr = New ThreadStart(AddressOf frmShowReport)
                                        myThread = New Thread(othr)
                                        myThread.SetApartmentState(ApartmentState.STA)
                                        myThread.Start()
                                    Else
                                        If dt.Rows.Count > 0 Then
                                            If SBO_Application.MessageBox("There are Outstanding Delivery Order more than 21 days at the end of the period. " + vbCrLf + "This will result the GST 03 report to be inaccrurate. " + vbCrLf + "Do you still want to Continue?", 1, "Yes", "No") = 1 Then
                                                Dim othr As ThreadStart, myThread As Thread
                                                othr = New ThreadStart(AddressOf frmShowReport)
                                                myThread = New Thread(othr)
                                                myThread.SetApartmentState(ApartmentState.STA)
                                                myThread.Start()
                                            Else
                                                Dim fr As New fOutstandingDO(SBO_Application)
                                            End If
                                        Else
                                            Dim othr As ThreadStart, myThread As Thread
                                            othr = New ThreadStart(AddressOf frmShowReport)
                                            myThread = New Thread(othr)
                                            myThread.SetApartmentState(ApartmentState.STA)
                                            myThread.Start()
                                        End If

                                    End If
                                    
                                    
                                End If

                        End Select
                End Select
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("ItemEvent Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub frmShowReport()
        Dim opr As New oPrint
        Dim str As String
        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            If rptFile = "GST03.rpt" Then
                str = opr.SAPPrintCrystalReport_HANA(PublicVariable.SAPPass, ArrayParaName, ArrayParaValue, "GST03-HANA.rpt") ' rptFile)
            Else
                str = opr.SAPPrintCrystalReport_HANA(PublicVariable.SAPPass, ArrayParaName, ArrayParaValue, "GST03DETAIL-HANA.rpt") ' rptFile)
            End If

        Else
            str = opr.SAPPrintCrystalReport(PublicVariable.SAPPass, ArrayParaName, ArrayParaValue, rptFile)
        End If

        If str <> "" Then
            SBO_Application.SetStatusBarMessage("Show report:  " + str, , True)
        End If
    End Sub
End Class
