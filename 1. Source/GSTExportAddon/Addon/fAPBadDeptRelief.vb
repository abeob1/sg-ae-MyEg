Imports SAPbobsCOM

Public Class fAPBadDeptRelief
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Dim f As SAPbouiCOM.Form
    Private oDBDataSource As SAPbouiCOM.DBDataSource
    Private oUserDataSource As SAPbouiCOM.UserDataSource

    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        DrawForm()
    End Sub
    Private Sub DrawForm()
        Try

            Dim oItem As SAPbouiCOM.Item
            Dim oLabel As SAPbouiCOM.StaticText
            Dim cp As SAPbouiCOM.FormCreationParams
            Dim oEdit As SAPbouiCOM.EditText
            cp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            cp.FormType = "fAPBadDeptRelief"
            'cp.ObjectType = "BADDEBT"

            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 1000
            f.ClientHeight = 440

            ' set the form title
            f.Title = "GST AP Bad Debt Relief"
            f.DataSources.UserDataSources.Add("nameds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("codeds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("dateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("date2ds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("ckds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("ck1ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.DataTables.Add("X")

            AddChooseFromList(f)

            ' ------------------------------Date---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oLabel = oItem.Specific
            oLabel.Caption = "To Date"
            oItem = f.Items.Add("txt1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 10
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "dateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            oItem = f.Items.Add("ck", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 205
            oItem.Top = 10
            oItem.Width = 100
            oItem.DisplayDesc = True
            Dim ck As SAPbouiCOM.CheckBox
            ck = oItem.Specific
            ck.Caption = "> 6 months"
            ck.DataBind.SetBound(True, "", "ckds")

            ' ------------------------------Debitbor---------------------------------
            oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oLabel = oItem.Specific
            oLabel.Caption = "Creditor"
            oItem = f.Items.Add("txt2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 30
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "codeds")
            oEdit.ChooseFromListUID = "clbp"
            oEdit.ChooseFromListAlias = "CardCode"

            oItem = f.Items.Add("txt4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 200
            oItem.Top = 30
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "nameds")

            ' ------------------------------Status---------------------------------
            oItem = f.Items.Add("lbl3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 50
            oLabel = oItem.Specific
            oLabel.Caption = "Bad Debt Status"
            oItem = f.Items.Add("cb1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = 95
            oItem.Top = 50
            oItem.Width = 100
            oItem.DisplayDesc = True
            Dim cb As SAPbouiCOM.ComboBox
            cb = oItem.Specific
            cb.ValidValues.Add("A", "ALL")
            cb.ValidValues.Add("Y", "YES")
            cb.ValidValues.Add("N", "NO")
            cb.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
            ' ------------------------------Claim Amount---------------------------------
            oItem = f.Items.Add("ck1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 10
            oItem.Top = 70
            oItem.Width = 100
            oItem.DisplayDesc = True
            ck = oItem.Specific
            ck.Caption = "Claim Amount"
            ck.DataBind.SetBound(True, "", "ck1ds")
            ck.Checked = True

            ' ------------------------------Posting Date---------------------------------
            oItem = f.Items.Add("lbl4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 90
            oLabel = oItem.Specific
            oLabel.Caption = "Posting Date"
            oItem = f.Items.Add("txt5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 90
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "date2ds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            ' ------------------------------Search---------------------------------
            oItem = f.Items.Add("btnsi", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 110
            Dim obtn As SAPbouiCOM.Button
            obtn = oItem.Specific
            obtn.Caption = "Search"

            AddMatrixToForm(f)

            ' -------------------Add the Create baddebt button----------------------
            oItem = f.Items.Add("btn1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 420
            oItem.Width = 150
            obtn = oItem.Specific
            obtn.Caption = "Create Bad Debt Relief"
            ' -------------------Add cancel bad debt button----------------------
            oItem = f.Items.Add("btn2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 170
            oItem.Top = 420
            oItem.Width = 150
            obtn = oItem.Specific
            obtn.Caption = "Cancel Bad Debt Relief"

            ' -------------------Add select all button----------------------
            oItem = f.Items.Add("btn3", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 330
            oItem.Top = 420
            obtn = oItem.Specific
            obtn.Caption = "Select All"

            ' -------------------Add unselect all button----------------------
            oItem = f.Items.Add("btn4", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 410
            oItem.Top = 420
            obtn = oItem.Specific
            obtn.Caption = "Unselect All"

            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 490
            oItem.Top = 420

            f.Visible = True
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DrawForm Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction = True Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                End If
                Select Case pVal.FormTypeEx
                    Case "fAPBadDeptRelief"
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                SBO_Application = Nothing
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                Dim omt As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
                                Dim Key As Integer
                                Select Case pVal.ColUID
                                    Case "InvNo"
                                        BubbleEvent = False
                                        Key = omt.Columns.Item("InvEntry").Cells.Item(pVal.Row).Specific.Value()
                                        SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoice, "", Key)
                                    Case "BadJENo"
                                        BubbleEvent = False
                                        Key = omt.Columns.Item("BadJE").Cells.Item(pVal.Row).Specific.Value()
                                        SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, "", Key)
                                End Select
                        End Select
                End Select

            Else

                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                End If
                Select Case pVal.FormTypeEx
                    Case "fAPBadDeptRelief"
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "btnsi"
                                        Calculate(oForm)
                                    Case "btn1" 'create bad debt
                                        If SBO_Application.MessageBox("Do you want to create bad debt for selected Invoices?", 2, "Yes", "No") = 1 Then
                                            Dim ret As String
                                            ret = CreateBadDebt(oForm)
                                            If ret <> "" Then
                                                SBO_Application.SetStatusBarMessage("CreateBadDebt Event: " + ret, , True)
                                            Else
                                                'Calculate(oForm)
                                                SBO_Application.SetStatusBarMessage("Operation successful!", , False)
                                            End If
                                        End If

                                    Case "btn2" 'cancel bad debt
                                        If SBO_Application.MessageBox("Do you want to cancel selected Bad Debt Invoices?", 2, "Yes", "No") = 1 Then
                                            Dim ret As String
                                            ret = CancelBadDebt(oForm)
                                            If ret <> "" Then
                                                SBO_Application.SetStatusBarMessage("CancelBadDebt Event: " + ret, , True)
                                            Else
                                                'Calculate(oForm)
                                                SBO_Application.SetStatusBarMessage("Operation successful!", , False)
                                            End If
                                        End If


                                    Case "btn3" 'select all
                                        SelectAll(oForm)
                                    Case "btn4" 'unselect all
                                        UnSelectAll(oForm)
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                Dim CustomerCode, CustomerName As String
                                Try
                                    CustomerCode = oDataTable.GetValue(0, 0)
                                    CustomerName = oDataTable.GetValue(1, 0)
                                Catch ex As Exception

                                End Try
                                If (pVal.ItemUID = "txt2") Then
                                    oForm.DataSources.UserDataSources.Item("codeds").ValueEx = CustomerCode
                                    oForm.DataSources.UserDataSources.Item("nameds").ValueEx = CustomerName

                                End If

                        End Select
                End Select
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("ItemEvent Event: " + ex.ToString, , True)
        End Try
    End Sub
    Sub AddMatrixToForm(ByVal f As SAPbouiCOM.Form)
        Try

            Dim oItem As SAPbouiCOM.Item
            Dim oColumns As SAPbouiCOM.Columns
            Dim oColumn As SAPbouiCOM.Column
            Dim oMatrix As SAPbouiCOM.Matrix
            Dim olink As SAPbouiCOM.LinkedButton

            oItem = f.Items.Add("mat", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
            oItem.Top = 145
            oItem.Left = 10
            oItem.Width = 990
            oItem.Height = 270

            oMatrix = oItem.Specific
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oColumns = oMatrix.Columns

            'Add Columns to the matrix
            'The # column
            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "#"
            oColumn.Width = 20
            oColumn.Editable = False

            oColumn = oColumns.Add("ck", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oColumn.TitleObject.Caption = ""
            oColumn.Width = 30

            oColumn = oColumns.Add("InvEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = ""
            oColumn.Width = 50
            oColumn.Editable = False
            oColumn.Visible = False

            oColumn = oColumns.Add("InvNo", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Invoice No."
            oColumn.Width = 80
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice

            oColumn = oColumns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Card Code"
            oColumn.Width = 70
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oColumn = oColumns.Add("CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Card Name"
            oColumn.Width = 100
            oColumn.Editable = False

            oColumn = oColumns.Add("InvDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Invoice Date"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("RefNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Cust. Ref. No."
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("SubTol", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Total before Tax"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Tax", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Total Tax"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Total", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Invoice Total"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("PaidAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Payment Amount"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Balance", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Balance Amount"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Status", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Bad Debt Relief Status"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("BadJENo", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Bad Debt Entry No"
            oColumn.Width = 90
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_JournalPosting

            oColumn = oColumns.Add("BadJE", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Bad Debt Entry"
            oColumn.Width = 90
            oColumn.Editable = False
            oColumn.Visible = False


            oColumn = oColumns.Add("DrAct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Debit Account"
            oColumn.Width = 90
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts


            oColumn = oColumns.Add("CrAct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Credit Account"
            oColumn.Width = 90
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts

            oColumn = oColumns.Add("TaxCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Tax Code"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Amount", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "Amount"
            oColumn.Width = 90
            oColumn.Editable = False


        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("AddMatrixToForm Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub Calculate(oForm As SAPbouiCOM.Form)
        Try
            Dim oedit As SAPbouiCOM.EditText
            Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
            Dim format As String = "yyyyMMdd"
            oedit = oForm.Items.Item("txt1").Specific
            Dim todate As String = oedit.Value
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                todate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
            Else
                todate = oedit.Value
            End If

            oedit = oForm.Items.Item("txt2").Specific
            Dim debitor As String = oedit.Value
            Dim ck As SAPbouiCOM.CheckBox
            ck = oForm.Items.Item("ck").Specific
            Dim month As Integer = 0
            If ck.Checked Then
                month = 6
            Else
                month = 0
            End If

            Dim Status As String
            Dim cb As SAPbouiCOM.ComboBox
            cb = oForm.Items.Item("cb1").Specific
            Status = cb.Selected.Value.ToString

            Dim ClaimAmt As String = "N"
            ck = oForm.Items.Item("ck1").Specific
            If ck.Checked Then ClaimAmt = "Y"

            oForm.Freeze(True)
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
            oForm.DataSources.DataTables.Item("X").Clear()
            Dim str As String = ""
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "Call sp_SAPB1Addon_GSTBadDebt_AP ('" + todate + "','" + debitor + "'," + CStr(month) + ",'" + ClaimAmt + "','" + Status + "')"
                dt = Functions.Hana_RunQuery(str)
            Else
                str = "exec sp_SAPB1Addon_GSTBadDebt_AP '" + todate + "','" + debitor + "'," + CStr(month) + ",'" + ClaimAmt + "','" + Status + "'"
                dt = Functions.DoQueryReturnDT(str)
            End If
            'oForm.DataSources.DataTables.Item("X").ExecuteQuery(str)
            'Dim ors As Recordset
            'ors = PublicVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            'ors.DoQuery(str)
            Dim dt1 As SAPbouiCOM.DataTable = f.DataSources.DataTables.Item("X")
            dt1.Columns.Add("ck", SAPbouiCOM.BoFieldsType.ft_Text, 1)
            dt1.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
            dt1.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 100)
            dt1.Columns.Add("NumAtCard", SAPbouiCOM.BoFieldsType.ft_Text, 100)
            dt1.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 100)
            dt1.Columns.Add("SubTol", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("VatSum", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("Balance", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("paidsum", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("docdate", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("U_BadDebtJE", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("BadDebtJENo", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("TaxCode", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("CrAct", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("DrAct", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Rows.Add(dt.Rows.Count)
            For i As Integer = 0 To dt.Rows.Count - 1
                dt1.SetValue("ck", i, dt.Rows(i).Item("ck").ToString())
                dt1.SetValue("DocNum", i, dt.Rows(i).Item("DocNum").ToString())
                dt1.SetValue("DocEntry", i, dt.Rows(i).Item("DocEntry").ToString())
                dt1.SetValue("CardCode", i, dt.Rows(i).Item("CardCode").ToString())
                dt1.SetValue("NumAtCard", i, dt.Rows(i).Item("NumAtCard").ToString())
                dt1.SetValue("CardName", i, dt.Rows(i).Item("CardName").ToString())
                dt1.SetValue("SubTol", i, dt.Rows(i).Item("SubTol").ToString())
                dt1.SetValue("VatSum", i, dt.Rows(i).Item("VatSum").ToString())
                dt1.SetValue("DocTotal", i, dt.Rows(i).Item("DocTotal").ToString())
                dt1.SetValue("Balance", i, dt.Rows(i).Item("Balance").ToString())
                dt1.SetValue("paidsum", i, dt.Rows(i).Item("paidsum").ToString())
                dt1.SetValue("Status", i, dt.Rows(i).Item("Status").ToString())
                dt1.SetValue("docdate", i, dt.Rows(i).Item("docdate").ToString().Substring(0, (dt.Rows(i).Item("docdate").ToString().Length - 8)))
                dt1.SetValue("U_BadDebtJE", i, dt.Rows(i).Item("U_BadDebtJE").ToString())
                dt1.SetValue("BadDebtJENo", i, dt.Rows(i).Item("BadDebtJENo").ToString())
                dt1.SetValue("TaxCode", i, dt.Rows(i).Item("TaxCode").ToString())
                dt1.SetValue("CrAct", i, dt.Rows(i).Item("CrAct").ToString())
                dt1.SetValue("DrAct", i, dt.Rows(i).Item("DrAct").ToString())
                dt1.SetValue("Amount", i, dt.Rows(i).Item("Amount").ToString())
                
            Next

            oMatrix.Columns.Item("ck").DataBind.Bind("X", "ck")
            oMatrix.Columns.Item("InvNo").DataBind.Bind("X", "DocNum")
            oMatrix.Columns.Item("InvEntry").DataBind.Bind("X", "DocEntry")
            oMatrix.Columns.Item("CardCode").DataBind.Bind("X", "CardCode")
            oMatrix.Columns.Item("RefNo").DataBind.Bind("X", "NumAtCard")
            oMatrix.Columns.Item("CardName").DataBind.Bind("X", "CardName")
            oMatrix.Columns.Item("SubTol").DataBind.Bind("X", "SubTol")
            oMatrix.Columns.Item("Tax").DataBind.Bind("X", "VatSum")
            oMatrix.Columns.Item("Total").DataBind.Bind("X", "DocTotal")
            oMatrix.Columns.Item("Balance").DataBind.Bind("X", "Balance")
            oMatrix.Columns.Item("PaidAmt").DataBind.Bind("X", "paidsum")
            oMatrix.Columns.Item("Status").DataBind.Bind("X", "Status")
            oMatrix.Columns.Item("InvDate").DataBind.Bind("X", "docdate")
            oMatrix.Columns.Item("BadJE").DataBind.Bind("X", "U_BadDebtJE")
            oMatrix.Columns.Item("BadJENo").DataBind.Bind("X", "BadDebtJENo")
            oMatrix.Columns.Item("TaxCode").DataBind.Bind("X", "TaxCode")
            oMatrix.Columns.Item("CrAct").DataBind.Bind("X", "CrAct")
            oMatrix.Columns.Item("DrAct").DataBind.Bind("X", "DrAct")
            oMatrix.Columns.Item("Amount").DataBind.Bind("X", "Amount")
            oMatrix.LoadFromDataSource()
            oForm.Freeze(False)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("Calculate Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub AddChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'CFL for BP
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 2
            oCFLCreationParams.UniqueID = "clbp"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("AddChooseFromList Event: " + ex.ToString, , True)
        Finally
            System.GC.Collect() 'Release the handle to the table
        End Try
    End Sub
    Private Sub SelectAll(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        For i = 0 To oForm.DataSources.DataTables.Item("X").Rows.Count - 1
            oForm.DataSources.DataTables.Item("X").SetValue("ck", i, "Y")
        Next
        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
        oMatrix.LoadFromDataSource()
        oForm.Freeze(False)
    End Sub
    Private Sub UnSelectAll(oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        For i = 0 To oForm.DataSources.DataTables.Item("X").Rows.Count - 1
            oForm.DataSources.DataTables.Item("X").SetValue("ck", i, "N")
        Next
        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
        oMatrix.LoadFromDataSource()
        oForm.Freeze(False)
    End Sub


#Region "Post JE"
    Private Function BuildTableOJDT() As DataTable
        Dim dt As New DataTable("OJDT")
        dt.Columns.Add("RefDate")
        dt.Columns.Add("Memo")
        dt.Columns.Add("U_BadDebt")
        dt.Columns.Add("U_ContraPayment")
        Return dt
    End Function
    Private Function BuildTableJDT1() As DataTable
        Dim dt As New DataTable("JDT1")
        dt.Columns.Add("Account")
        'dt.Columns.Add("ShortName")
        dt.Columns.Add("Debit")
        dt.Columns.Add("Credit")
        dt.Columns.Add("VatGroup")
        dt.Columns.Add("U_InvoiceEntry")
        dt.Columns.Add("Ref3Line")
        Return dt
    End Function
    Private Function InsertIntoOJDT(dt As DataTable, RefDate As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("RefDate") = RefDate
        drNew("Memo") = "Bad Debt Entry"
        drNew("U_BadDebt") = "Y"
        'drNew("U_ContraPayment") = IsContraPayment
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoJDT1(dt As DataTable, DebAct As String, CreAct As String, Amount As String, TaxCode As String, InvoiceEntry As String, Optional CrTax As Boolean = False, Optional Ref3Line As String = "") As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("Account") = DebAct
        'drNew("ShortName") = DebAct
        drNew("Debit") = Amount
        drNew("Credit") = 0
        drNew("Ref3Line") = Ref3Line
        If Not CrTax Then
            drNew("VatGroup") = TaxCode
        End If

        drNew("U_InvoiceEntry") = InvoiceEntry
        dt.Rows.Add(drNew)

        drNew = dt.NewRow
        drNew("Account") = CreAct
        'drNew("ShortName") = CreAct
        drNew("Debit") = 0
        drNew("Credit") = Amount
        drNew("U_InvoiceEntry") = InvoiceEntry
        drNew("Ref3Line") = Ref3Line
        If CrTax Then
            drNew("VatGroup") = TaxCode
        End If
        dt.Rows.Add(drNew)

        Return dt
    End Function
    Public Function CreateBadDebt(oForm As SAPbouiCOM.Form) As String
        Try
            Dim OneJE As String = "N"
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.DoQueryReturnDT("select ""U_Value"" from ""@GSTSETUP"" where ""Code""='OneJE'")
            Else
                dt = Functions.DoQueryReturnDT("select U_Value from [@GSTSETUP] where code='OneJE'")
            End If
            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    OneJE = dt.Rows(0).Item("U_Value").ToString
                End If
            End If
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
            Dim ck As SAPbouiCOM.CheckBox
            Dim oEdit As SAPbouiCOM.EditText
            Dim Ret As String, InvoiceEntry As Integer, RefDate As String, DebAct As String, CreAct As String, Amount As String, TaxCode As String
            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()

            For i As Integer = 1 To oMatrix.RowCount
                ck = oMatrix.Columns.Item("ck").Cells.Item(i).Specific
                If ck.Checked = True Then
                    oEdit = oMatrix.Columns.Item("Status").Cells.Item(i).Specific
                    If oEdit.Value = "NO" Then
                        oEdit = oMatrix.Columns.Item("InvNo").Cells.Item(i).Specific
                        InvoiceEntry = oEdit.Value
                        oEdit = oForm.Items.Item("txt5").Specific
                        RefDate = oEdit.Value
                        oEdit = oMatrix.Columns.Item("DrAct").Cells.Item(i).Specific
                        DebAct = oEdit.Value
                        oEdit = oMatrix.Columns.Item("CrAct").Cells.Item(i).Specific
                        CreAct = oEdit.Value
                        oEdit = oMatrix.Columns.Item("Amount").Cells.Item(i).Specific
                        Amount = oEdit.Value
                        oEdit = oMatrix.Columns.Item("TaxCode").Cells.Item(i).Specific
                        TaxCode = oEdit.Value
                        If Amount <> 0 Then
                            If OneJE = "N" Then
                                Ret = CreateJE(InvoiceEntry, RefDate, DebAct, CreAct, Amount, TaxCode)
                                If Ret <> "" Then
                                    Return Ret
                                End If
                            Else
                                dtHeader = InsertIntoOJDT(dtHeader, RefDate)
                                dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceEntry)
                            End If

                        End If
                    End If
                End If
            Next
            If OneJE = "Y" Then
                If dtHeader.Rows.Count > 0 Then
                    Ret = CreateOneJE(dtHeader, dtLine)
                    Return Ret
                End If
            End If
            oMatrix.Clear()
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Private Function CreateOneJE(dtHeader As DataTable, dtLine As DataTable) As String
        Try
            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""
            Dim dtheader2 As DataTable = dtHeader.Clone
            dtheader2.ImportRow(dtHeader.Rows(0))
            ds.Tables.Add(dtheader2.Copy)
            ds.Tables.Add(dtLine.Copy)

            xmlstr = oXML.ToXMLStringFromDS("30", ds)
            Dim oinvoice As SAPbobsCOM.Documents
            oinvoice = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            PublicVariable.oCompany.StartTransaction()
            ret = oXML.CreateMarketingDocument(xmlstr, "30")
            If ret = "" Then
                Dim JENo As Integer = PublicVariable.oCompany.GetNewObjectKey
                For i As Integer = 0 To dtLine.Rows.Count - 1

                    If oinvoice.GetByKey(ReturnInvEntryFromInvNum_AP(dtLine.Rows(i).Item("U_InvoiceEntry").ToString)) Then
                        oinvoice.UserFields.Fields.Item("U_BadDebt").Value = "Y"
                        oinvoice.UserFields.Fields.Item("U_BadDebtJE").Value = CStr(JENo)
                        ret = oinvoice.Update()
                        If ret <> "0" Then
                            ret = PublicVariable.oCompany.GetLastErrorDescription
                        Else
                            ret = ""
                        End If
                    End If

                Next
            End If
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            Return ret
        Catch ex As Exception
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Return ex.ToString
        End Try
    End Function
    Private Function CreateJE(InvoiceEntry As Integer, RefDate As String, DebAct As String, CreAct As String, Amount As String, TaxCode As String) As String
        Try
            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""

            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim oinvoice As SAPbobsCOM.Documents
            oinvoice = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            Dim ObjType As String = "30"
            dtHeader = InsertIntoOJDT(dtHeader, RefDate)
            dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceEntry)
            dtHeader.TableName = "OJDT"
            dtLine.TableName = "JDT1"

            ds.Tables.Add(dtHeader.Copy)
            ds.Tables.Add(dtLine.Copy)
            xmlstr = oXML.ToXMLStringFromDS(ObjType, ds)
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            PublicVariable.oCompany.StartTransaction()
            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)
            If ret = "" Then
                Dim JENo As Integer = PublicVariable.oCompany.GetNewObjectKey

                If oinvoice.GetByKey(ReturnInvEntryFromInvNum_AP(InvoiceEntry)) Then
                    oinvoice.UserFields.Fields.Item("U_BadDebt").Value = "Y"
                    oinvoice.UserFields.Fields.Item("U_BadDebtJE").Value = CStr(JENo)
                    ret = oinvoice.Update()
                    If ret <> "0" Then
                        ret = PublicVariable.oCompany.GetLastErrorDescription
                    Else
                        ret = ""
                    End If
                End If

            End If
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

            Return ret
        Catch ex As Exception
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Return ex.ToString
        End Try
    End Function
    Private Function ReturnInvEntryFromInvNum_AP(DocNum As String) As String
        Try
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.DoQueryReturnDT("select ""DocEntry"" from ""OPCH"" where ""DocNum""='" + DocNum + "'")
            Else
                dt = Functions.DoQueryReturnDT("select DocEntry from OPCH with(nolock) where DocNum='" + DocNum + "'")
            End If

            Return dt.Rows(0).Item("DocEntry").ToString
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Public Function CancelBadDebt(oForm As SAPbouiCOM.Form) As String
        Try

            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
            Dim ck As SAPbouiCOM.CheckBox
            Dim oEdit As SAPbouiCOM.EditText
            Dim Ret As String, JENo As String, InvoiceEntry As Integer, RefDate As String, DebAct As String, CreAct As String, Amount As String, TaxCode As String
            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()

            For i As Integer = 1 To oMatrix.RowCount
                ck = oMatrix.Columns.Item("ck").Cells.Item(i).Specific
                If ck.Checked = True Then
                    oEdit = oMatrix.Columns.Item("InvNo").Cells.Item(i).Specific
                    InvoiceEntry = oEdit.Value
                    oEdit = oForm.Items.Item("txt5").Specific
                    RefDate = oEdit.Value
                    oEdit = oMatrix.Columns.Item("DrAct").Cells.Item(i).Specific
                    DebAct = oEdit.Value
                    oEdit = oMatrix.Columns.Item("CrAct").Cells.Item(i).Specific
                    CreAct = oEdit.Value
                    oEdit = oMatrix.Columns.Item("Amount").Cells.Item(i).Specific
                    Amount = oEdit.Value
                    oEdit = oMatrix.Columns.Item("TaxCode").Cells.Item(i).Specific
                    TaxCode = oEdit.Value
                    oEdit = oMatrix.Columns.Item("BadJE").Cells.Item(i).Specific
                    JENo = oEdit.Value
                    If JENo <> "" Then
                        Ret = CancelJE(JENo)
                        If Ret = "" Then
                            UpdateNonBadDebtInvoice(JENo)
                        Else
                            Return Ret
                        End If

                    End If
                End If
            Next
            oMatrix.Clear()
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
    Private Function UpdateNonBadDebtInvoice(JENo As String) As String
        Try
            Dim str As String
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "update ""OPCH"" set ""U_BadDebt""='N' where ""DocNum"" in (select distinct ""U_InvoiceEntry"" from ""JDT1"" where ""TransId""=" + JENo + ")"
            Else
                str = "update OPCH set U_BadDebt='N' where Docnum in (select distinct U_InvoiceEntry from JDT1 where TransID=" + JENo + ")"
            End If

            Functions.DoQueryReturnDT(str)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Private Function CancelJE(JENo As String) As String
        Dim oJE As SAPbobsCOM.JournalEntries = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        If oJE.GetByKey(JENo) Then
            If oJE.Cancel = 0 Then
                Return ""
            Else
                Return PublicVariable.oCompany.GetLastErrorDescription
            End If
        End If
        Return ""
    End Function
#End Region
End Class
