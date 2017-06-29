Imports SAPbobsCOM

Public Class fOutstandingDO
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
            cp.FormType = "fOutstandingDO"
            'cp.ObjectType = "BADDEBT"

            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 1000
            f.ClientHeight = 440

            ' set the form title
            f.Title = "Outstanding Delivery Order"
            f.DataSources.UserDataSources.Add("frdateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("ckds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.DataTables.Add("X")

            ' ------------------------------From Date---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oLabel = oItem.Specific
            oLabel.Caption = "Date"
            oItem = f.Items.Add("txt0", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 125
            oItem.Top = 10
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "frdateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")

            ' ------------------------------Status---------------------------------
            oItem = f.Items.Add("lbl3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oItem.Width = 100
            oLabel = oItem.Specific
            oLabel.Caption = "21 Day rules Applied"
            oItem = f.Items.Add("cb1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = 125
            oItem.Top = 30
            oItem.Width = 115
            oItem.DisplayDesc = True
            Dim cb As SAPbouiCOM.ComboBox
            cb = oItem.Specific
            cb.ValidValues.Add("A", "ALL")
            cb.ValidValues.Add("Y", "YES")
            cb.ValidValues.Add("N", "NO")
            cb.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
            ' ------------------------------Search---------------------------------
            oItem = f.Items.Add("btnsi", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 50
            Dim obtn As SAPbouiCOM.Button
            obtn = oItem.Specific
            obtn.Caption = "Search"

            AddMatrixToForm(f)

            ' -------------------Add the Create baddebt button----------------------
            oItem = f.Items.Add("btn1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 420
            obtn = oItem.Specific
            obtn.Caption = "Post GL"

            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
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
                    Case "fOutstandingDO"
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
                                        SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_DeliveryNotes, "", Key)
                                    Case "JE"
                                        BubbleEvent = False
                                        Key = omt.Columns.Item("JE").Cells.Item(pVal.Row).Specific.Value()
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
                    Case "fOutstandingDO"
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "btnsi"
                                        Calculate(oForm)
                                    Case "btn1" 'create bad debt
                                        If SBO_Application.MessageBox("Do you want to create GL Posting?", 2, "Yes", "No") = 1 Then
                                            'Dim oclje As New clJE
                                            Dim ret As String

                                            ret = CreateJE21Day(oForm) 'oForm.DataSources.DataTables.Item("X")
                                            If ret <> "" Then
                                                SBO_Application.SetStatusBarMessage("CreateJE Event: " + ret, , True)
                                            Else
                                                Calculate(oForm)
                                                SBO_Application.SetStatusBarMessage("Operation successful!", , False)
                                            End If
                                        End If
                                End Select
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
            oItem.Top = 70
            oItem.Left = 10
            oItem.Width = 990
            oItem.Height = 300

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
            oColumn.TitleObject.Caption = "Delivery No."
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
            oColumn.Width = 150
            oColumn.Editable = False

            oColumn = oColumns.Add("InvDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Invoice Date"
            oColumn.Width = 90
            oColumn.Editable = False

            oColumn = oColumns.Add("Days", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Outstanding Days"
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

            oColumn = oColumns.Add("Status", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Status"
            oColumn.Width = 50
            oColumn.Editable = False

            oColumn = oColumns.Add("JE", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oColumn.TitleObject.Caption = "JE"
            oColumn.Width = 50
            oColumn.Editable = False
            olink = oColumn.ExtendedObject
            olink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_JournalPosting

            oColumn = oColumns.Add("TaxAct", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Tax Act."
            oColumn.Width = 50
            oColumn.Editable = False

            oColumn = oColumns.Add("OVatAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oColumn.TitleObject.Caption = "Outstanding Tax Amt."
            oColumn.Width = 120
            oColumn.Editable = False
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("AddMatrixToForm Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub Calculate(oForm As SAPbouiCOM.Form)
        Try
            Dim oedit As SAPbouiCOM.EditText
            oedit = oForm.Items.Item("txt0").Specific
            Dim Fromdate As String
            Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
            Dim format As String = "yyyyMMdd"
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Fromdate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
            Else
                Fromdate = oedit.Value
            End If

            Dim Status As String
            Dim cb As SAPbouiCOM.ComboBox
            cb = oForm.Items.Item("cb1").Specific
            Status = cb.Selected.Value.ToString

            oForm.Freeze(True)
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
            Dim str As String
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "Call sp_SAPB1Addon_21DO ('" + Fromdate + "','" + Status + "')"
            Else
                str = "exec sp_SAPB1Addon_21DO '" + Fromdate + "','" + Status + "'"
            End If
            oForm.DataSources.DataTables.Item("X").Clear()
            'oForm.DataSources.DataTables.Item("X").ExecuteQuery(str)
            Dim ors As Recordset
            ors = PublicVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            ors.DoQuery(str)
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
            dt1.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("U_21DayJE", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("Days", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("TaxAct", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("OutstandVatAmt", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Rows.Add(ors.RecordCount)
            Dim i As Integer = 0
            ors.MoveFirst()
            Do Until ors.EoF
                dt1.SetValue("ck", i, ors.Fields.Item("ck").Value.ToString())
                dt1.SetValue("DocNum", i, ors.Fields.Item("DocNum").Value.ToString())
                dt1.SetValue("DocEntry", i, ors.Fields.Item("DocEntry").Value.ToString())
                dt1.SetValue("CardCode", i, ors.Fields.Item("CardCode").Value.ToString())
                dt1.SetValue("NumAtCard", i, ors.Fields.Item("NumAtCard").Value.ToString())
                dt1.SetValue("CardName", i, ors.Fields.Item("CardName").Value.ToString())
                dt1.SetValue("SubTol", i, ors.Fields.Item("SubTol").Value.ToString())
                dt1.SetValue("VatSum", i, ors.Fields.Item("VatSum").Value.ToString())
                dt1.SetValue("DocTotal", i, ors.Fields.Item("DocTotal").Value.ToString())
                dt1.SetValue("Status", i, ors.Fields.Item("Status").Value.ToString())
                dt1.SetValue("DocDate", i, ors.Fields.Item("DocDate").Value.ToString().Substring(0, (ors.Fields.Item("docdate").Value.ToString().Length - 8)))
                dt1.SetValue("U_21DayJE", i, ors.Fields.Item("U_21DayJE").Value.ToString())
                dt1.SetValue("Days", i, ors.Fields.Item("Days").Value.ToString())
                dt1.SetValue("TaxAct", i, ors.Fields.Item("TaxAct").Value.ToString())
                dt1.SetValue("OutstandVatAmt", i, ors.Fields.Item("OutstandVatAmt").Value.ToString())
                i = i + 1
                ors.MoveNext()
            Loop

            oMatrix.Columns.Item("ck").DataBind.Bind("X", "ck")
            oMatrix.Columns.Item("InvNo").DataBind.Bind("X", "DocNum")
            oMatrix.Columns.Item("InvEntry").DataBind.Bind("X", "DocEntry")
            oMatrix.Columns.Item("InvDate").DataBind.Bind("X", "DocDate")
            oMatrix.Columns.Item("CardCode").DataBind.Bind("X", "CardCode")
            oMatrix.Columns.Item("RefNo").DataBind.Bind("X", "NumAtCard")
            oMatrix.Columns.Item("CardName").DataBind.Bind("X", "CardName")
            oMatrix.Columns.Item("SubTol").DataBind.Bind("X", "SubTol")
            oMatrix.Columns.Item("Tax").DataBind.Bind("X", "VatSum")
            oMatrix.Columns.Item("Status").DataBind.Bind("X", "Status")
            oMatrix.Columns.Item("JE").DataBind.Bind("X", "U_21DayJE")
            oMatrix.Columns.Item("Total").DataBind.Bind("X", "DocTotal")
            oMatrix.Columns.Item("Days").DataBind.Bind("X", "Days")
            oMatrix.Columns.Item("TaxAct").DataBind.Bind("X", "TaxAct")
            oMatrix.Columns.Item("OVatAmt").DataBind.Bind("X", "OutstandVatAmt")

            oMatrix.LoadFromDataSource()
            oForm.Freeze(False)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("Calculate Event: " + ex.ToString, , True)
        End Try
    End Sub


#Region "POSTING JE"


    Private Function BuildTableOJDT() As DataTable
        Dim dt As New DataTable("OJDT")
        dt.Columns.Add("RefDate")
        dt.Columns.Add("Memo")
        dt.Columns.Add("U_21Day")
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

    Private Function InsertIntoOJDT(dt As DataTable, RefDate As String, Is21Day As String, Remark As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("RefDate") = RefDate
        drNew("Memo") = Remark
        drNew("U_21Day") = Is21Day
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoJDT1_1side(dt As DataTable, DrAmt As Decimal, CrAmt As Decimal, Account As String, BaseEntry As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("Account") = Account
        drNew("Debit") = DrAmt
        drNew("Credit") = CrAmt
        drNew("U_InvoiceEntry") = BaseEntry
        dt.Rows.Add(drNew)

        Return dt
    End Function

    Public Function CreateJE21Day(oForm As SAPbouiCOM.Form) As String
        'Cr OUTPUT tax, DR Relief recovery
        Try
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mat").Specific
            Dim oEdit As SAPbouiCOM.EditText
            Dim ck As SAPbouiCOM.CheckBox

            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""
            Dim oDeliveryOrder As SAPbobsCOM.Documents
            oDeliveryOrder = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)

            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim ObjType As String = "30"
            oEdit = oForm.Items.Item("txt0").Specific
            dtHeader = InsertIntoOJDT(dtHeader, oEdit.Value.ToString, "Y", "Outstanding DO GST Posting")
            Dim TotalInput As Decimal = 0
            Dim TotalOutput As Decimal = 0
            Dim ContraAct As String = ""
            Dim dtsetup As DataTable

            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dtsetup = Functions.DoQueryReturnDT("Select * from ""@GSTSETUP"" where ""Code""='DOAct'")
            Else
                dtsetup = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] where code='DOAct'")
            End If
            If dtsetup.Rows.Count > 0 Then
                ContraAct = dtsetup.Rows(0).Item("U_Value").ToString
                If ContraAct = "" Then
                    Return "Please setup 21 days rules - Accrual Account "
                End If
            Else
                Return "Please setup 21 days rules - Accrual Account "
            End If
            Dim Total As Decimal = 0
            For i As Integer = 1 To oMatrix.RowCount
                ck = oMatrix.Columns.Item("ck").Cells.Item(i).Specific
                If ck.Checked = True Then
                    oEdit = oMatrix.Columns.Item("Status").Cells.Item(i).Specific
                    If oEdit.Value = "NO" Then
                        oEdit = oMatrix.Columns.Item("OVatAmt").Cells.Item(i).Specific

                        Dim VatAmt As Decimal = 0
                        VatAmt = oEdit.Value
                        Total = Total + VatAmt

                        Dim VatAct As String = ""
                        oEdit = oMatrix.Columns.Item("TaxAct").Cells.Item(i).Specific
                        VatAct = oEdit.Value

                        Dim DocNum As String = ""
                        oEdit = oMatrix.Columns.Item("InvNo").Cells.Item(i).Specific
                        DocNum = oEdit.Value

                        dtLine = InsertIntoJDT1_1side(dtLine, 0, VatAmt, VatAct, DocNum)
                    End If
                End If

            Next
            If dtLine.Rows.Count = 0 Then
                Return "No open DO is selected!"
            End If
            dtLine = InsertIntoJDT1_1side(dtLine, Total, 0, ContraAct, "")

            dtHeader.TableName = "OJDT"
            dtLine.TableName = "JDT1"

            ds.Tables.Add(dtHeader.Copy)
            ds.Tables.Add(dtLine.Copy)
            xmlstr = oXML.ToXMLStringFromDS(ObjType, ds)
            PublicVariable.oCompany.StartTransaction()
            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)
            If ret = "" Then
                Dim JENo As Integer = PublicVariable.oCompany.GetNewObjectKey
                For i As Integer = 0 To dtLine.Rows.Count - 1
                    If dtLine.Rows(i).Item("U_InvoiceEntry").ToString <> "" Then
                        If oDeliveryOrder.GetByKey(ReturnInvEntryFromInvNum_DO(dtLine.Rows(i).Item("U_InvoiceEntry"))) Then
                            oDeliveryOrder.UserFields.Fields.Item("U_21Day").Value = "Y"
                            oDeliveryOrder.UserFields.Fields.Item("U_21DayJE").Value = CStr(JENo)
                            ret = oDeliveryOrder.Update()
                            If ret <> "0" Then
                                ret = PublicVariable.oCompany.GetLastErrorDescription
                                If PublicVariable.oCompany.InTransaction Then
                                    PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                            Else
                                ret = ""
                            End If
                        End If
                    End If

                Next
            Else
                If PublicVariable.oCompany.InTransaction Then
                    PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
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

    Private Function ReturnInvEntryFromInvNum_DO(DocNum As String) As String
        Try
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.DoQueryReturnDT("select ""DocEntry"" from ""ODLN"" where ""DocNum""='" + DocNum + "'")
            Else
                dt = Functions.DoQueryReturnDT("select DocEntry from ODLN with(nolock) where DocNum='" + DocNum + "'")
            End If

            Return dt.Rows(0).Item("DocEntry").ToString
        Catch ex As Exception
            Return ""
        End Try

    End Function
#End Region
End Class
