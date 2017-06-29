Imports SAPbobsCOM

Public Class fPaymentContra
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Dim f As SAPbouiCOM.Form
    Private oDBDataSource As SAPbouiCOM.DBDataSource
    Private oUserDataSource As SAPbouiCOM.UserDataSource
    Dim dt As DataTable = Nothing
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
            cp.FormType = "fPaymentContra"
            'cp.ObjectType = "BADDEBT"

            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 1000
            f.ClientHeight = 440

            ' set the form title
            f.Title = "Payment Contra"
            f.DataSources.UserDataSources.Add("frdateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("todateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("ptdateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.DataTables.Add("X")

            ' ------------------------------From Date---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oLabel = oItem.Specific
            oLabel.Caption = "From Date"
            oItem = f.Items.Add("txt0", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 10
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "frdateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")

            ' ------------------------------To Date---------------------------------
            oItem = f.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oLabel = oItem.Specific
            oLabel.Caption = "To Date"
            oItem = f.Items.Add("txt1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 30
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "Todateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")


            ' ------------------------------Posting Date---------------------------------
            oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 50
            oLabel = oItem.Specific
            oLabel.Caption = "Posting Date"
            oItem = f.Items.Add("txt2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 50
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "ptdateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            ' ------------------------------Search---------------------------------
            oItem = f.Items.Add("btnsi", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 200
            oItem.Top = 45
            Dim obtn As SAPbouiCOM.Button
            obtn = oItem.Specific
            obtn.Caption = "Search"

            '----------------------------Add Grid------------------------
            oItem = f.Items.Add("mat", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Top = 80
            oItem.Left = 10
            oItem.Width = 990
            oItem.Height = 340

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
                    Case "fPaymentContra"
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                SBO_Application = Nothing
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                Dim omt As SAPbouiCOM.Grid = oForm.Items.Item("mat").Specific
                                Dim Key As Integer
                                Select Case pVal.ColUID
                                    Case "InvNo"
                                        BubbleEvent = False
                                        Key = omt.Columns.Item("InvEntry").Cells.Item(pVal.Row).Specific.Value()
                                        SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Invoice, "", Key)
                                    Case "TransID"
                                        BubbleEvent = False
                                        Key = omt.DataTable.GetValue("TransID", omt.GetDataTableRowIndex(pVal.Row))  ' omt.Columns.Item("TransID").Cells.Item(pVal.Row).Specific.Value()
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
                    Case "fPaymentContra"
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "btnsi"
                                        Calculate(oForm)
                                    Case "btn1" 'create bad debt
                                        If SBO_Application.MessageBox("Do you want to create GL Posting?", 2, "Yes", "No") = 1 Then
                                            Dim ret As String
                                            Dim oedit As SAPbouiCOM.EditText = oForm.Items.Item("txt2").Specific
                                            Dim PostingDate As String = oedit.Value

                                            Dim Remark As String = ""
                                            oedit = oForm.Items.Item("txt0").Specific
                                            Dim Fromdate As String = oedit.Value
                                            oedit = oForm.Items.Item("txt1").Specific
                                            Dim ToDate As String = oedit.Value
                                            Remark = "Payment Contra " + Fromdate + " - " + ToDate
                                            ret = CreateJEPaymentContra(dt, PostingDate, Remark) 'oForm.DataSources.DataTables.Item("X")
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

    Private Sub Calculate(oForm As SAPbouiCOM.Form)
        Try
            Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
            Dim format As String = "yyyyMMdd"
            Dim oedit As SAPbouiCOM.EditText
            '-----FROM DATE---------
            oedit = oForm.Items.Item("txt0").Specific
            Dim Fromdate As String
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Fromdate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
            Else
                Fromdate = oedit.Value
            End If

            '---TO DATE--------
            oedit = oForm.Items.Item("txt1").Specific
            Dim ToDate As String
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                ToDate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
            Else
                ToDate = oedit.Value
            End If

            oForm.Freeze(True)
            oForm.DataSources.DataTables.Item("X").Clear()
            Dim str As String = ""
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "Call sp_SAPB1Addon_PaymentContra ('" + Fromdate + "','" + ToDate + "')"
                dt = Functions.Hana_RunQuery(str)
            Else
                str = "exec sp_SAPB1Addon_PaymentContra '" + Fromdate + "','" + ToDate + "'"
                dt = Functions.DoQueryReturnDT(str)
            End If
            
            If PublicVariable.IsDebug = "Y" Then
                Functions.WriteLog(str)
            End If

            If IsNothing(dt) Then
                SBO_Application.MessageBox("There's no data!!")
                oForm.Freeze(False)
                Return
            End If

            Dim dt1 As SAPbouiCOM.DataTable = f.DataSources.DataTables.Item("X")
            dt1.Columns.Add("Category", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("TotalBalance", SAPbouiCOM.BoFieldsType.ft_Float, 20)
            dt1.Columns.Add("VatGroup", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("TransID", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("Debit", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("Credit", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("Balance", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("BaseSum", SAPbouiCOM.BoFieldsType.ft_Float, 10)
            dt1.Columns.Add("VatRat", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("Memo", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("transtype", SAPbouiCOM.BoFieldsType.ft_Text, 10)
            dt1.Columns.Add("TaxAccount", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Columns.Add("ContraAccount", SAPbouiCOM.BoFieldsType.ft_Text, 20)
            dt1.Rows.Add(dt.Rows.Count)
            For i As Integer = 0 To dt.Rows.Count - 1
                dt1.SetValue("Category", i, dt.Rows(i).Item("Category").ToString())
                dt1.SetValue("TotalBalance", i, dt.Rows(i).Item("TotalBalance").ToString())
                dt1.SetValue("VatGroup", i, dt.Rows(i).Item("VatGroup").ToString())
                dt1.SetValue("TransID", i, dt.Rows(i).Item("TransID").ToString())
                dt1.SetValue("Debit", i, dt.Rows(i).Item("Debit").ToString())
                dt1.SetValue("Credit", i, dt.Rows(i).Item("Credit").ToString())
                dt1.SetValue("Balance", i, dt.Rows(i).Item("Balance").ToString())
                dt1.SetValue("BaseSum", i, dt.Rows(i).Item("BaseSum").ToString())
                dt1.SetValue("VatRat", i, dt.Rows(i).Item("VatRat").ToString())
                dt1.SetValue("Memo", i, dt.Rows(i).Item("Memo").ToString())
                dt1.SetValue("transtype", i, dt.Rows(i).Item("transtype").ToString())
                dt1.SetValue("TaxAccount", i, dt.Rows(i).Item("TaxAccount").ToString())
                dt1.SetValue("ContraAccount", i, dt.Rows(i).Item("ContraAccount").ToString())
            Next
            Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("mat").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("X")

            oGrid.CollapseLevel = 1
            oGrid.CollapseLevel = 2
            oGrid.CollapseLevel = 3
            oGrid.Rows.CollapseAll()
            'Dim ocl As SAPbouiCOM.EditTextColumn
            'ocl = oGrid.Columns.Item(6)
            'ocl.ColumnSetting.DisplayType=

            'ocl = oGrid.Columns.Item(7)
            'ocl.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            'ocl = oGrid.Columns.Item(3)
            'ocl.LinkedObjectType = "30"
            oForm.Freeze(False)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("Calculate Event: " + ex.ToString, , True)
        End Try
    End Sub
#Region "Build Table Structure"
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

#End Region
    Private Function InsertIntoOJDT(dt As DataTable, RefDate As String, IsBadDebt As String, IsContraPayment As String, Remark As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("RefDate") = RefDate
        drNew("Memo") = Remark
        drNew("U_BadDebt") = IsBadDebt
        drNew("U_ContraPayment") = IsContraPayment
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
    Public Function CreateJEPaymentContra(dt As DataTable, PostingDate As String, Remark As String) As String
        'Cr OUTPUT tax, DR Relief recovery
        Try

            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""

            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim ObjType As String = "30"
            dtHeader = InsertIntoOJDT(dtHeader, PostingDate, "", "Y", Remark)
            Dim TotalInput As Decimal = 0
            Dim TotalOutput As Decimal = 0
            Dim ContraAct As String = ""
            Dim dtsetup As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dtsetup = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" where ""Code""='ContraAct'")
            Else
                dtsetup = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] where code='ContraAct'")
            End If

            If dtsetup.Rows.Count > 0 Then
                ContraAct = dtsetup.Rows(0).Item("U_Value").ToString
            End If
            Dim dtInput As DataTable = dt.Clone
            Dim dtOutput As DataTable = dt.Clone
            If dt.Copy.Select("Category='Input'").Length > 0 Then
                dtInput = dt.Copy.Select("Category='Input'").CopyToDataTable
                TotalInput = Convert.ToDecimal(dtInput.Compute("Sum(Balance)", String.Empty))
            End If


            If dt.Copy.Select("Category='Output'").Length > 0 Then
                dtOutput = dt.Copy.Select("Category='Output'").CopyToDataTable
                TotalOutput = Convert.ToDecimal(dtOutput.Compute("Sum(Balance)", String.Empty))
            End If


            
            If TotalInput >= TotalOutput Then
                For i As Integer = 0 To dtInput.Rows.Count - 1
                    dtLine = InsertIntoJDT1_1side(dtLine, 0, dtInput.Rows(i).Item("Balance"), dtInput.Rows(i).Item("TaxAccount"), dtInput.Rows(i).Item("TransID"))
                Next
                For i As Integer = 0 To dtOutput.Rows.Count - 1
                    dtLine = InsertIntoJDT1_1side(dtLine, dtOutput.Rows(i).Item("Balance"), 0, dtOutput.Rows(i).Item("TaxAccount"), dtOutput.Rows(i).Item("TransID"))
                Next
                dtLine = InsertIntoJDT1_1side(dtLine, TotalInput - TotalOutput, 0, ContraAct, 0)
            Else
                For i As Integer = 0 To dtOutput.Rows.Count - 1
                    dtLine = InsertIntoJDT1_1side(dtLine, dtOutput.Rows(i).Item("Balance"), 0, dtOutput.Rows(i).Item("TaxAccount"), dtOutput.Rows(i).Item("TransID"))
                Next
                For i As Integer = 0 To dtInput.Rows.Count - 1
                    dtLine = InsertIntoJDT1_1side(dtLine, 0, dtInput.Rows(i).Item("Balance"), dtInput.Rows(i).Item("TaxAccount"), dtInput.Rows(i).Item("TransID"))
                Next
                dtLine = InsertIntoJDT1_1side(dtLine, 0, TotalOutput - TotalInput, ContraAct, 0)
            End If

            dtHeader.TableName = "OJDT"
            dtLine.TableName = "JDT1"

            ds.Tables.Add(dtHeader.Copy)
            ds.Tables.Add(dtLine.Copy)
            xmlstr = oXML.ToXMLStringFromDS(ObjType, ds)
            If PublicVariable.IsDebug = "Y" Then
                Functions.WriteLog(xmlstr)
            End If
            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)
            Return ret
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
End Class
