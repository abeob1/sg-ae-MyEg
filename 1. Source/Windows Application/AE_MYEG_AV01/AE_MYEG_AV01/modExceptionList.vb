Imports System.Data
Module modExceptionList

    Private oEdit As SAPbouiCOM.EditText
    Private oCheck As SAPbouiCOM.CheckBox
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oComboBox As SAPbouiCOM.ComboBox
    Private oMatrix As SAPbouiCOM.Matrix
    Private sFileName As String
    Private iRandomNo As Integer

#Region "Initialize Form"
    Private Sub InitializeExpListForm(ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "InitializeExpListForm"
        Dim sErrDesc As String = String.Empty
        Try
            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            objForm.EnableMenu("6913", False) 'User Defined windows
            objForm.EnableMenu("1290", False) 'Move First Record
            objForm.EnableMenu("1288", False) 'Move Next Record
            objForm.EnableMenu("1289", False) 'Move Previous Record
            objForm.EnableMenu("1291", False) 'Move Last Record
            objForm.EnableMenu("1281", False) 'Find Record
            objForm.EnableMenu("1282", False) 'Add New Record
            objForm.EnableMenu("1292", False) 'Add New Row

            AddUserDatasources(objForm)
            objForm.DataBrowser.BrowseBy = "5"

            oMatrix = objForm.Items.Item("10").Specific
            'oMatrix.AddRow(1)
            oMatrix.Columns.Item("V_21").Editable = False
            oMatrix.Columns.Item("V_19").Editable = False
            oMatrix.AutoResizeColumns()

            AddChooseFromList(objForm)
            CFLDataBinding(objForm)

            iRandomNo = GetRandomeCode()

            objForm.Freeze(False)
            objForm.Update()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Add DataSources to the Form"
    Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
        objForm.DataSources.UserDataSources.Add("uDtFrom", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oEdit = objForm.Items.Item("5").Specific
        oEdit.DataBind.SetBound(True, "", "uDtFrom")

        objForm.DataSources.UserDataSources.Add("uDtTo", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oEdit = objForm.Items.Item("7").Specific
        oEdit.DataBind.SetBound(True, "", "uDtTo")

        objForm.DataSources.UserDataSources.Add("uBCodeFrm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oEdit = objForm.Items.Item("12").Specific
        oEdit.DataBind.SetBound(True, "", "uBCodeFrm")

        objForm.DataSources.UserDataSources.Add("uBCodeTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oEdit = objForm.Items.Item("14").Specific
        oEdit.DataBind.SetBound(True, "", "uBCodeTo")

        oMatrix = objForm.Items.Item("10").Specific
        objForm.DataSources.UserDataSources.Add("uLineId", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
        oMatrix.Columns.Item("V_-1").DataBind.SetBound(True, "", "uLineId")

        objForm.DataSources.UserDataSources.Add("uChoose", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_15").DataBind.SetBound(True, "", "uChoose")

        objForm.DataSources.UserDataSources.Add("uActNo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_14").DataBind.SetBound(True, "", "uActNo")

        objForm.DataSources.UserDataSources.Add("uMActName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_18").DataBind.SetBound(True, "", "uMActName")

        objForm.DataSources.UserDataSources.Add("uInvRefNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
        oMatrix.Columns.Item("V_13").DataBind.SetBound(True, "", "uInvRefNo")

        objForm.DataSources.UserDataSources.Add("uDueDate", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oMatrix.Columns.Item("V_12").DataBind.SetBound(True, "", "uDueDate")

        objForm.DataSources.UserDataSources.Add("uPostDate", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oMatrix.Columns.Item("V_11").DataBind.SetBound(True, "", "uPostDate")

        objForm.DataSources.UserDataSources.Add("uCustomer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_10").DataBind.SetBound(True, "", "uCustomer")

        objForm.DataSources.UserDataSources.Add("uRemarks", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_9").DataBind.SetBound(True, "", "uRemarks")

        objForm.DataSources.UserDataSources.Add("uAmount", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oMatrix.Columns.Item("V_8").DataBind.SetBound(True, "", "uAmount")

        objForm.DataSources.UserDataSources.Add("uStatus", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_7").DataBind.SetBound(True, "", "uStatus")

        objForm.DataSources.UserDataSources.Add("uErrMsg", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_6").DataBind.SetBound(True, "", "uErrMsg")

        objForm.DataSources.UserDataSources.Add("uId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_5").DataBind.SetBound(True, "", "uId")

        objForm.DataSources.UserDataSources.Add("uStNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_4").DataBind.SetBound(True, "", "uStNo")

        objForm.DataSources.UserDataSources.Add("uPref", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_3").DataBind.SetBound(True, "", "uPref")

        objForm.DataSources.UserDataSources.Add("uTime", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_2").DataBind.SetBound(True, "", "uTime")

        objForm.DataSources.UserDataSources.Add("uSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_1").DataBind.SetBound(True, "", "uSource")

        objForm.DataSources.UserDataSources.Add("uBCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_16").DataBind.SetBound(True, "", "uBCode")

        objForm.DataSources.UserDataSources.Add("uMemo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        oMatrix.Columns.Item("V_0").DataBind.SetBound(True, "", "uMemo")

        objForm.DataSources.UserDataSources.Add("uParRcpt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_22").DataBind.SetBound(True, "", "uParRcpt")

        objForm.DataSources.UserDataSources.Add("uPartAmt", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oMatrix.Columns.Item("V_21").DataBind.SetBound(True, "", "uPartAmt")

        objForm.DataSources.UserDataSources.Add("uBalAmt", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oMatrix.Columns.Item("V_23").DataBind.SetBound(True, "", "uBalAmt")

        objForm.DataSources.UserDataSources.Add("uMultRcpt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
        oMatrix.Columns.Item("V_20").DataBind.SetBound(True, "", "uMultRcpt")

        objForm.DataSources.UserDataSources.Add("uSelMCust", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
        oMatrix.Columns.Item("V_19").DataBind.SetBound(True, "", "uSelMCust")

        objForm.DataSources.UserDataSources.Add("uPayDoc", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_17").DataBind.SetBound(True, "", "uPayDoc")

    End Sub
#End Region
#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByRef objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Customer Code
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "frozenFor"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

            'Account Code From
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "FrozenFor"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'Account Code To
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL3"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "FrozenFor"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CFLDataBinding(ByRef objForm As SAPbouiCOM.Form)
        'Customer Code
        oMatrix = objForm.Items.Item("10").Specific
        oMatrix.Columns.Item("V_10").ChooseFromListUID = "CFL1"
        oMatrix.Columns.Item("V_10").ChooseFromListAlias = "CardCode"

        'Account Code From
        oEdit = objForm.Items.Item("12").Specific
        oEdit.ChooseFromListUID = "CFL2"
        oEdit.ChooseFromListAlias = "FormatCode"

        'Account Code To
        oEdit = objForm.Items.Item("14").Specific
        oEdit.ChooseFromListUID = "CFL3"
        oEdit.ChooseFromListAlias = "FormatCode"

    End Sub
#End Region
#Region "Generate Random Number Code"
    Private Function GetRandomeCode() As String
        Dim s As String = String.Empty
        Dim iloop As Int32
        Dim random As Random
        random = New Random()
        For iloop = 0 To 7
            s = String.Concat(s, random.Next(10).ToString())
        Next iloop
        Return s
    End Function
#End Region   
#Region "Check Fields"
    Private Function CheckFields(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True

        oEdit = objForm.Items.Item("5").Specific
        If oEdit.Value = "" Then
            sErrDesc = "From Date should not be Empty"
            bCheck = False
            Return bCheck
            Exit Function
        End If

        oEdit = objForm.Items.Item("7").Specific
        If oEdit.Value = "" Then
            sErrDesc = "To Date should not be Empty"
            bCheck = False
            Return bCheck
            Exit Function
        End If

        Return bCheck
    End Function
#End Region
#Region "Load Matrix Datas"
    Private Sub LoadMatrix(ByVal objForm As SAPbouiCOM.Form)
        Dim sAcctCode As String = String.Empty
        Dim sAcctName As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sAcctCodeFrom As String = String.Empty
        Dim sAcctCodeTo As String = String.Empty
        Dim dtExecption As New DataTable
        Dim dtFromDate, dtToDate As Date
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oEdit = objForm.Items.Item("5").Specific
        dtFromDate = GetDateTimeValue(oEdit.String)
        oEdit = objForm.Items.Item("7").Specific
        dtToDate = GetDateTimeValue(oEdit.String)
        oEdit = objForm.Items.Item("12").Specific
        sAcctCodeFrom = oEdit.Value
        oEdit = objForm.Items.Item("14").Specific
        sAcctCodeTo = oEdit.Value

        Dim sQuery As String
        sQuery = "SELECT ID ,Entity ,AcctCode ,InvoiceRef ,to_char(DueDate, 'DD.MM.YY') DueDate ,Memo ,COALESCE(BalanceAmt,Amount) ""Amount"",ST_No,PaymentRef,Time,Source,BranchCode " & _
                 " FROM AB_STATEMENTUPLOAD where DueDate between '" & dtFromDate.ToString("yyyy-MM-dd") & "' and '" & dtToDate.ToString("yyyy-MM-dd") & "' AND Status = 'FAIL' " & _
                 " AND AcctCode BETWEEN '" & sAcctCodeFrom & "' AND '" & sAcctCodeTo & "' " & _
                 " AND (COALESCE(BalanceAmt,Amount) > 0) " & _
                 " UNION ALL " & _
                 " SELECT ID ,Entity ,AcctCode ,InvoiceRef ,to_char(DueDate, 'DD.MM.YY') DueDate ,Memo ,COALESCE(BalanceAmt,Amount) ""Amount"",ST_No,PaymentRef,Time,Source,BranchCode   " & _
                 " FROM AB_STATEMENTUPLOAD where DueDate between '" & dtFromDate.ToString("yyyy-MM-dd") & "' and '" & dtToDate.ToString("yyyy-MM-dd") & "' AND Status = 'SUCCESS'" & _
                 " AND AcctCode BETWEEN '" & sAcctCodeFrom & "' AND '" & sAcctCodeTo & "' " & _
                 " AND (COALESCE(BalanceAmt,Amount) > 0) "
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery()", sFuncName)
        dtExecption = ExecuteSQLQueryDataTable(sQuery, sErrDesc)

        oMatrix = objForm.Items.Item("10").Specific
        oMatrix.Clear()

        If Not dtExecption Is Nothing Then
            If dtExecption.Rows.Count >= 1 Then
                For Each oDr As DataRow In dtExecption.Rows
                    oMatrix.AddRow(1)
                    sAcctCode = oDr("AcctCode").ToString.Trim()

                    sSQL = "SELECT ""AcctName"" FROM ""OACT"" WHERE ""AcctCode"" = '" & sAcctCode & "'"
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sAcctName = oRecordSet.Fields.Item("AcctName").Value
                    End If

                    oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.value = oMatrix.RowCount
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Memo").ToString.Trim()
                    oMatrix.Columns.Item("V_16").Cells.Item(oMatrix.RowCount).Specific.value = oDr("BranchCode").ToString.Trim()
                    oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Source").ToString.Trim()
                    oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Time").ToString.Trim()
                    oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific.value = oDr("PaymentRef").ToString.Trim()
                    oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific.value = oDr("ID").ToString.Trim()
                    oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Amount").ToString.Trim()
                    oMatrix.Columns.Item("V_12").Cells.Item(oMatrix.RowCount).Specific.string = oDr("DueDate").ToString.Trim()
                    oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific.string = oDr("DueDate").ToString.Trim()
                    oMatrix.Columns.Item("V_13").Cells.Item(oMatrix.RowCount).Specific.value = oDr("InvoiceRef").ToString.Trim()
                    oMatrix.Columns.Item("V_18").Cells.Item(oMatrix.RowCount).Specific.value = sAcctName
                    oMatrix.Columns.Item("V_14").Cells.Item(oMatrix.RowCount).Specific.value = sAcctCode
                    oCheck = oMatrix.Columns.Item("V_15").Cells.Item(oMatrix.RowCount).Specific
                    oCheck.Checked = False
                Next
            End If
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
    End Sub
#End Region
#Region "Open Customer Selection Form"
    Private Sub OpenCustSelection(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer)
        Dim sSQL As String = String.Empty
        Dim iCount As Integer = 0
        Dim iTableCount As Integer = 0
        Dim sId, sInvRefNo, sCustSelDocNo, sPayDocNo As String
        Dim dAmount As Double
        Dim oDs, oDs1 As New DataSet
        sInvRefNo = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
        sId = oMatrix.Columns.Item("V_5").Cells.Item(iLine).Specific.value
        dAmount = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
        sPayDocNo = oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value

        sSQL = "SELECT COUNT(*) ""MNO"" FROM PG_TABLES WHERE UPPER(schemaname) ='PUBLIC' AND UPPER(TABLENAME) = 'AB_SELECTEDCUSTOMER'"
        oDs = ExecuteSQLQueryDataset(sSQL, sErrDesc)

        If oDs.Tables(0).Rows.Count > 0 Then
            iTableCount = oDs.Tables(0).Rows(0).Item(0).ToString

            If iTableCount = 0 Then
                sSQL = "CREATE TABLE AB_SELECTEDCUSTOMER(RANDOMNO INTEGER,DOCNUM INTEGER,ID VARCHAR(10),INVREFNO VARCHAR(50),LINE VARCHAR(10), " & _
                       " AMOUNT NUMERIC(18,3),CUSTCODE VARCHAR(50),CUSTNAME VARCHAR(100),CUSTAMT NUMERIC(18,3),PAYMENTDOCNUM VARCHAR(10),INVDOCENTRY VARCHAR(10))"
                If ExecuteSQLNonQuery(sSQL, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If
        End If

        sSQL = "SELECT COUNT(CUSTCODE),DOCNUM FROM AB_SELECTEDCUSTOMER WHERE ID = '" & sId & "' AND INVREFNO = '" & sInvRefNo & "' AND LINE = '" & iLine & "' AND RANDOMNO = '" & iRandomNo & "' "
        sSQL = sSQL & " GROUP BY DOCNUM"
        oDs1 = ExecuteSQLQueryDataset(sSQL, sErrDesc)
        If oDs1.Tables(0).Rows.Count > 0 Then
            iCount = oDs1.Tables(0).Rows(0).Item(0).ToString
            sCustSelDocNo = oDs1.Tables(0).Rows(0).Item(1).ToString
        End If
        If iCount = 0 Then
            InitializeCustSelectionForm(sId, iLine, dAmount, sInvRefNo, iRandomNo)
        ElseIf iCount > 0 Then
            CustSelectionFindForm(sCustSelDocNo, sPayDocNo)
        End If

    End Sub
#End Region
#Region "Function for Retry button"
    Private Sub RetryFunction(ByVal objForm As SAPbouiCOM.Form)
        oMatrix = objForm.Items.Item("10").Specific
        Dim dAmount As Double = 0.0
        Dim dPayAmount As Double = 0.0

        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True And oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value = "" Then
                oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "Processing..."
                If oMatrix.Columns.Item("V_11").Cells.Item(i).Specific.value = "" Then
                    oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                    oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Posting date is blank"
                    Continue For
                End If
                Try
                    dAmount = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific.value
                Catch ex As Exception
                    dAmount = 0.0
                End Try
                If dAmount = 0.0 Then
                    oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                    oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Amount column value should be greater than zero"
                    Continue For
                End If

                objForm.Items.Item("4").Enabled = False
                objForm.Items.Item("17").Enabled = False

                oCheck = oMatrix.Columns.Item("V_22").Cells.Item(i).Specific
                oCheckbox = oMatrix.Columns.Item("V_20").Cells.Item(i).Specific
                If oCheck.Checked = True And oCheckbox.Checked = True Then
                    oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                    oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Cannot select both partial receipt and multiple customer receipt checkbox"
                    Continue For
                ElseIf oCheck.Checked = True And oCheckbox.Checked = False Then
                    If oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value = "" Then
                        If oMatrix.Columns.Item("V_10").Cells.Item(i).Specific.value = "" Then
                            oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                            oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Choose the Customer"
                            Continue For
                        End If
                    End If

                    Try
                        dAmount = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific.value
                    Catch ex As Exception
                    End Try

                    Try
                        dPayAmount = oMatrix.Columns.Item("V_21").Cells.Item(i).Specific.value
                    Catch ex As Exception
                        dPayAmount = 0.0
                    End Try

                    If dPayAmount = 0.0 Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Payment amount should be greater than zero"
                        Continue For
                    End If
                    If dPayAmount > dAmount Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Payment amount should not be greater than the amount value"
                        Continue For
                    End If

                    objForm.Items.Item("3").Enabled = False
                    objForm.Items.Item("4").Enabled = False

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ArIncomingPayment_ParialReceipts()", sFuncName)

                    If ArIncomingPayment_ParialReceipts(objForm, i, sErrDesc) = False Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = sErrDesc
                    Else
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = ""
                    End If
                ElseIf oCheck.Checked = False And oCheckbox.Checked = True Then
                     If oMatrix.Columns.Item("V_19").Cells.Item(i).Specific.value = "" Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Select the list of customers"
                        Continue For
                    End If

                    objForm.Items.Item("3").Enabled = False
                    objForm.Items.Item("4").Enabled = False

                    p_oDICompany.StartTransaction()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ArIncomingPayment_MulitpleCustomers()", sFuncName)

                    If ArIncomingPayment_MulitpleCustomers(objForm, i, sErrDesc) = False Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = sErrDesc
                        If p_oDICompany.InTransaction = True Then
                            p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Else
                            p_oDICompany.StartTransaction()
                            p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Else
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = ""
                        p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If
                ElseIf oCheck.Checked = False And oCheckbox.Checked = False Then
                    If oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value = "" Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Invoice No. is blank"
                        Continue For
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ARIncoimingPayment()", sFuncName)

                    If ARIncoimingPayment(objForm, i, sErrDesc) = False Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = sErrDesc
                    Else
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = ""
                    End If
                End If
            End If
        Next

        Dim sID, sStatus, sErrorMessage, sPayDocNo, sQuery, sInvRef As String
        Dim dBalanceAmt As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                'sID = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific.value
                'sStatus = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value
                'sErrorMessage = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value
                'sPayDocNo = oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value
                'sInvRef = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value

                'If oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS" Then
                '    sQuery = "UPDATE AB_STATEMENTUPLOAD  SET InvoiceRef = '" & sInvRef & "', SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "', " & _
                '              " Status = '" & sStatus & "', ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "', " & _
                '              " PaymentDocnum = '" & sPayDocNo & "',BalanceAmt = '0' WHERE ID = '" & sID & "' "
                'Else
                '    sQuery = "UPDATE AB_STATEMENTUPLOAD SET InvoiceRef = '" & sInvRef & "' , Status = '" & sStatus & "', ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "' " & _
                '             " WHERE ID = '" & sID & "' "
                'End If
                sID = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific.value
                sStatus = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value
                sErrorMessage = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value
                sInvRef = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value
                sPayDocNo = oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value
                Try
                    dAmount = oMatrix.Columns.Item("V_8").Cells.Item(i).Specific.value
                Catch ex As Exception
                End Try
                Try
                    dPayAmount = oMatrix.Columns.Item("V_21").Cells.Item(i).Specific.value
                Catch ex As Exception
                    dPayAmount = 0.0
                End Try

                oCheck = oMatrix.Columns.Item("V_22").Cells.Item(i).Specific
                oCheckbox = oMatrix.Columns.Item("V_20").Cells.Item(i).Specific

                If oCheck.Checked = True And oCheckbox.Checked = False Then
                    dBalanceAmt = dAmount - dPayAmount
                Else
                    dBalanceAmt = 0
                End If

                If oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS" Then
                    sQuery = "UPDATE AB_STATEMENTUPLOAD  SET InvoiceRef = '" & sInvRef & "',SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',Status = '" & sStatus & "', " & _
                             " ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',PaymentDocnum = '" & sPayDocNo & "',BalanceAmt = '" & dBalanceAmt & "' " & _
                             " WHERE ID = '" & sID & "'"
                Else
                    sQuery = "UPDATE AB_STATEMENTUPLOAD SET InvoiceRef = '" & sInvRef & "' ,Status = '" & sStatus & "',ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "' " & _
                             " WHERE ID = '" & sID & "' "
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If
        Next

    End Sub
#End Region
#Region "AR incoming Payment Document"
    Private Function ARIncoimingPayment(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty
        Dim sARDocEntry As String = String.Empty
        Dim sNumAtCard As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        oARInvoice = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim dtDocDate As Date
        oMatrix = objForm.Items.Item("10").Specific
        sNumAtCard = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
        dtDocDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)

        sQuery = "SELECT ""DocNum"",""DocEntry"",""NumAtCard"" FROM OINV WHERE ""NumAtCard"" = '" & sNumAtCard & "'"
        oRecordSet.DoQuery(sQuery)
        If oRecordSet.RecordCount > 0 Then
            sARDocEntry = oRecordSet.Fields.Item("DocEntry").Value

            If oARInvoice.GetByKey(sARDocEntry) Then
                oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                oIncomingPayment.CardCode = oARInvoice.CardCode
                oIncomingPayment.DocDate = dtDocDate
                oIncomingPayment.DueDate = dtDocDate
                oIncomingPayment.TaxDate = dtDocDate

                oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

                oIncomingPayment.Invoices.DocEntry = oARInvoice.DocEntry
                oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value
                oIncomingPayment.Invoices.SumApplied = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
                oIncomingPayment.Invoices.Add()

                'Bank Transfer
                oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
                oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
                oIncomingPayment.CashSum = 0

                oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
                oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
                lRetCode = oIncomingPayment.Add()

                If lRetCode <> 0 Then
                    sErrDesc = p_oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    bCheck = False
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    p_oDICompany.GetNewObjectCode(sPayDocEntry)
                    If oIncomingPayment.GetByKey(sPayDocEntry) Then
                        sPayDocEntry = oIncomingPayment.DocNum
                    End If

                    oMatrix.Columns.Item("V_17").Editable = True
                    oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
                    objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oMatrix.Columns.Item("V_17").Editable = False

                    oARInvoice.NumAtCard = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
                    oARInvoice.Update()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    bCheck = True
                End If
            End If
        Else
            sErrDesc = "Invoice Not Found"
            Call WriteToLogFile(sErrDesc, sFuncName)
            bCheck = False
            Return bCheck
            Exit Function
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

        Return bCheck
    End Function
#End Region
#Region "Function for On Account based Button"
    Private Sub OnAccountFunction(ByVal objForm As SAPbouiCOM.Form)
        oMatrix = objForm.Items.Item("10").Specific

        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                If oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value = "" Then
                    oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "Processing..."
                    If oMatrix.Columns.Item("V_11").Cells.Item(i).Specific.value = "" Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Posting date is blank"
                        Continue For
                    End If
                    If oMatrix.Columns.Item("V_10").Cells.Item(i).Specific.value = "" Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Choose the Customer"
                        Continue For
                    End If

                    objForm.Items.Item("3").Enabled = False
                    objForm.Items.Item("17").Enabled = False

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ArIncomingPayment_onAccount()", sFuncName)

                    If ArIncomingPayment_onAccount(objForm, i, sErrDesc) = False Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = sErrDesc
                    Else
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = ""
                    End If
                End If
            End If
        Next

        Dim sID, sStatus, sErrorMessage, sPayDocNo, sInvRef, sQuery As String
        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                sID = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific.value
                sStatus = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value
                sErrorMessage = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value
                sInvRef = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value
                sPayDocNo = oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value

                If oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS" Then
                    sQuery = "UPDATE AB_STATEMENTUPLOAD  SET InvoiceRef = '" & sInvRef & "',SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',Status = '" & sStatus & "', " & _
                             " ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',PaymentDocnum = '" & sPayDocNo & "',BalanceAmt = '0' " & _
                             " WHERE ID = '" & sID & "'"
                Else
                    sQuery = "UPDATE AB_STATEMENTUPLOAD SET InvoiceRef = '" & sInvRef & "' ,Status = '" & sStatus & "',ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "' " & _
                             " WHERE ID = '" & sID & "' "
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If
        Next
    End Sub
#End Region
#Region "AR incoming payment on Account based"
    Private Function ArIncomingPayment_onAccount(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True

        Dim sFuncName As String = "ArIncomingPayment_onAccount"
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oMatrix = objForm.Items.Item("10").Specific

        oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

        oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
        oIncomingPayment.CardCode = oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value
        oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
        oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
        oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value


        oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

        ''----- Bank Transfer

        oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
        oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
        oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
        '' oIncomingPayment.CashSum = 0

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
        lRetCode = oIncomingPayment.Add()

        If lRetCode <> 0 Then
            sErrDesc = p_oDICompany.GetLastErrorDescription
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            bCheck = False
        Else
            sErrDesc = String.Empty
            p_oDICompany.GetNewObjectCode(sPayDocEntry)
            If oIncomingPayment.GetByKey(sPayDocEntry) Then
                sPayDocEntry = oIncomingPayment.DocNum
            End If

            oMatrix.Columns.Item("V_17").Editable = True
            oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
            objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oMatrix.Columns.Item("V_17").Editable = False

            Dim sQuery As String
            sQuery = "INSERT INTO AB_RECEIPTS (Entity ,receipt_no ,updated_datetime ,receipt_amount,prepaid_acct_no ,account_no ,CustomerName ,InvoiceNumber) " & _
              "VALUES ('" & p_oDICompany.CompanyDB & "', '" & sPayDocEntry & "','" & oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.value & "'," & oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value & ", " & _
              " '" & oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value & "','" & oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value & "','" & oIncomingPayment.CardName & "', '') "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
            If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            bCheck = True
        End If

        Return bCheck
    End Function
#End Region
#Region "Delete Sub Form Datas"
    Private Sub DeleteSubFormDatas(ByVal objForm As SAPbouiCOM.Form)
        Dim sSql As String = String.Empty
        Dim sId, sInvRefNo, sLine As String
        Dim oDs As New DataSet
        Dim iTableCount As Integer
       
        oMatrix = objForm.Items.Item("10").Specific

        sSql = "SELECT COUNT(*) ""MNO"" FROM PG_TABLES WHERE UPPER(schemaname) ='PUBLIC' AND UPPER(TABLENAME) = 'AB_SELECTEDCUSTOMER'"
        oDs = ExecuteSQLQueryDataset(sSql, sErrDesc)

        If oDs.Tables(0).Rows.Count > 0 Then
            iTableCount = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        If iTableCount > 0 Then
            For i As Integer = 1 To oMatrix.RowCount
                oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
                If oCheck.Checked = True And oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value = "" Then

                    sInvRefNo = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value
                    sId = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific.value
                    sLine = oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.value

                    sSql = "DELETE FROM AB_SELECTEDCUSTOMER WHERE ID = '" & sId & "' AND INVREFNO = '" & sInvRefNo & "' AND LINE = '" & sLine & "' AND RANDOMNO = '" & iRandomNo & "' "
                    If ExecuteSQLNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            Next
        End If

    End Sub
#End Region
#Region "Delete the Customer Selction"
    Private Sub DeleteCustSelectionLine(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer)
        Dim sSql As String = String.Empty
        Dim sId, sInvRefNo, sLine As String

        oMatrix = objForm.Items.Item("10").Specific
        sInvRefNo = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
        sId = oMatrix.Columns.Item("V_5").Cells.Item(iLine).Specific.value
        sLine = oMatrix.Columns.Item("V_-1").Cells.Item(iLine).Specific.value

        sSql = "DELETE FROM AB_SELECTEDCUSTOMER WHERE ID = '" & sId & "' AND INVREFNO = '" & sInvRefNo & "' AND LINE = '" & sLine & "' AND RANDOMNO = '" & iRandomNo & "' "
        If ExecuteSQLNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        oMatrix.Columns.Item("V_19").Cells.Item(iLine).Specific.value = ""
        objForm.Update()
    End Sub
#End Region
#Region "AR incoming payment on Account based for Partial Receipt"
    Private Function ArIncomingPayment_ParialReceipts(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True

        Dim sFuncName As String = "ArIncomingPayment_onAccount"
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty
        Dim sInvRefNo As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sARDocEntry As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oMatrix = objForm.Items.Item("10").Specific
        sInvRefNo = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the payment object", sFuncName)

        oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        oARInvoice = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

        If sInvRefNo = "" Then
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Ref No is empty", sFuncName)
            oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer

            oIncomingPayment.CardCode = oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value
            oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
            oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
            oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

            oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
            oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
            oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
            oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

            ''----- Bank Transfer

            oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
            oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
            oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_21").Cells.Item(iLine).Specific.value
            '' oIncomingPayment.CashSum = 0

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
            lRetCode = oIncomingPayment.Add()

            If lRetCode <> 0 Then
                sErrDesc = p_oDICompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                bCheck = False
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment added successful", sFuncName)
                sErrDesc = String.Empty
                p_oDICompany.GetNewObjectCode(sPayDocEntry)
                If oIncomingPayment.GetByKey(sPayDocEntry) Then
                    sPayDocEntry = oIncomingPayment.DocNum
                End If

                oMatrix.Columns.Item("V_17").Editable = True
                oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
                objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oMatrix.Columns.Item("V_17").Editable = False

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating receipts table", sFuncName)

                sQuery = "INSERT INTO AB_RECEIPTS (Entity ,receipt_no ,updated_datetime ,receipt_amount,prepaid_acct_no ,account_no ,CustomerName ,InvoiceNumber) " & _
                  "VALUES ('" & p_oDICompany.CompanyDB & "', '" & sPayDocEntry & "','" & oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.value & "'," & oMatrix.Columns.Item("V_21").Cells.Item(iLine).Specific.value & ", " & _
                  " '" & oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value & "','" & oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value & "','" & oIncomingPayment.CardName & "', '') "

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                bCheck = True
            End If
        Else
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting invoice for the Ref No" & sInvRefNo, sFuncName)

            sQuery = "SELECT ""DocNum"",""DocEntry"",""NumAtCard"" FROM OINV WHERE ""NumAtCard"" = '" & sInvRefNo & "'"
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sQuery)
            If oRecordSet.RecordCount > 0 Then
                sARDocEntry = oRecordSet.Fields.Item("DocEntry").Value

                If oARInvoice.GetByKey(sARDocEntry) Then
                    oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                    oIncomingPayment.CardCode = oARInvoice.CardCode
                    oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                    oIncomingPayment.DueDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                    oIncomingPayment.TaxDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)

                    oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                    oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                    oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                    oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

                    oIncomingPayment.Invoices.DocEntry = oARInvoice.DocEntry
                    oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                    'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                    'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                    'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                    'oIncomingPayment.Invoices.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value
                    oIncomingPayment.Invoices.SumApplied = oMatrix.Columns.Item("V_21").Cells.Item(iLine).Specific.value
                    oIncomingPayment.Invoices.Add()

                    'Bank Transfer
                    oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
                    oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                    oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_21").Cells.Item(iLine).Specific.value
                    oIncomingPayment.CashSum = 0

                    oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
                    oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
                    lRetCode = oIncomingPayment.Add()

                    If lRetCode <> 0 Then
                        sErrDesc = p_oDICompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                        bCheck = False
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                        p_oDICompany.GetNewObjectCode(sPayDocEntry)
                        If oIncomingPayment.GetByKey(sPayDocEntry) Then
                            sPayDocEntry = oIncomingPayment.DocNum
                        End If

                        oMatrix.Columns.Item("V_17").Editable = True
                        oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
                        objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oMatrix.Columns.Item("V_17").Editable = False

                        oARInvoice.NumAtCard = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
                        oARInvoice.Update()

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                        bCheck = True
                    End If
                End If
            Else
                sErrDesc = "Invoice Not Found"
                Call WriteToLogFile(sErrDesc, sFuncName)
                bCheck = False
                Return bCheck
                Exit Function
            End If
        End If

        Return bCheck
    End Function
#End Region
#Region "AR Incoming Payment on Account based - Multiple customers"
    Private Function ArIncomingPayment_MulitpleCustomers(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True

        Dim sFuncName As String = "ArIncomingPayment_MulitpleCustomers"
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty
        Dim sId, sInvRefNo, sCustSelDocNo, sSQL, sCardCode, sInvDocEntry As String
        Dim dCustSelAmount As Double = 0.0
        Dim oDt As New DataTable
        Dim sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset

        sInvRefNo = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
        sId = oMatrix.Columns.Item("V_5").Cells.Item(iLine).Specific.value
        sCustSelDocNo = oMatrix.Columns.Item("V_19").Cells.Item(iLine).Specific.value

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oMatrix = objForm.Items.Item("10").Specific

        sSQL = "SELECT * FROM AB_SELECTEDCUSTOMER WHERE DOCNUM = '" & sCustSelDocNo & "' AND ID = '" & sId & "' AND INVREFNO = '" & sInvRefNo & "' " & _
               " AND LINE = '" & iLine & "' AND RANDOMNO = '" & iRandomNo & "' "
        oDt = ExecuteSQLQueryDataTable(sSQL, sErrDesc)
        If Not oDt Is Nothing Then
            If oDt.Rows.Count >= 1 Then
                For Each oDr As DataRow In oDt.Rows
                    sCardCode = oDr("CUSTCODE").ToString.Trim()
                    sInvDocEntry = oDr("INVDOCENTRY").ToString.Trim()
                    dCustSelAmount = oDr("CUSTAMT").ToString.Trim()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the payment object", sFuncName)

                    oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    oARInvoice = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                    If sInvDocEntry = "" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Docentry is empty", sFuncName)

                        oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                        oIncomingPayment.CardCode = sCardCode
                        oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
                        oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
                        oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

                        oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

                        ''----- Bank Transfer

                        oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
                        oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
                        oIncomingPayment.TransferSum = dCustSelAmount
                        '' oIncomingPayment.CashSum = 0

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
                        lRetCode = oIncomingPayment.Add()

                        If lRetCode <> 0 Then
                            sErrDesc = p_oDICompany.GetLastErrorDescription
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                            bCheck = False
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment added successfully", sFuncName)

                            sErrDesc = String.Empty
                            p_oDICompany.GetNewObjectCode(sPayDocEntry)
                            If oIncomingPayment.GetByKey(sPayDocEntry) Then
                                sPayDocEntry = oIncomingPayment.DocNum
                            End If

                            oMatrix.Columns.Item("V_17").Editable = True
                            oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
                            objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oMatrix.Columns.Item("V_17").Editable = False

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating AB_RECEIPTS table", sFuncName)

                            sQuery = "INSERT INTO AB_RECEIPTS (Entity ,receipt_no ,updated_datetime ,receipt_amount,prepaid_acct_no ,account_no ,CustomerName ,InvoiceNumber) " & _
                              "VALUES ('" & p_oDICompany.CompanyDB & "', '" & sPayDocEntry & "','" & oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.value & "'," & dCustSelAmount & ", " & _
                              " '" & oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value & "','" & oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value & "','" & oIncomingPayment.CardName & "', '') "

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                            If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating AB_SELECTEDCUSTOMER table", sFuncName)

                            sQuery = "UPDATE AB_SELECTEDCUSTOMER  SET PaymentDocnum = '" & sPayDocEntry & "'  WHERE DOCNUM = '" & sCustSelDocNo & "' AND ID = '" & sId & "' " & _
                                     " AND INVREFNO = '" & sInvRefNo & "' AND LINE = '" & iLine & "' AND RANDOMNO = '" & iRandomNo & "' AND CUSTCODE = '" & sCardCode & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                            If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            bCheck = True
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing invoice Docentry " & sInvDocEntry, sFuncName)

                        sQuery = "SELECT ""DocNum"",""DocEntry"",""NumAtCard"" FROM OINV WHERE ""DocEntry"" = '" & sInvDocEntry & "'"
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery(sQuery)
                        If oRecordSet.RecordCount > 0 Then
                            sInvDocEntry = oRecordSet.Fields.Item("DocEntry").Value

                            If oARInvoice.GetByKey(sInvDocEntry) Then
                                oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                                oIncomingPayment.CardCode = oARInvoice.CardCode
                                oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                                oIncomingPayment.DueDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                                oIncomingPayment.TaxDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)

                                oIncomingPayment.Invoices.DocEntry = oARInvoice.DocEntry
                                oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                                oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                                oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                                oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
                                oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

                                oIncomingPayment.Invoices.SumApplied = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
                                oIncomingPayment.Invoices.Add()

                                'Bank Transfer
                                oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
                                oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_12").Cells.Item(iLine).Specific.string)
                                oIncomingPayment.TransferSum = dCustSelAmount 'oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
                                oIncomingPayment.CashSum = 0

                                oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
                                oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
                                lRetCode = oIncomingPayment.Add()

                                If lRetCode <> 0 Then
                                    sErrDesc = p_oDICompany.GetLastErrorDescription
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                                    bCheck = False
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                                    p_oDICompany.GetNewObjectCode(sPayDocEntry)
                                    If oIncomingPayment.GetByKey(sPayDocEntry) Then
                                        sPayDocEntry = oIncomingPayment.DocNum
                                    End If

                                    oMatrix.Columns.Item("V_17").Editable = True
                                    oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
                                    objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oMatrix.Columns.Item("V_17").Editable = False

                                    oARInvoice.NumAtCard = oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value
                                    oARInvoice.Update()

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                                    bCheck = True
                                End If
                            End If
                        End If
                    End If

                Next
            End If
        End If

        Return bCheck
    End Function
#End Region
#Region "Function for Unaccounted receipts button"
    Private Sub UnaccountedReceipts(ByVal objForm As SAPbouiCOM.Form)
        Dim dAmount As Double = 0.0
        Dim dPayAmount As Double = 0.0
        oMatrix = objForm.Items.Item("10").Specific

        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                If oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value = "" Then
                    oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "Processing..."
                    If oMatrix.Columns.Item("V_11").Cells.Item(i).Specific.value = "" Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = "Posting date is blank"
                        Continue For
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InPayment_onAccount_UnaccountedRcpts()", sFuncName)

                    If InPayment_onAccount_UnaccountedRcpts(objForm, i, sErrDesc) = False Then
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = sErrDesc
                    Else
                        oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS"
                        oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value = ""
                    End If
                End If
            End If
        Next

        Dim sID, sStatus, sErrorMessage, sPayDocNo, sInvRef, sQuery As String
        Dim dBalanceAmt As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            oCheck = oMatrix.Columns.Item("V_15").Cells.Item(i).Specific
            If oCheck.Checked = True Then
                sID = oMatrix.Columns.Item("V_5").Cells.Item(i).Specific.value
                sStatus = oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value
                sErrorMessage = oMatrix.Columns.Item("V_6").Cells.Item(i).Specific.value
                sInvRef = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value
                sPayDocNo = oMatrix.Columns.Item("V_17").Cells.Item(i).Specific.value

                If oMatrix.Columns.Item("V_7").Cells.Item(i).Specific.value = "SUCCESS" Then
                    sQuery = "UPDATE AB_STATEMENTUPLOAD  SET InvoiceRef = '" & sInvRef & "',SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',Status = '" & sStatus & "', " & _
                             " ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',PaymentDocnum = '" & sPayDocNo & "',BalanceAmt = '" & dBalanceAmt & "' " & _
                             " WHERE ID = '" & sID & "'"
                Else
                    sQuery = "UPDATE AB_STATEMENTUPLOAD SET InvoiceRef = '" & sInvRef & "' ,Status = '" & sStatus & "',ErrMsg = '" & sErrorMessage & "', LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "' " & _
                             " WHERE ID = '" & sID & "' "
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If
        Next
    End Sub
#End Region
#Region "AR incoming payment on Account based - Unaccounted receipts"
    Private Function InPayment_onAccount_UnaccountedRcpts(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True

        Dim sFuncName As String = "InPayment_onAccount_UnaccountedRcpts"
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oMatrix = objForm.Items.Item("10").Specific

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing payment object", sFuncName)
        oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

        oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
        oIncomingPayment.CardCode = p_oCompDef.sDummyCust
        oIncomingPayment.DocDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
        oIncomingPayment.Remarks = oMatrix.Columns.Item("V_0").Cells.Item(iLine).Specific.value
        oIncomingPayment.JournalRemarks = oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.value

        oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_1").Cells.Item(iLine).Specific.value
        oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_16").Cells.Item(iLine).Specific.value

        ''----- Bank Transfer

        oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value
        oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.string)
        oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
        '' oIncomingPayment.CashSum = 0

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
        lRetCode = oIncomingPayment.Add()

        If lRetCode <> 0 Then
            sErrDesc = p_oDICompany.GetLastErrorDescription
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            bCheck = False
        Else
            sErrDesc = String.Empty
            p_oDICompany.GetNewObjectCode(sPayDocEntry)
            If oIncomingPayment.GetByKey(sPayDocEntry) Then
                sPayDocEntry = oIncomingPayment.DocNum
            End If

            oMatrix.Columns.Item("V_17").Editable = True
            oMatrix.Columns.Item("V_17").Cells.Item(iLine).Specific.value = sPayDocEntry
            objForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oMatrix.Columns.Item("V_17").Editable = False

            'Dim sQuery As String
            'sQuery = "INSERT INTO AB_RECEIPTS (Entity ,receipt_no ,updated_datetime ,receipt_amount,prepaid_acct_no ,account_no ,CustomerName ,InvoiceNumber) " & _
            '  "VALUES ('" & p_oDICompany.CompanyDB & "', '" & sPayDocEntry & "','" & oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.value & "'," & oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value & ", " & _
            '  " '" & oMatrix.Columns.Item("V_10").Cells.Item(iLine).Specific.value & "','" & oMatrix.Columns.Item("V_14").Cells.Item(iLine).Specific.value & "','" & oIncomingPayment.CardName & "', '') "

            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
            'If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            bCheck = True
        End If

        Return bCheck
    End Function
#End Region

#Region "Item Event"
    Public Sub ExpList_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "ExpList_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.CharPressed = "9" Then
                            If pval.ItemUID = "10" Then
                                oMatrix = objForm.Items.Item("10").Specific
                                If pval.ColUID = "V_10" Then
                                    'CFLDataBinding(objForm)
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "8" Then
                            If CheckFields(objForm, sErrDesc) = False Then
                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            Else
                                objForm.Freeze(True)
                                p_oSBOApplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                LoadMatrix(objForm)
                                p_oSBOApplication.StatusBar.SetText("Process completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objForm.Freeze(False)
                            End If
                        ElseIf pval.ItemUID = "2" Then
                            DeleteSubFormDatas(objForm)
                        ElseIf pval.ItemUID = "3" Then 'Retry Button
                            oMatrix = objForm.Items.Item("10").Specific
                            If oMatrix.RowCount > 0 Then
                                RetryFunction(objForm)
                            End If
                        ElseIf pval.ItemUID = "4" Then 'OnAccount button
                            oMatrix = objForm.Items.Item("10").Specific
                            If oMatrix.RowCount > 0 Then
                                OnAccountFunction(objForm)
                            End If
                        ElseIf pval.ItemUID = "17" Then 'Unaccounted Receipts
                            oMatrix = objForm.Items.Item("10").Specific
                            If oMatrix.RowCount > 0 Then
                                UnaccountedReceipts(objForm)
                            End If
                        ElseIf pval.ItemUID = "18" Then 'Refresh 
                            p_oSBOApplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objForm.Items.Item("3").Enabled = True
                            objForm.Items.Item("4").Enabled = True
                            objForm.Items.Item("17").Enabled = True
                            objForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            p_oSBOApplication.StatusBar.SetText("Process completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "10" Then
                            oMatrix = objForm.Items.Item("10").Specific
                            If pval.ColUID = "V_19" Then
                                If oMatrix.Columns.Item("V_19").Cells.Item(pval.Row).Specific.value <> "" Then
                                    OpenCustSelection(objForm, pval.Row)
                                End If
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.CharPressed = "9" Then
                            If pval.ItemUID = "10" Then
                                oMatrix = objForm.Items.Item("10").Specific
                                If pval.ColUID = "V_21" Then
                                    oCheck = oMatrix.Columns.Item("V_15").Cells.Item(pval.Row).Specific
                                    oCheckbox = oMatrix.Columns.Item("V_22").Cells.Item(pval.Row).Specific
                                    If oCheck.Checked = True Then
                                        If oCheckbox.Checked = True Then
                                            Dim dAmt, dPayAmt As Double
                                            Try
                                                dAmt = oMatrix.Columns.Item("V_8").Cells.Item(pval.Row).Specific.value
                                                dPayAmt = oMatrix.Columns.Item("V_21").Cells.Item(pval.Row).Specific.value
                                            Catch ex As Exception
                                            End Try
                                            oMatrix.Columns.Item("V_23").Cells.Item(pval.Row).Specific.value = dAmt - dPayAmt
                                        Else
                                            oMatrix.Columns.Item("V_23").Cells.Item(pval.Row).Specific.value = 0.0
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "10" Then
                            If pval.Row <= oMatrix.RowCount Then
                                oMatrix = objForm.Items.Item("10").Specific
                                If pval.ColUID = "V_22" Then
                                    oCheck = oMatrix.Columns.Item("V_15").Cells.Item(pval.Row).Specific
                                    If oCheck.Checked = True Then
                                        oCheck = oMatrix.Columns.Item("V_22").Cells.Item(pval.Row).Specific
                                        If oMatrix.Columns.Item("V_17").Cells.Item(pval.Row).Specific.value = "" Then
                                            If oCheck.Checked = True Then
                                                oMatrix.Columns.Item("V_21").Editable = True
                                            Else
                                                oMatrix.Columns.Item("V_21").Editable = False
                                            End If
                                        End If
                                    End If
                                ElseIf pval.ColUID = "V_20" Then
                                    oCheck = oMatrix.Columns.Item("V_15").Cells.Item(pval.Row).Specific
                                    If oCheck.Checked = True Then
                                        oCheckbox = oMatrix.Columns.Item("V_20").Cells.Item(pval.Row).Specific
                                        oCheck = oMatrix.Columns.Item("V_22").Cells.Item(pval.Row).Specific
                                        If oCheck.Checked = True And oCheckbox.Checked = True Then
                                            p_oSBOApplication.StatusBar.SetText("Cannot Select both partial receipt and multiple customer receipt", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            oCheck = oMatrix.Columns.Item("V_20").Cells.Item(pval.Row).Specific
                                            If oCheck.Checked = True Then
                                                OpenCustSelection(objForm, oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row).Specific.value)
                                            ElseIf oCheck.Checked = False Then
                                                If oMatrix.Columns.Item("V_19").Cells.Item(pval.Row).Specific.value <> "" Then
                                                    DeleteCustSelectionLine(objForm, pval.Row)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        Try
                            Dim oItem, objItem As SAPbouiCOM.Item
                            oItem = objForm.Items.Item("10")
                            objItem = objForm.Items.Item("9")
                            objItem.Top = oItem.Top - 5
                            objItem.Height = oItem.Height + 7
                            objItem.Width = oItem.Width + 5

                        Catch ex As Exception

                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pval
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim val As String
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = objForm.ChooseFromLists.Item(sCFL_ID)
                        Try
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                If pval.ItemUID = "10" Then
                                    val = oDataTable.GetValue("CardCode", 0)
                                    oMatrix = objForm.Items.Item("10").Specific
                                    oMatrix.Columns.Item("V_10").Cells.Item(pval.Row).Specific.value = val
                                ElseIf pval.ItemUID = "12" Then
                                    val = oDataTable.GetValue("FormatCode", 0)
                                    oEdit = objForm.Items.Item("12").Specific
                                    oEdit.Value = val
                                ElseIf pval.ItemUID = "14" Then
                                    val = oDataTable.GetValue("FormatCode", 0)
                                    oEdit = objForm.Items.Item("14").Specific
                                    oEdit.Value = val
                                End If
                            End If
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objForm.Update()
                        End Try
                End Select
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Menu Event"
    Public Sub ExpList_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form
                If pVal.MenuUID = "AE_EL" Then
                    LoadFromXML("Exception List.srf", p_oSBOApplication)
                    objForm = p_oSBOApplication.Forms.Item("EXPL")
                    objForm.Visible = True
                    InitializeExpListForm(objForm)
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Menu event error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region

End Module
