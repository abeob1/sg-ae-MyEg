﻿Imports System.IO
Imports System.Data

Module modUploadStatement

    Private oEdit As SAPbouiCOM.EditText
    Private oComboBox As SAPbouiCOM.ComboBox
    Private oMatrix As SAPbouiCOM.Matrix
    Private oRecordSet As SAPbobsCOM.Recordset
    Private sFileName As String

#Region "Initialize Form"
    Private Sub InitializeUploadForm(ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "InitializeForm"
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
            objForm.DataBrowser.BrowseBy = "6"

            oMatrix = objForm.Items.Item("7").Specific
            oMatrix.AddRow(1)
            oMatrix.AutoResizeColumns()

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
#Region "Add Data Sources to form"
    Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
        objForm.DataSources.UserDataSources.Add("uFileUplod", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit = objForm.Items.Item("6").Specific
        oEdit.DataBind.SetBound(True, "", "uFileUplod")

        oMatrix = objForm.Items.Item("7").Specific
        objForm.DataSources.UserDataSources.Add("uLineId", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
        oMatrix.Columns.Item("V_-1").DataBind.SetBound(True, "", "uLineId")

        objForm.DataSources.UserDataSources.Add("uMActNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_5").DataBind.SetBound(True, "", "uMActNo")

        objForm.DataSources.UserDataSources.Add("uMActName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_6").DataBind.SetBound(True, "", "uMActName")

        objForm.DataSources.UserDataSources.Add("uInvRefNo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_4").DataBind.SetBound(True, "", "uInvRefNo")

        objForm.DataSources.UserDataSources.Add("uPostDate", SAPbouiCOM.BoDataType.dt_DATE, 50)
        oMatrix.Columns.Item("V_3").DataBind.SetBound(True, "", "uPostDate")

        objForm.DataSources.UserDataSources.Add("uAmount", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oMatrix.Columns.Item("V_2").DataBind.SetBound(True, "", "uAmount")

        objForm.DataSources.UserDataSources.Add("uStatus", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_1").DataBind.SetBound(True, "", "uStatus")

        objForm.DataSources.UserDataSources.Add("uErrMsg", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_0").DataBind.SetBound(True, "", "uErrMsg")

        objForm.DataSources.UserDataSources.Add("uID", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_10").DataBind.SetBound(True, "", "uID")

        objForm.DataSources.UserDataSources.Add("uSTNO", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_7").DataBind.SetBound(True, "", "uSTNO")

        objForm.DataSources.UserDataSources.Add("uTIME", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_8").DataBind.SetBound(True, "", "uTIME")

        objForm.DataSources.UserDataSources.Add("uSOURCE", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_9").DataBind.SetBound(True, "", "uSOURCE")

        objForm.DataSources.UserDataSources.Add("uBRANCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_11").DataBind.SetBound(True, "", "uBRANCH")

        objForm.DataSources.UserDataSources.Add("uPayDoc", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oMatrix.Columns.Item("V_13").DataBind.SetBound(True, "", "uPayDoc")

    End Sub
#End Region
#Region "Show OpenFileDialog"
    Private Sub showOpenFileDialog(ByRef objForm As SAPbouiCOM.Form)
        Dim myThread As New System.Threading.Thread(AddressOf OpenFileDialog)
        myThread.SetApartmentState(Threading.ApartmentState.STA)
        myThread.Start()
        myThread.Join()
    End Sub

    Private Sub OpenFileDialog()
        Dim DummyForm As New frmOpenFileDialog
        sFileName = ""
        DummyForm.Show()
        DummyForm.OpenFileDialog1.ShowDialog()
        sFileName = DummyForm.OpenFileDialog1.FileName
        System.Threading.Thread.CurrentThread.Abort()
    End Sub
#End Region
#Region "Check Fields before pressing upload button"
    Private Function CheckBeforeUpload(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim v_Check As Boolean
        v_Check = True
        sErrDesc = ""

        oEdit = objForm.Items.Item("6").Specific
        If oEdit.Value = "" Then
            sErrDesc = "Choose the Bank statement file for upload.."
            v_Check = False
            Return v_Check
            Exit Function
        Else
            sFileName = oEdit.Value
        End If

        If Not File.Exists(oEdit.Value) Then
            sErrDesc = "Invalid File Path"
            v_Check = False
            Return v_Check
            Exit Function
        End If

        Dim sSQL As String
        Dim oDs As DataSet
        sSQL = "SELECT ID FROM AB_STATEMENTUPLOAD WHERE FILENAME = '" & sFileName & "'"
        oDs = ExecuteSQLQueryDataset(sSQL, sErrDesc)
        If oDs.Tables(0).Rows.Count > 0 Then
            v_Check = False
            sErrDesc = "File already uploaded"
            Return v_Check
            Exit Function
        End If

        Dim oDT_Bankstat As DataTable
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetDataTableFromExcel()", sFuncName)
        oDT_Bankstat = GetDataSetFromExcel(sFileName, sErrDesc)

        Dim newColumn As New Data.DataColumn("FileName", GetType(System.String))
        newColumn.DefaultValue = sFileName
        oDT_Bankstat.Columns.Add(newColumn)

        For Each Datarows As DataRow In oDT_Bankstat.Rows
            Dim dt_tmp As DateTime
            ' dDate = CDate(Datarows(3).ToString.Trim())

            Dim sDate As String = Datarows(3).ToString.Trim()
            Dim dDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "M/d/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDate)
            
            'dt_tmp = Datarows(7)

            'sSQL = "INSERT INTO AB_STATEMENTUPLOAD (Entity ,AcctCode ,InvoiceRef ,DueDate,Memo ,Amount,PaymentRef ,Time ,Source ,BranchCode,TransactionCode, FileName) " & _
            '          "VALUES ('', '" & Datarows(1).ToString.Trim() & "','" & Datarows(2).ToString.Trim() & "','" & Replace(Datarows(3).ToString.Trim(), "'", "''") & "','" & Datarows(4).ToString.Trim() & "', " & _
            '          " " & Datarows(5).ToString.Trim() & ",'" & Datarows(6).ToString.Trim() & "','" & dt_tmp.ToString("HH:mm:ss") & "','" & Datarows(8).ToString.Trim() & "','" & Datarows(9).ToString.Trim() & "', " & _
            '          " '" & Datarows(10).ToString.Trim() & "' , '" & Datarows(11).ToString.Trim() & "' ) "

            If IsDBNull(Datarows(7)) Then

                sSQL = "INSERT INTO AB_STATEMENTUPLOAD (Entity ,AcctCode ,InvoiceRef ,DueDate,Memo ,Amount,PaymentRef ,Time ,Source ,BranchCode,TransactionCode, FileName) " & _
                       "VALUES ('', '" & Datarows(1).ToString.Trim() & "','" & Datarows(2).ToString.Trim() & "','" & dDate.ToString("yyyy-MM-dd") & "','" & Datarows(4).ToString.Trim() & "', " & _
                       " " & Datarows(5).ToString.Trim() & ",'" & Datarows(6).ToString.Trim() & "','','" & Datarows(8).ToString.Trim() & "','" & Datarows(9).ToString.Trim() & "', " & _
                       " '" & Datarows(10).ToString.Trim() & "' , '" & Datarows(11).ToString.Trim() & "' ) "
            Else
                dt_tmp = Datarows(7)

                sSQL = "INSERT INTO AB_STATEMENTUPLOAD (Entity ,AcctCode ,InvoiceRef ,DueDate,Memo ,Amount,PaymentRef ,Time ,Source ,BranchCode,TransactionCode, FileName) " & _
                       "VALUES ('', '" & Datarows(1).ToString.Trim() & "','" & Datarows(2).ToString.Trim() & "','" & dDate.ToString("yyyy-MM-dd") & "','" & Datarows(4).ToString.Trim() & "', " & _
                       " " & Datarows(5).ToString.Trim() & ",'" & Datarows(6).ToString.Trim() & "','" & dt_tmp.ToString("HH:mm:ss") & "','" & Datarows(8).ToString.Trim() & "','" & Datarows(9).ToString.Trim() & "', " & _
                       " '" & Datarows(10).ToString.Trim() & "' , '" & Datarows(11).ToString.Trim() & "' ) "
            End If

            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()", sFuncName)
            If ExecuteSQLNonQuery(sSQL, sErrDesc) <> RTN_SUCCESS Then
                sSQL = "DELETE FROM AB_STATEMENTUPLOAD WHERE FileName = '" & sFileName & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while inserting values. Deleteing the inserted values", sFuncName)
                If ExecuteSQLNonQuery(sSQL, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Throw New ArgumentException("Error occured while inserting values in excel/File is '" & sFileName & "' ")
            End If
        Next

        sSQL = "SELECT ID ,Entity ,AcctCode ,InvoiceRef ,to_char(DueDate, 'DD.MM.YY') DueDate ,Memo ,Amount,PaymentRef,Time,Source,BranchCode,TransactionCode  " & _
              " FROM AB_STATEMENTUPLOAD WHERE COALESCE(Status,'') = '' "

        oDT_Bankstat = New DataTable
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery()", sFuncName)
        oDT_Bankstat = ExecuteSQLQueryDataTable(sSQL, sErrDesc)

        objForm.Freeze(True)
        p_oSBOApplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        LoadDatasinMatrix(objForm, oDT_Bankstat, sErrDesc)
        p_oSBOApplication.StatusBar.SetText("Process completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        objForm.Freeze(False)

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IncomingPayment_OnCustomer()", sFuncName)
        If IncomingPayment_OnCustomer(objForm, oDT_Bankstat, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        Return v_Check
    End Function
#End Region
#Region "Load Datas in Matrix"
    Private Sub LoadDatasinMatrix(ByVal objForm As SAPbouiCOM.Form, ByVal dataTable As DataTable, ByRef sErrDesc As String)
        Dim sFuncName As String = "LoadDatasinMatrix"
        Dim sSql As String = String.Empty
        Dim sAcctCode As String = String.Empty
        Dim sAcctName As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        oMatrix = objForm.Items.Item("7").Specific
        oMatrix.Clear()

        For Each oDr As DataRow In dataTable.Rows
            oMatrix.AddRow(1)
            sAcctCode = oDr("AcctCode").ToString.Trim()

            sSql = "SELECT ""AcctName"" FROM ""OACT"" WHERE ""AcctCode"" = '" & sAcctCode & "'"
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sSql)
            If oRecordSet.RecordCount > 0 Then
                sAcctName = oRecordSet.Fields.Item("AcctName").Value
            End If

            oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.value = oMatrix.RowCount
            oMatrix.Columns.Item("V_5").Cells.Item(oMatrix.RowCount).Specific.value = sAcctCode
            oMatrix.Columns.Item("V_6").Cells.Item(oMatrix.RowCount).Specific.value = sAcctName
            oMatrix.Columns.Item("V_4").Cells.Item(oMatrix.RowCount).Specific.value = oDr("InvoiceRef").ToString.Trim()
            oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific.string = oDr("DueDate").ToString.Trim()
            oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Amount").ToString.Trim()
            oMatrix.Columns.Item("V_10").Cells.Item(oMatrix.RowCount).Specific.value = oDr("ID").ToString.Trim()
            oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific.value = oDr("TransactionCode").ToString.Trim()
            oMatrix.Columns.Item("V_8").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Time").ToString.Trim()
            oMatrix.Columns.Item("V_9").Cells.Item(oMatrix.RowCount).Specific.value = oDr("Source").ToString.Trim()
            oMatrix.Columns.Item("V_11").Cells.Item(oMatrix.RowCount).Specific.value = oDr("BranchCode").ToString.Trim()
        Next

    End Sub
#End Region
#Region "Incoming Payment onCustomer"
    Public Function IncomingPayment_OnCustomer(ByVal objForm As SAPbouiCOM.Form, ByVal dataTable As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim icount As Integer = 0
        Dim oDV_Payments As DataView = dataTable.DefaultView
        Dim sQuery As String = String.Empty

        Try
            sFuncName = "IncomingPayment_OnCustomer()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oMatrix = objForm.Items.Item("7").Specific
            For i As Integer = 1 To oMatrix.RowCount
                If oMatrix.Columns.Item("V_10").Cells.Item(i).Specific.value <> "" Then
                    oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value = "Processing..."
                    If oMatrix.Columns.Item("V_4").Cells.Item(i).Specific.value <> "" Then
                        oDV_Payments.RowFilter = "ID='" & oMatrix.Columns.Item("V_10").Cells.Item(i).Specific.value & "'"

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AR_IncomingPayment()", sFuncName)
                        If AR_IncomingPayment(objForm, i, sErrDesc) = False Then
                            oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value = "FAIL"
                            oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value = sErrDesc
                        Else
                            oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value = "SUCCESS"
                        End If
                    Else
                        oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value = "FAIL"
                        oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value = "Not Matched"
                    End If
                End If
            Next

            p_oSBOApplication.StatusBar.SetText("Uploaded successfully the bank statement...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Dim sID, sStatus, sErrorMessage, sPayDocNo As String
            Dim dAmount As Double = 0.0
            For i As Integer = 1 To oMatrix.RowCount
                sID = oMatrix.Columns.Item("V_10").Cells.Item(i).Specific.value
                sStatus = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value
                sErrorMessage = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value
                sPayDocNo = oMatrix.Columns.Item("V_13").Cells.Item(i).Specific.value
                Try
                    dAmount = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific.value
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                If oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value = "SUCCESS" Then
                    sQuery = "UPDATE AB_STATEMENTUPLOAD SET UploadDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "', " & _
                         " Status = '" & sStatus & "', ErrMsg = '" & sErrorMessage & "',LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',PaymentDocnum = '" & sPayDocNo & "',BalanceAmt = '0'" & _
                         " WHERE ID = '" & sID & "' "
                Else
                    sQuery = "UPDATE AB_STATEMENTUPLOAD SET UploadDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',SAPSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "', " & _
                         " Status = '" & sStatus & "', ErrMsg = '" & sErrorMessage & "',LastSyncDate = '" & Date.Now.Date.ToString("yyyy-MM-dd") & "',PaymentDocnum = '" & sPayDocNo & "'" & _
                         " WHERE ID = '" & sID & "' "
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLNonQuery()" & sQuery, sFuncName)
                If ExecuteSQLNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Next

            IncomingPayment_OnCustomer = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            sErrDesc = ex.Message
            IncomingPayment_OnCustomer = RTN_ERROR
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function
#End Region
#Region "AR Incoming payment"
    Private Function AR_IncomingPayment(ByVal objForm As SAPbouiCOM.Form, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim bCheck As Boolean
        bCheck = True
        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments = Nothing
        Dim oARInvoice As SAPbobsCOM.Documents = Nothing
        Dim sPayDocEntry As String = String.Empty
        Dim sARDocEntry As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sNumAtCard As String = String.Empty

        Try
            sFuncName = "AR_IncomingPayment"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing payment object", sFuncName)
            oIncomingPayment = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing Invoice object", sFuncName)
            oARInvoice = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            Dim dtDocDate As Date
            oMatrix = objForm.Items.Item("7").Specific
            sNumAtCard = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
            dtDocDate = GetDateTimeValue(oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.string)

            Dim dXcelAmount As Double = 0.0
            Dim dInvoiceSum As Double = 0.0
            dXcelAmount = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value

            sQuery = "SELECT SUM(""DocTotal"") AS ""DocTotal"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            oRecordSet.DoQuery(sQuery)
            If oRecordSet.RecordCount > 0 Then
                dInvoiceSum = oRecordSet.Fields.Item("DocTotal").Value
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total invoice sum is " & Math.Round(dInvoiceSum, 2), sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total Excel amount is " & Math.Round(dXcelAmount, 2), sFuncName)

            If Math.Round(dXcelAmount, 2) <> Math.Round(dInvoiceSum, 2) Then
                sErrDesc = "Amount in Excel and Invoice total amount does not match"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sQuery = "SELECT DISTINCT ""CardCode"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            oRecordSet.DoQuery(sQuery)
            If oRecordSet.RecordCount > 1 Then
                sErrDesc = "invoice reference does not match customer name"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sQuery = "SELECT ""CardCode"",""DocNum"",""DocEntry"",""NumAtCard"" FROM ""OINV"" WHERE ""NumAtCard"" = '" & sNumAtCard & "' " & _
                     " GROUP BY ""CardCode"",""DocNum"",""DocEntry"",""NumAtCard"" "
            oRecordSet.DoQuery(sQuery)
            If Not (oRecordSet.BoF And oRecordSet.EoF) Then
                oRecordSet.MoveFirst()
                Do Until oRecordSet.EoF
                    sARDocEntry = oRecordSet.Fields.Item("DocEntry").Value

                    If oARInvoice.GetByKey(sARDocEntry) Then
                        oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                        oIncomingPayment.CardCode = oARInvoice.CardCode
                        oIncomingPayment.DocDate = dtDocDate
                        oIncomingPayment.DueDate = dtDocDate
                        oIncomingPayment.TaxDate = dtDocDate
                        oIncomingPayment.UserFields.Fields.Item("U_AB_STNO").Value = oMatrix.Columns.Item("V_7").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_TIME").Value = oMatrix.Columns.Item("V_8").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_SOURCE").Value = oMatrix.Columns.Item("V_9").Cells.Item(iLine).Specific.value
                        oIncomingPayment.UserFields.Fields.Item("U_AB_BRANCHCODE").Value = oMatrix.Columns.Item("V_11").Cells.Item(iLine).Specific.value

                        oIncomingPayment.Invoices.DocEntry = oARInvoice.DocEntry
                        oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                        oIncomingPayment.Invoices.SumApplied = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                        oIncomingPayment.Invoices.Add()
                    End If

                    oRecordSet.MoveNext()
                Loop

                'Bank Transfer
                oIncomingPayment.TransferAccount = oMatrix.Columns.Item("V_5").Cells.Item(iLine).Specific.value
                oIncomingPayment.TransferDate = GetDateTimeValue(oMatrix.Columns.Item("V_3").Cells.Item(iLine).Specific.string)
                oIncomingPayment.TransferSum = oMatrix.Columns.Item("V_2").Cells.Item(iLine).Specific.value
                oIncomingPayment.CashSum = 0

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

                    oMatrix.Columns.Item("V_13").Editable = True
                    oMatrix.Columns.Item("V_13").Cells.Item(iLine).Specific.value = sPayDocEntry
                    objForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oMatrix.Columns.Item("V_13").Editable = False

                    oARInvoice.NumAtCard = oMatrix.Columns.Item("V_4").Cells.Item(iLine).Specific.value
                    oARInvoice.Update()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    bCheck = True
                End If
            Else
                sErrDesc = "Not Matched"
                bCheck = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

            Return bCheck
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_IncomingPayment = RTN_ERROR
        End Try
    End Function
#End Region
    
#Region "Item Event"
    Public Sub UploadStatement_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "RP_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "3" Then
                            showOpenFileDialog(objForm)
                            oEdit = objForm.Items.Item("6").Specific
                            oEdit.Value = sFileName
                            If oEdit.Value = "OpenFileDialog1" Or oEdit.Value = "" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If oEdit.Value = "OpenFileDialog1" Then
                                    oEdit.Value = ""
                                End If
                            End If
                        ElseIf pval.ItemUID = "4" Then
                            If CheckBeforeUpload(objForm, sErrDesc) = False Then
                                p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)

                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        Try
                            Dim oItem, objItem As SAPbouiCOM.Item
                            oItem = objForm.Items.Item("7")
                            objItem = objForm.Items.Item("5")
                            objItem.Top = oItem.Top - 5
                            objItem.Height = oItem.Height + 7
                            objItem.Width = oItem.Width + 5

                            oItem = objForm.Items.Item("6")
                            objItem = objForm.Items.Item("3")
                            objItem.Left = oItem.Left + oItem.Width + 10
                            objItem.Top = oItem.Top
                        Catch ex As Exception

                        End Try

                End Select
            End If
        Catch ex As Exception
            oEdit = objForm.Items.Item("6").Specific
            sFileName = oEdit.Value
            Dim sSQL As String
            sSQL = "DELETE FROM AB_STATEMENTUPLOAD WHERE FileName = '" & sFileName & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error occured while inserting values. Deleteing the inserted values", sFuncName)
            If ExecuteSQLNonQuery(sSQL, sErrDesc) <> RTN_SUCCESS Then
                Call WriteToLogFile("Unable to delete old values in DB/Delete it manually for file " & sFileName, sFuncName)
            End If

            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Menu Event"
    Public Sub UploadStatement_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form
                If pVal.MenuUID = "AE_US" Then
                    LoadFromXML("Upload Statement.srf", p_oSBOApplication)
                    objForm = p_oSBOApplication.Forms.Item("BUPS")
                    objForm.Visible = True
                    InitializeUploadForm(objForm)
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
