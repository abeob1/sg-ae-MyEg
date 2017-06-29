Imports System.Text

Module modProcess

    Private dtBP As DataTable
    Private dtItemCode As DataTable
    Private dtMerchantId As DataTable
    Private dtVatGroup As DataTable
    Private dtValidation As DataTable
    Private sCostCenter5, sCostCenter4, sCostCenter3, sCostCenter2, sCostCenter As String

#Region "Start"
    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty
        Dim sSql As String = String.Empty
        Dim oDataView As DataView
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'sSql = "SELECT ID,Entity,Agency,""Service Type"",Receipt_no,Update_datetime,Tx_amount,eservice_amount,gst_amount,voucher_amount,summons_amount,ppz_amount, " & _
            '       " jpj_amount,comptest_amount,inq_amt,agency_amount,delamount,levifee_amount,deliveryfee,processfee,passfee,visafee,fomafee,insfee,merchant_tx_id,payment_type_id,summons_id, " & _
            '       " summon_type,offence_datetime,offender_name,offender_ic,vehicle_no,law_code2,law_code3,jpj_rev_code,replace_type,""user_id"",id_no,comp_no,account_no,bill_date, " & _
            '       " car_registration_no,prepaid_acct_no,license_class,revenue_code,veh_owner_name,emp_icno,emp_name,passportno,applicantname,sector,print_status,fis_amount, " & _
            '       " photo_amount,ag_code,pay_mode,agency_account_no,zakat_id,req_id,credit_card_no,contact_no,zakat_agency_id,booking_ID,covernote_number,email,ins_company, " & _
            '       " invoiceid,New_Passport_no,""A/P Invoice No"",Section_code,fw_id,trans_id " & _
            '       " FROM public.AB_REVENUEANDCOST WHERE COALESCE(Status,'FAIL') = 'FAIL' OR COALESCE(Status,'') = ''  " & _
            '       " UNION ALL " & _
            '       " SELECT ID,Entity,Agency,""Service Type"",Receipt_no,Update_datetime,Tx_amount,eservice_amount,gst_amount,voucher_amount,summons_amount,ppz_amount, " & _
            '       " jpj_amount,comptest_amount,inq_amt,agency_amount,delamount,levifee_amount,deliveryfee,processfee,passfee,visafee,fomafee,insfee,merchant_tx_id,payment_type_id,summons_id, " & _
            '       " summon_type,offence_datetime,offender_name,offender_ic,vehicle_no,law_code2,law_code3,jpj_rev_code,replace_type,""user_id"",id_no,comp_no,account_no,bill_date, " & _
            '       " car_registration_no,prepaid_acct_no,license_class,revenue_code,veh_owner_name,emp_icno,emp_name,passportno,applicantname,sector,print_status,fis_amount, " & _
            '       " photo_amount,ag_code,pay_mode,agency_account_no,zakat_id,req_id,credit_card_no,contact_no,zakat_agency_id,booking_ID,covernote_number,email,ins_company, " & _
            '       " invoiceid,New_Passport_no,""A/P Invoice No"",Section_code,fw_id,trans_id " & _
            '       " FROM public.AB_REVENUEANDCOST WHERE COALESCE(Status,'FAIL') = 'SUCCESS' AND Agency = 'IMMI' AND COALESCE(""A/P Invoice No2"",'0') = '0' AND print_status = 'SUCCESS' " & _
            '       " ORDER BY ID "
            sSql = "SELECT ID,Entity,Agency,""Service Type"",Receipt_no,Update_datetime,Tx_amount,eservice_amount,gst_amount,voucher_amount,summons_amount,ppz_amount,  " & _
                   " jpj_amount,comptest_amount,inq_amt,agency_amount,delamount,levifee_amount,deliveryfee,processfee,passfee,visafee,fomafee,insfee,merchant_tx_id,payment_type_id, " & _
                   " summons_id,  summon_type,offence_datetime,offender_name,offender_ic,vehicle_no,law_code2,law_code3,jpj_rev_code,replace_type,""user_id"",id_no,comp_no,account_no,bill_date,  " & _
                   " car_registration_no,prepaid_acct_no,license_class,revenue_code,veh_owner_name,emp_icno,emp_name,passportno,applicantname,sector,print_status,fis_amount, " & _
                   " photo_amount,ag_code,pay_mode,agency_account_no,zakat_id,req_id,credit_card_no,contact_no,zakat_agency_id,booking_ID,covernote_number,email,ins_company,  " & _
                   " invoiceid,New_Passport_no,""A/P Invoice No"",Section_code,fw_id,trans_id FROM ( " & _
                   " SELECT ROW_NUMBER() OVER (PARTITION BY Entity) SNO, * FROM (SELECT ID,Entity,Agency,""Service Type"",Receipt_no,Update_datetime,Tx_amount,eservice_amount,gst_amount,voucher_amount,summons_amount,ppz_amount,  " & _
                   " jpj_amount,comptest_amount,inq_amt,agency_amount,delamount,levifee_amount,deliveryfee,processfee,passfee,visafee,fomafee,insfee,merchant_tx_id,payment_type_id, " & _
                   " summons_id,  summon_type,offence_datetime,offender_name,offender_ic,vehicle_no,law_code2,law_code3,jpj_rev_code,replace_type,""user_id"",id_no,comp_no,account_no,bill_date,  " & _
                   " car_registration_no,prepaid_acct_no,license_class,revenue_code,veh_owner_name,emp_icno,emp_name,passportno,applicantname,sector,print_status,fis_amount,  " & _
                   " photo_amount,ag_code,pay_mode,agency_account_no,zakat_id,req_id,credit_card_no,contact_no,zakat_agency_id,booking_ID,covernote_number,email,ins_company,  " & _
                   " invoiceid,New_Passport_no,""A/P Invoice No"",Section_code,fw_id,trans_id  " & _
                   " FROM public.AB_REVENUEANDCOST WHERE COALESCE(Status,'FAIL') = 'FAIL' OR COALESCE(Status,'') = '' " & _
                   " UNION ALL  " & _
                   " SELECT ID,Entity,Agency,""Service Type"",Receipt_no,Update_datetime,Tx_amount,eservice_amount,gst_amount,voucher_amount,summons_amount,ppz_amount,  " & _
                   " jpj_amount,comptest_amount,inq_amt,agency_amount,delamount,levifee_amount,deliveryfee,processfee,passfee,visafee,fomafee,insfee,merchant_tx_id,payment_type_id,summons_id,  " & _
                   " summon_type,offence_datetime,offender_name,offender_ic,vehicle_no,law_code2,law_code3,jpj_rev_code,replace_type,""user_id"",id_no,comp_no,account_no,bill_date,  " & _
                   " car_registration_no,prepaid_acct_no,license_class,revenue_code,veh_owner_name,emp_icno,emp_name,passportno,applicantname,sector,print_status,fis_amount,  " & _
                   " photo_amount,ag_code,pay_mode,agency_account_no,zakat_id,req_id,credit_card_no,contact_no,zakat_agency_id,booking_ID,covernote_number,email,ins_company,  " & _
                   " invoiceid,New_Passport_no,""A/P Invoice No"",Section_code,fw_id,trans_id  " & _
                   " FROM public.AB_REVENUEANDCOST WHERE COALESCE(Status,'FAIL') = 'SUCCESS' AND Agency = 'IMMI' AND print_status = 'SUCCESS' AND COALESCE(""A/P Invoice No2"",'0') = '0'  ) as fin ) AS TAB " & _
                   " ORDER BY ID " 'WHERE TAB.sno <= 5000 

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            oDataView = GetDataView(sSql)

            If Not oDataView Is Nothing Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessDatas()", sFuncName)
                If ProcessDatas(oDataView, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Else
                Console.WriteLine("No Data's found for integration in integration database")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found for integration in Integration database", sFuncName)
                End
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub
#End Region
#Region "Process Datas"
    Public Function ProcessDatas(ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas"
        Dim sEntity As String = String.Empty
        Dim sIntegId As String = String.Empty
        Dim sServiceType As String = String.Empty
        Dim sPrintStatus As String = String.Empty
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "Entity")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "ENTITY") Then
                    oDv.RowFilter = "Entity = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "

                    If oDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CategorizeInvoice()", sFuncName)
                        Dim oDt As DataTable = oDv.ToTable
                        Dim oNewDv As DataView = New DataView(oDt)
                        If CategorizeInvoice(oNewDv, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessDatas = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessDatas = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Process the datas to create invoice"
    Private Function CategorizeInvoice(ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CategorizeInvoice"
        Dim sEntity, sAgency, sPrintStatus, sSQL, sIntegId, sApinvNo As String
        Dim sAGCode As String = String.Empty
        Dim oRecSet As SAPbobsCOM.Recordset = Nothing

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sEntity = oDv(0)(1).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToTargetCompany(p_oCompany, sEntity, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Console.WriteLine("Company Connection Successful")

            If p_oCompany.Connected Then

                oRecSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                sSQL = "SELECT ""Code"",""Name"",""U_ServiceType"" ""ServiceType"",UPPER(""U_REVCOSTCODE"") ""RevCostCode"" FROM " & p_oCompany.CompanyDB & ".""@AE_ITEMCODEMAPPING"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtItemCode = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""CardCode"" FROM " & p_oCompany.CompanyDB & ".""OCRD"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtBP = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT * FROM " & p_oCompany.CompanyDB & ".""@AE_MERCHANT_ID"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtMerchantId = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""ItemCode"",""VatGourpSa"",""VatGroupPu"" FROM " & p_oCompany.CompanyDB & ".""OITM"" WHERE ""frozenFor""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtVatGroup = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                For i As Integer = 0 To oDv.Count - 1
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation()", sFuncName)
                    If Validation(oDv, i, sErrDesc) = RTN_SUCCESS Then

                        sIntegId = oDv(i)(0).ToString.Trim
                        sAgency = oDv(i)(2).ToString.Trim
                        sPrintStatus = oDv(i)(51).ToString.Trim
                        sApinvNo = oDv(i)(68).ToString.Trim
                        sAGCode = oDv(i)(54).ToString.Trim

                        sCostCenter = String.Empty
                        sCostCenter2 = String.Empty
                        sCostCenter3 = String.Empty
                        sCostCenter4 = String.Empty
                        sCostCenter5 = String.Empty

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for ID " & sIntegId, sFuncName)

                        If sPrintStatus = "" Then
                            Console.WriteLine("Agency is " & sAgency)
                        Else
                            Console.WriteLine("Agency is " & sAgency & " and Print Status is " & sPrintStatus)
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Ag Code is " & sAGCode, sFuncName)

                        If sAGCode <> "" Then
                            sSQL = "SELECT ""PrcCode"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE UPPER(""U_AGCODE"") = '" & sAGCode.ToUpper() & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            oRecSet.DoQuery(sSQL)
                            If oRecSet.RecordCount > 0 Then
                                sCostCenter5 = oRecSet.Fields.Item("PrcCode").Value
                            Else
                                sCostCenter5 = ""
                            End If

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cost center for dim5 is " & sCostCenter5, sFuncName)

                            If sCostCenter5 <> "" Then
                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter5 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                oRecSet.DoQuery(sSQL)
                                If oRecSet.RecordCount > 0 Then
                                    sCostCenter4 = oRecSet.Fields.Item("U_DIMENSION_LINK").Value
                                Else
                                    sCostCenter4 = ""
                                End If

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter4 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                oRecSet.DoQuery(sSQL)
                                If oRecSet.RecordCount > 0 Then
                                    sCostCenter3 = oRecSet.Fields.Item("U_DIMENSION_LINK").Value
                                Else
                                    sCostCenter3 = ""
                                End If

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter3 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                oRecSet.DoQuery(sSQL)
                                If oRecSet.RecordCount > 0 Then
                                    sCostCenter2 = oRecSet.Fields.Item("U_DIMENSION_LINK").Value
                                Else
                                    sCostCenter2 = ""
                                End If

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter2.ToUpper() & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                oRecSet.DoQuery(sSQL)
                                If oRecSet.RecordCount > 0 Then
                                    sCostCenter = oRecSet.Fields.Item("U_DIMENSION_LINK").Value
                                Else
                                    sCostCenter = ""
                                End If
                            End If

                            Else
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Ag Code is null/Not getting Cost Center" & sAGCode, sFuncName)
                            End If

                            If sAgency = "IMMI" Then
                                If sPrintStatus = "" And sApinvNo = "" Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARinvoice followed by CreateAPInvoice", sFuncName)

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                                    If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                    If CreateARInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                        If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Continue For
                                    ElseIf CreateAPInvoice_IMMI(oDv, i, sPrintStatus, sErrDesc) = RTN_ERROR Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ", sFuncName)
                                        If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Continue For
                                    Else
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                        If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                ElseIf sPrintStatus.ToUpper() = "SUCCESS" And sApinvNo <> "" Then

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                                    If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                    If CreateAPInvoice_Second(oDv, i, sErrDesc) = RTN_ERROR Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                        If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Continue For
                                    ElseIf CreateCreditNote(oDv, i, sErrDesc) = RTN_ERROR Then
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                        If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Continue For
                                    Else
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                        If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                End If
                            Else
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARinvoice followed by CreateAPInvoice", sFuncName)

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                If CreateARInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                ElseIf CreateAPInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                    If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If

                        Else
                            Continue For
                        End If
                Next
            End If

            Console.WriteLine("Disconnecting Company connection" & sEntity)
            p_oCompany.Disconnect()
            Console.WriteLine("Company disconnected successfully")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CategorizeInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CategorizeInvoice = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Process the datas to create invoice BACKUP"
    Private Function CategorizeInvoice_BACKUP(ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CategorizeInvoice_BACKUP"
        Dim sEntity, sAgency, sPrintStatus, sSQL, sIntegId, sApinvNo As String
        Dim sAGCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sEntity = oDv(0)(1).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToTargetCompany(p_oCompany, sEntity, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            Console.WriteLine("Company Connection Successful")

            '1. Duplicate in Postg sql data
            ' odv.dup
            '2. Duplicate check agaist SAP
            ' sSQL = "SELECT DISTINCT UPPER(""NumAtCard"") AS ""MERCHANTID"", UPPER(""U_AI_InvRefNo"") AS ""RECEIPTNO"",UPPER(""U_TRANS_ID"") AS ""TRANSID"", " & _
            '           " UPPER(""U_FWID"") AS ""FWID"",UPPER(""U_SUMMONSID"") AS ""SUMMONSID"",UPPER(""U_COMPNO"") AS ""COMPOUNDNO"",UPPER(""U_COVERNOTENO"") AS ""COVERNOTENO"" " & _
            ''           " FROM " & p_oCompany.CompanyDB & ".""OINV"" "
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            ' Recordset
            '3. 

            If p_oCompany.Connected Then

                For i As Integer = 0 To oDv.Count - 1

                    sSQL = "SELECT DISTINCT UPPER(""NumAtCard"") AS ""MERCHANTID"", UPPER(""U_AI_InvRefNo"") AS ""RECEIPTNO"",UPPER(""U_TRANS_ID"") AS ""TRANSID"", " & _
                       " UPPER(""U_FWID"") AS ""FWID"",UPPER(""U_SUMMONSID"") AS ""SUMMONSID"",UPPER(""U_COMPNO"") AS ""COMPOUNDNO"",UPPER(""U_COVERNOTENO"") AS ""COVERNOTENO"" " & _
                       " FROM " & p_oCompany.CompanyDB & ".""OINV"" "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                    dtValidation = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Validation()", sFuncName)
                    If Validation(oDv, i, sErrDesc) = RTN_SUCCESS Then

                        sIntegId = oDv(i)(0).ToString.Trim
                        sAgency = oDv(i)(2).ToString.Trim
                        sPrintStatus = oDv(i)(51).ToString.Trim
                        sApinvNo = oDv(i)(68).ToString.Trim
                        sAGCode = oDv(i)(54).ToString.Trim

                        sCostCenter = String.Empty
                        sCostCenter2 = String.Empty
                        sCostCenter3 = String.Empty
                        sCostCenter4 = String.Empty
                        sCostCenter5 = String.Empty

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for ID " & sIntegId, sFuncName)

                        sSQL = "SELECT ""Code"",""Name"",""U_ServiceType"" ""ServiceType"",UPPER(""U_REVCOSTCODE"") ""RevCostCode"" FROM " & p_oCompany.CompanyDB & ".""@AE_ITEMCODEMAPPING"" "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                        dtItemCode = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                        sSQL = "SELECT ""CardCode"" FROM " & p_oCompany.CompanyDB & ".""OCRD"" "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                        dtBP = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                        sSQL = "SELECT * FROM " & p_oCompany.CompanyDB & ".""@AE_MERCHANT_ID"" "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                        dtMerchantId = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                        sSQL = "SELECT ""ItemCode"",""VatGourpSa"",""VatGroupPu"" FROM " & p_oCompany.CompanyDB & ".""OITM"" WHERE ""frozenFor""='N'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                        dtVatGroup = ExecuteQueryReturnDataTable_HANA(sSQL, p_oCompany.CompanyDB)

                        If sPrintStatus = "" Then
                            Console.WriteLine("Agency is " & sAgency)
                        Else
                            Console.WriteLine("Agency is " & sAgency & " and Print Status is " & sPrintStatus)
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Ag Code is " & sAGCode, sFuncName)

                        If sAGCode <> "" Then
                            sSQL = "SELECT ""PrcCode"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE UPPER(""U_AGCODE"") = '" & sAGCode.ToUpper() & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            sCostCenter5 = GetStringValue(sSQL)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cost center for dim5 is " & sCostCenter5, sFuncName)

                            If sCostCenter5 <> "" Then
                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter5 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                sCostCenter4 = GetStringValue(sSQL)

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter4 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                sCostCenter3 = GetStringValue(sSQL)

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter3 & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                sCostCenter2 = GetStringValue(sSQL)

                                sSQL = "SELECT ""U_DIMENSION_LINK"" FROM " & p_oCompany.CompanyDB & ".""OPRC"" WHERE ""PrcCode"" = '" & sCostCenter2.ToUpper() & "' "
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                                sCostCenter = GetStringValue(sSQL)
                            End If

                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Ag Code is null/Not getting Cost Center" & sAGCode, sFuncName)
                        End If

                        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)
                        'If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        'If sAgency = "IMMI" And sPrintStatus.ToUpper = "PENDING" Then
                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice()", sFuncName)
                        '    CreateARInvoice(oDv, i, sErrDesc)
                        'ElseIf sAgency = "IMMI" And sPrintStatus.ToUpper = "PRINTED" Then
                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice()", sFuncName)
                        '    CreateAPInvoice(oDv, i, sErrDesc)
                        'Else
                        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARinvoice followed by CreateAPInvoice", sFuncName)
                        '    If CreateARInvoice(oDv, i, sErrDesc) = RTN_SUCCESS Then
                        '        CreateAPInvoice(oDv, i, sErrDesc)
                        '    End If
                        'End If

                        If sAgency = "IMMI" Then
                            If sPrintStatus = "" And sApinvNo = "" Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARinvoice followed by CreateAPInvoice", sFuncName)

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                If CreateARInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                ElseIf CreateAPInvoice_IMMI(oDv, i, sPrintStatus, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                    If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            ElseIf sPrintStatus.ToUpper() = "SUCCESS" And sApinvNo <> "" Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                If CreateAPInvoice_Second(oDv, i, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                ElseIf CreateCreditNote(oDv, i, sErrDesc) = RTN_ERROR Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Continue For
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                    If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARinvoice followed by CreateAPInvoice", sFuncName)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                            If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            If CreateARInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Continue For
                            ElseIf CreateAPInvoice(oDv, i, sErrDesc) = RTN_ERROR Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
                                If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Continue For
                            Else
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction()", sFuncName)
                                If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If

                        'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
                        'If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    Else
                        Continue For
                    End If
                Next
            End If

            Console.WriteLine("Disconnecting Company connection" & sEntity)
            p_oCompany.Disconnect()
            Console.WriteLine("Company disconnected successfully")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CategorizeInvoice_BACKUP = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction()", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CategorizeInvoice_BACKUP = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Validation"
    Private Function Validation(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "Validation"
        Dim sSQL As String = String.Empty
        Dim sIntegId As String = String.Empty
        Dim sAgency As String = String.Empty
        Dim sServiceType As String = String.Empty
        Dim sMerChantid As String = String.Empty
        Dim sReceiptNo As String = String.Empty
        Dim sTransId As String = String.Empty
        Dim sFwId As String = String.Empty
        Dim sSummonsID As String = String.Empty
        Dim sCompoundNo As String = String.Empty
        Dim sCoverNoteNo As String = String.Empty
        Dim sApinvNo As String = String.Empty

        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            sIntegId = oDv(iLine)(0).ToString.Trim.ToUpper()
            sAgency = oDv(iLine)(2).ToString.Trim.ToUpper()
            sServiceType = oDv(iLine)(3).ToString.Trim.ToUpper()
            sReceiptNo = oDv(iLine)(4).ToString.Trim.ToUpper()
            sMerChantid = oDv(iLine)(24).ToString.Trim.ToUpper()
            sSummonsID = oDv(iLine)(26).ToString.Trim.ToUpper()
            sCompoundNo = oDv(iLine)(38).ToString.Trim.ToUpper()
            sCoverNoteNo = oDv(iLine)(63).ToString.Trim.ToUpper()
            sApinvNo = oDv(iLine)(68).ToString.Trim.ToUpper()
            sFwId = oDv(iLine)(70).ToString.Trim.ToUpper()
            sTransId = oDv(iLine)(71).ToString.Trim.ToUpper()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Service type is " & sServiceType.ToUpper(), sFuncName)

            Select Case sServiceType.ToUpper()
                Case "BOOKING", "CDL", "LDL", "RTX", "STMS", "ETMS"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Receipt No " & sReceiptNo.ToUpper(), sFuncName)

                    sSQL = "SELECT ""U_AI_InvRefNo"" FROM " & p_oCompany.CompanyDB & ".""OINV"" WHERE UPPER(""U_AI_InvRefNo"") = '" & sReceiptNo.ToUpper() & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "RECEIPTNO ::" & sReceiptNo & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

                Case "JPJSUMMONS"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and Transid " & sTransId, sFuncName)

                    sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_TRANS_ID"" AS ""TRANSID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                           " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_TRANS_ID"") = '" & sTransId.ToUpper() & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and TRANSID ::" & sTransId & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

                Case "ZAKAT", "ASSESSMENT", "JIM", "JPN", "ZAKATPPZ", "ZAKATLZS", "ZAKATPKZP", "ZAKATMAINJ", "ZAKATPZNS", "ZAKATMAIP", "ZAKATJZNK"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper(), sFuncName)

                    sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

                Case "MAIDPR"
                    If sAgency.ToUpper() = "IMMI" And sApinvNo = "" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper(), sFuncName)

                        sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sErrDesc = "MERCHANTID ::" & sMerChantid & " already exist in SAP. Function " & sFuncName
                            Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Case "PDRM"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and SummonsId " & sSummonsID.ToUpper(), sFuncName)

                    sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_SUMMONSID"" AS ""SUMMONSID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                           " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_SUMMONSID"") = '" & sSummonsID.ToUpper() & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and SUMMONSID ::" & sSummonsID & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

                Case "COMPOUND"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and Compound no " & sCompoundNo.ToUpper(), sFuncName)

                    sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_COMPNO"" AS ""COMPOUNDNO"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                           " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_COMPNO"") = '" & sCompoundNo.ToUpper() & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and COMPOUNDNO ::" & sCompoundNo & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

                Case "FOREIGN WORKER", "PATI"
                    If sAgency.ToUpper() = "IMMI" Then
                        If sApinvNo = "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and fwid " & sFwId.ToUpper(), sFuncName)

                            sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_FWID"" AS ""FWID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                                   " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_FWID"") = '" & sFwId.ToUpper() & "' "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                            oRecordSet.DoQuery(sSQL)
                            If oRecordSet.RecordCount > 0 Then
                                sErrDesc = "MERCHANTID ::" & sMerChantid & " and FWID ::" & sFwId & " already exist in SAP. Function " & sFuncName
                                Throw New ArgumentException(sErrDesc)
                            End If

                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and fwid " & sFwId.ToUpper(), sFuncName)
                       
                        sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_FWID"" AS ""FWID"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                               " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_FWID"") = '" & sFwId.ToUpper() & "' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        oRecordSet.DoQuery(sSQL)
                        If oRecordSet.RecordCount > 0 Then
                            sErrDesc = "MERCHANTID ::" & sMerChantid & " and FWID ::" & sFwId & " already exist in SAP. Function " & sFuncName
                            Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Case "VEHICLE INSURANCE"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and CoverNote NO " & sCoverNoteNo.ToUpper(), sFuncName)

                    sSQL = "SELECT DISTINCT ""NumAtCard"" AS ""MERCHANTID"",""U_COVERNOTENO"" AS ""COVERNOTENO"" FROM " & p_oCompany.CompanyDB & ".""OINV"" " & _
                           " WHERE UPPER(""NumAtCard"") = '" & sMerChantid.ToUpper() & "' AND UPPER(""U_COVERNOTENO"") = '" & sCoverNoteNo.ToUpper() & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    oRecordSet.DoQuery(sSQL)
                    If oRecordSet.RecordCount > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and COVERNOTENO ::" & sCoverNoteNo & " already exist in SAP. Function " & sFuncName
                        Throw New ArgumentException(sErrDesc)
                    End If

            End Select
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Validation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Validation = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Validation BACKUP"
    Private Function Validation_BACKUP(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "Validation_BACKUP"
        Dim sSQL As String = String.Empty
        Dim sIntegId As String = String.Empty
        Dim sAgency As String = String.Empty
        Dim sServiceType As String = String.Empty
        Dim sMerChantid As String = String.Empty
        Dim sReceiptNo As String = String.Empty
        Dim sTransId As String = String.Empty
        Dim sFwId As String = String.Empty
        Dim sSummonsID As String = String.Empty
        Dim sCompoundNo As String = String.Empty
        Dim sCoverNoteNo As String = String.Empty
        Dim sApinvNo As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            sIntegId = oDv(iLine)(0).ToString.Trim.ToUpper()
            sAgency = oDv(iLine)(2).ToString.Trim.ToUpper()
            sServiceType = oDv(iLine)(3).ToString.Trim.ToUpper()
            sReceiptNo = oDv(iLine)(4).ToString.Trim.ToUpper()
            sMerChantid = oDv(iLine)(24).ToString.Trim.ToUpper()
            sSummonsID = oDv(iLine)(26).ToString.Trim.ToUpper()
            sCompoundNo = oDv(iLine)(38).ToString.Trim.ToUpper()
            sCoverNoteNo = oDv(iLine)(63).ToString.Trim.ToUpper()
            sApinvNo = oDv(iLine)(68).ToString.Trim.ToUpper()
            sFwId = oDv(iLine)(70).ToString.Trim.ToUpper()
            sTransId = oDv(iLine)(71).ToString.Trim.ToUpper()

            dtValidation.DefaultView.RowFilter = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Service type is " & sServiceType.ToUpper(), sFuncName)

            Select Case sServiceType.ToUpper()
                Case "BOOKING", "CDL", "LDL", "RTX", "STMS", "ETMS"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Receipt No " & sReceiptNo.ToUpper(), sFuncName)
                    dtValidation.DefaultView.RowFilter = "RECEIPTNO = '" & sReceiptNo.ToUpper() & "'"
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "RECEIPTNO ::" & sReceiptNo & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                Case "JPJSUMMONS"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and Transid " & sTransId, sFuncName)
                    dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND TRANSID = '" & sTransId.ToUpper() & "'"
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and TRANSID ::" & sTransId & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                Case "ZAKAT", "ASSESSMENT", "JIM", "JPN", "ZAKATPPZ", "ZAKATLZS", "ZAKATPKZP", "ZAKATMAINJ", "ZAKATPZNS", "ZAKATMAIP", "ZAKATJZNK"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper(), sFuncName)
                    dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' "
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                Case "MAIDPR"
                    If sAgency.ToUpper() = "IMMI" And sApinvNo = "" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper(), sFuncName)
                        dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' "
                        If dtValidation.DefaultView.Count > 0 Then
                            sErrDesc = "MERCHANTID ::" & sMerChantid & " already exist in SAP. Function " & sFuncName

                            Dim sQuery As String
                            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                     " WHERE ID = '" & sIntegId & "'"
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                Case "PDRM"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and SummonsId " & sSummonsID.ToUpper(), sFuncName)
                    dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND SUMMONSID = '" & sSummonsID.ToUpper() & "' "
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and SUMMONSID ::" & sSummonsID & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                Case "COMPOUND"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and Compound no " & sCompoundNo.ToUpper(), sFuncName)
                    dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND COMPOUNDNO = '" & sCompoundNo.ToUpper() & "' "
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and COMPOUNDNO ::" & sCompoundNo & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                Case "FOREIGN WORKER", "PATI"
                    If sAgency.ToUpper() = "IMMI" Then
                        If sApinvNo = "" Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and fwid " & sFwId.ToUpper(), sFuncName)
                            dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND FWID = '" & sFwId.ToUpper() & "' "
                            If dtValidation.DefaultView.Count > 0 Then
                                sErrDesc = "MERCHANTID ::" & sMerChantid & " and FWID ::" & sFwId & " already exist in SAP. Function " & sFuncName

                                Dim sQuery As String
                                sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                         " WHERE ID = '" & sIntegId & "'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                                If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and fwid " & sFwId.ToUpper(), sFuncName)
                        dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND FWID = '" & sFwId.ToUpper() & "' "
                        If dtValidation.DefaultView.Count > 0 Then
                            sErrDesc = "MERCHANTID ::" & sMerChantid & " and FWID ::" & sFwId & " already exist in SAP. Function " & sFuncName

                            Dim sQuery As String
                            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                     " WHERE ID = '" & sIntegId & "'"
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                Case "VEHICLE INSURANCE"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking Merchant id " & sMerChantid.ToUpper() & " and CoverNote NO " & sCoverNoteNo.ToUpper(), sFuncName)
                    dtValidation.DefaultView.RowFilter = "MERCHANTID = '" & sMerChantid.ToUpper() & "' AND COVERNOTENO = '" & sCoverNoteNo.ToUpper() & "' "
                    If dtValidation.DefaultView.Count > 0 Then
                        sErrDesc = "MERCHANTID ::" & sMerChantid & " and COVERNOTENO ::" & sCoverNoteNo & " already exist in SAP. Function " & sFuncName

                        Dim sQuery As String
                        sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                                 " WHERE ID = '" & sIntegId & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                        If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Validation_BACKUP = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Validation_BACKUP = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create A/R Invoice"
    Public Function CreateARInvoice(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice"
        Dim oArInovice As SAPbobsCOM.Documents
        Dim sMerChantid As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sIntegId As String = String.Empty
        Dim iCount As Integer
        Dim sItemCode As String = String.Empty
        Dim sAgency As String = String.Empty
        Dim sVatGroup As String = String.Empty
        Dim dEservice As Double = 0.0
        Dim sSql, sEservice_Taxcode As String
        Dim sServiceType As String = String.Empty
        Dim sItemDesc As String = String.Empty
        Dim bLineAdded As Boolean = False
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sIntegId = oDv(iLine)(0).ToString.Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data based on id no " & sIntegId, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating AR Invoice document", sFuncName)
            Console.WriteLine("Creating AR Invoice document")

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSql = "SELECT ""Code"" FROM ""OVTG"" WHERE ""Code"" = '" & p_oCompDef.sEserviceTax & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            oRs.DoQuery(sSql)
            If oRs.RecordCount > 0 Then
                sEservice_Taxcode = oRs.Fields.Item("Code").Value
            Else
                sEservice_Taxcode = ""
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing AR Invoice object", sFuncName)
            oArInovice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            sAgency = oDv(iLine)(2).ToString.Trim
            sServiceType = oDv(iLine)(3).ToString.Trim
            sMerChantid = oDv(iLine)(24).ToString.Trim
            'sAGCode = oDv(iLine)(54).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Service Type is " & sServiceType, sFuncName)

            If sServiceType.ToUpper() = "BOOKING" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Cost center dim5 is " & sCostCenter5, sFuncName)
                If sCostCenter5 = String.Empty Then
                    sErrDesc = "Cost center for dimension 5 is mandatory for booking type"
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            If sMerChantid = "" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Merchant id is null so getting account no as cardcode", sFuncName)
                sCardCode = oDv(iLine)(39).ToString.Trim
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Merchant id is " & sMerChantid, sFuncName)
                For Each c As Char In sMerChantid
                    If Char.IsLetter(c) Then
                        If sCardCode = "" Then
                            sCardCode = c
                        Else
                            sCardCode = sCardCode & c
                        End If
                    Else
                        Exit For
                    End If
                Next
                If sCardCode <> "" Then
                    If sCardCode.Substring(0, 2) = "PP" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Merchant id starts with PP", sFuncName)
                        sCardCode = oDv(iLine)(39).ToString.Trim
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Code in Merchant id is " & sCardCode, sFuncName)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking merchant id in merchant table", sFuncName)
                        dtMerchantId.DefaultView.RowFilter = "Code = '" & sCardCode & "'"
                        If dtMerchantId.DefaultView.Count = 0 Then

                        Else
                            sCardCode = dtMerchantId.DefaultView.Item(0)(1).ToString().Trim()
                        End If
                    End If
                Else
                    sCardCode = oDv(iLine)(39).ToString.Trim
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sCardCode, sFuncName)

            If sCardCode <> "" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking cardcode " & sCardCode & " in BP Master", sFuncName)
                dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                If dtBP.DefaultView.Count = 0 Then
                    sErrDesc = "Cardcode ::" & sCardCode & " provided does not exist in SAP. Function " & sFuncName
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "Cardcode is Mandatory. Function " & sFuncName
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            oArInovice.CardCode = sCardCode
            oArInovice.NumAtCard = sMerChantid
            If Not (oDv(iLine)(5).ToString.Trim = String.Empty) Then
                oArInovice.DocDate = CDate(oDv(iLine)(5).ToString.Trim)
            End If
            oArInovice.BPL_IDAssignedToInvoice = "1"
            oArInovice.UserFields.Fields.Item("U_SERVICETYPE").Value = sServiceType
            oArInovice.UserFields.Fields.Item("U_AI_InvRefNo").Value = oDv(iLine)(4).ToString.Trim
            oArInovice.Comments = "From Integration database. Refer id no " & sIntegId
            oArInovice.JournalMemo = "A/R Invoices - " & sCardCode & " " & sMerChantid

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning UDF fields", sFuncName)

            If Not (oDv(iLine)(25).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_AE_PAYMENTTYPE").Value = oDv(iLine)(25).ToString.Trim
            End If
            If Not (oDv(iLine)(26).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_SUMMONSID").Value = oDv(iLine)(26).ToString.Trim
            End If
            If Not (oDv(iLine)(27).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_SUMMONTYPE").Value = oDv(iLine)(27).ToString.Trim
            End If
            If Not (oDv(iLine)(28).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_OFFENCEDATE").Value = oDv(iLine)(28).ToString.Trim
            End If
            If Not (oDv(iLine)(29).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_OFFENDERNAME").Value = oDv(iLine)(29).ToString.Trim
            End If
            If Not (oDv(iLine)(30).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_OFFENDERIC").Value = oDv(iLine)(30).ToString.Trim
            End If
            If Not (oDv(iLine)(31).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_VEHICLENO").Value = oDv(iLine)(31).ToString.Trim
            End If
            If Not (oDv(iLine)(32).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_LAWCODE2").Value = oDv(iLine)(32).ToString.Trim
            End If
            If Not (oDv(iLine)(33).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_LAWCODE3").Value = oDv(iLine)(33).ToString.Trim
            End If
            If Not (oDv(iLine)(34).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_JPJREVCODE").Value = oDv(iLine)(34).ToString.Trim
            End If
            If Not (oDv(iLine)(35).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_REPLACETYPE").Value = oDv(iLine)(35).ToString.Trim
            End If
            If Not (oDv(iLine)(36).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_USERID").Value = oDv(iLine)(36).ToString.Trim
            End If
            If Not (oDv(iLine)(37).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_IDNO").Value = oDv(iLine)(37).ToString.Trim
            End If
            If Not (oDv(iLine)(38).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_COMPNO").Value = oDv(iLine)(38).ToString.Trim
            End If
            If Not (oDv(iLine)(39).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_ACCOUNTNO").Value = oDv(iLine)(39).ToString.Trim
            End If
            If Not (oDv(iLine)(40).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_BILLDATE").Value = oDv(iLine)(40).ToString.Trim
            End If
            If Not (oDv(iLine)(41).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_CARREGNO").Value = oDv(iLine)(41).ToString.Trim
            End If
            If Not (oDv(iLine)(42).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_PREPAIDACCTNO").Value = oDv(iLine)(42).ToString.Trim
            End If
            If Not (oDv(iLine)(43).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_LICENSECLASS").Value = oDv(iLine)(43).ToString.Trim
            End If
            If Not (oDv(iLine)(44).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_REVENUECODE").Value = oDv(iLine)(44).ToString.Trim
            End If
            If Not (oDv(iLine)(45).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_VEHOWNERNAME").Value = oDv(iLine)(45).ToString.Trim
            End If
            If Not (oDv(iLine)(46).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_EMPICNO").Value = oDv(iLine)(46).ToString.Trim
            End If
            If Not (oDv(iLine)(47).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_EMPNAME").Value = oDv(iLine)(47).ToString.Trim
            End If
            If Not (oDv(iLine)(48).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_PASSPORTNO").Value = oDv(iLine)(48).ToString.Trim
            End If
            If Not (oDv(iLine)(49).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_APPLICANTNAME").Value = oDv(iLine)(49).ToString.Trim
            End If
            If Not (oDv(iLine)(50).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_SECTOR").Value = oDv(iLine)(50).ToString.Trim
            End If
            If Not (oDv(iLine)(51).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_PRINTSTATUS").Value = oDv(iLine)(51).ToString.Trim
            End If
            If Not (oDv(iLine)(54).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_AGCODE").Value = oDv(iLine)(54).ToString.Trim
            End If
            If Not (oDv(iLine)(55).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_PAYMODE").Value = oDv(iLine)(55).ToString.Trim
            End If
            If Not (oDv(iLine)(56).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_ICNO").Value = oDv(iLine)(56).ToString.Trim
            End If
            If Not (oDv(iLine)(57).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_ZAKATID").Value = oDv(iLine)(57).ToString.Trim
            End If
            If Not (oDv(iLine)(58).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_REQID").Value = oDv(iLine)(58).ToString.Trim
            End If
            If Not (oDv(iLine)(59).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_CREDITCARDNO").Value = oDv(iLine)(59).ToString.Trim
            End If
            If Not (oDv(iLine)(60).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_CONTACTNO").Value = oDv(iLine)(60).ToString.Trim
            End If
            If Not (oDv(iLine)(61).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_ZAKATAGENCYID").Value = oDv(iLine)(61).ToString.Trim
            End If
            If Not (oDv(iLine)(62).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_BOOKINGID").Value = oDv(iLine)(62).ToString.Trim
            End If
            If Not (oDv(iLine)(63).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_COVERNOTENO").Value = oDv(iLine)(63).ToString.Trim
            End If
            If Not (oDv(iLine)(64).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_EMAIL").Value = oDv(iLine)(64).ToString.Trim
            End If
            If Not (oDv(iLine)(65).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_VEHINSURANCE").Value = oDv(iLine)(65).ToString.Trim
            End If
            If Not (oDv(iLine)(66).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_FWORKERINSURANCE").Value = oDv(iLine)(66).ToString.Trim
            End If
            If Not (oDv(iLine)(69).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_SECTIONCODE").Value = oDv(iLine)(69).ToString.Trim
            End If
            If Not (oDv(iLine)(70).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_FWID").Value = oDv(iLine)(70).ToString.Trim
            End If
            If Not (oDv(iLine)(71).ToString.Trim = String.Empty) Then
                oArInovice.UserFields.Fields.Item("U_TRANS_ID").Value = oDv(iLine)(71).ToString.Trim
            End If

            iCount = iCount + 1

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line items", sFuncName)

            If Not (oDv(iLine)(7).ToString = String.Empty) Then
                Try
                    dEservice = CDbl(oDv(iLine)(7).ToString.Trim())
                Catch ex As Exception
                    dEservice = 0.0
                End Try
            End If

            '*****ESERVICE AMOUNT COLUMN
            If Not (oDv(iLine)(7).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(7).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    sItemDesc = "eservice_amount" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing Datas for Eservice amount", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for " & sItemDesc, sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'ESERVICE_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''eservice_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(7).ToString.Trim)
                    If sEservice_Taxcode <> "" Then
                        oArInovice.Lines.VatGroup = sEservice_Taxcode
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****GST AMOUNT COLUMN
            If Not (oDv(iLine)(8).ToString = String.Empty) Then
                Dim dGst, dRate As Double
                Try
                    dGst = CDbl(oDv(iLine)(8).ToString.Trim())
                Catch ex As Exception
                    dGst = 0.0
                End Try
                If sAgency.ToUpper = "JPJ" And sServiceType.ToUpper = "BOOKING" Then

                Else        'If sAgency.ToUpper = "JZNK" And sServiceType.ToUpper = "ZAKAT" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Code for GST amount column", sFuncName)

                    If dGst > 0.0 And dEservice = 0.0 Then
                        sSql = "SELECT ""Rate"" FROM ""OVTG"" WHERE ""Code"" = '" & p_oCompDef.sEserviceTax & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                        dRate = GetDoubleValue(sSql)
                        If dRate > 0.0 Then
                            dRate = CDbl(dRate / 100)

                            Dim dEserv_Amount As Double
                            dEserv_Amount = Math.Round(CDbl(dGst / dRate), 2)
                            If dEserv_Amount > 0.0 Then
                                bLineAdded = True
                                If iCount > 1 Then
                                    oArInovice.Lines.Add()
                                End If
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for GST", sFuncName)

                                sItemDesc = "eservice_amount" & "-" & sServiceType
                                dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                                If dtItemCode.DefaultView.Count = 0 Then

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                                    dtItemCode.DefaultView.RowFilter = Nothing
                                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'ESERVICE_AMOUNT'"
                                    If dtItemCode.DefaultView.Count = 0 Then
                                        sErrDesc = "ItemCode ::''eservice_amount'' provided does not exist in SAP(Mapping Table)."
                                        Call WriteToLogFile(sErrDesc, sFuncName)
                                        Throw New ArgumentException(sErrDesc)
                                    Else
                                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                    End If
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If

                                oArInovice.Lines.ItemCode = sItemCode
                                oArInovice.Lines.Quantity = 1
                                oArInovice.Lines.UnitPrice = dEserv_Amount
                                If sEservice_Taxcode <> "" Then
                                    oArInovice.Lines.VatGroup = sEservice_Taxcode
                                End If
                                If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                                    If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                        oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                                    End If
                                Else
                                    If Not (sCostCenter = String.Empty) Then
                                        oArInovice.Lines.CostingCode = sCostCenter
                                    End If
                                    If Not (sCostCenter2 = String.Empty) Then
                                        oArInovice.Lines.CostingCode2 = sCostCenter2
                                    End If
                                    If Not (sCostCenter3 = String.Empty) Then
                                        oArInovice.Lines.CostingCode3 = sCostCenter3
                                    End If
                                    If Not (sCostCenter4 = String.Empty) Then
                                        oArInovice.Lines.CostingCode4 = sCostCenter4
                                    End If
                                    If Not (sCostCenter5 = String.Empty) Then
                                        oArInovice.Lines.CostingCode5 = sCostCenter5
                                    End If
                                End If

                                iCount = iCount + 1
                            End If
                        End If

                    End If
                End If
            End If

            '*****VOUCHER AMOUNT COLUMN
            If Not (oDv(iLine)(9).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(9).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for voucher amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VOUCHER_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''voucher_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for Voucher amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(9).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****SUMMONS AMOUNT COLUMN
            If Not (oDv(iLine)(10).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(10).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for summons amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'SUMMONS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''summons_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for summons amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(10).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****PPZ AMOUNT COLUMN
            If Not (oDv(iLine)(11).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(11).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PPZ amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PPZ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''ppz_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PPZ Amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(11).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****JPJ AMOUNT COLUMN
            If Not (oDv(iLine)(12).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(12).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for JPJ Amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'JPJ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''jpj_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for JPJ Amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(12).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****COMPTEST_AMOUNT COLUMN
            If Not (oDv(iLine)(13).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(13).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    sItemDesc = "comptest_amount" & "-" & sServiceType
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for EHAK amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for comptest_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'COMPTEST_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''comptest_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for EHAK amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(13).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oArInovice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oArInovice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oArInovice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oArInovice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oArInovice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****INQ AMOUNT COLUMN
            If Not (oDv(iLine)(14).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(14).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing Data for INQ Amount", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INQ_AMT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''inq_amt'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for INQ Amount", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(14).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*****AMOUNT COLUMN
            If Not (oDv(iLine)(15).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(15).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    Dim dTxAmount As Double
                    Try
                        dTxAmount = CDbl(oDv(iLine)(6).ToString.Trim())
                    Catch ex As Exception
                        dTxAmount = 0.0
                    End Try
                    Dim dAmount As Double
                    Try
                        dAmount = CDbl(oDv(iLine)(15).ToString.Trim())
                    Catch ex As Exception
                        dAmount = 0.0
                    End Try

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT column", sFuncName)

                    sItemDesc = "agency_amount" & "-" & sServiceType

                    If sAgency.ToUpper = "JPJ" And sServiceType.ToUpper = "JPJSUMMONS" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If
                        If dInvValue = 0.0 Then
                            sErrDesc = "Cannot create invoice.Doctotal is 0"
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If

                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for Amount Column, JPJ and JPJSUMMONS", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                        iCount = iCount + 1

                        If dEservice = 0.0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Eservice Item for Amount Column, JPJ and JPJSUMMONS", sFuncName)

                            If iCount > 1 Then
                                oArInovice.Lines.Add()
                            End If
                            sItemDesc = "eservice_amount" & "-" & sServiceType
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                            If dtItemCode.DefaultView.Count = 0 Then

                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                                dtItemCode.DefaultView.RowFilter = Nothing
                                dtItemCode.DefaultView.RowFilter = "RevCostCode = 'ESERVICE_AMOUNT'"
                                If dtItemCode.DefaultView.Count = 0 Then
                                    sErrDesc = "ItemCode ::''eservice_amount'' provided does not exist in SAP(Mapping Table)."
                                    Call WriteToLogFile(sErrDesc, sFuncName)
                                    Throw New ArgumentException(sErrDesc)
                                Else
                                    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                                End If
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            oArInovice.Lines.ItemCode = sItemCode
                            oArInovice.Lines.Quantity = 1
                            oArInovice.Lines.UnitPrice = CDbl(2)
                            If sEservice_Taxcode <> "" Then
                                oArInovice.Lines.VatGroup = sEservice_Taxcode
                            End If
                            If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                            End If
                            iCount = iCount + 1
                        End If
                    ElseIf sAgency.ToUpper = "JPJ" And sServiceType.ToUpper = "LDL" Then
                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for Amount Column, JPJ and LDL", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(dAmount)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                        iCount = iCount + 1

                        'If dEservice = 0.0 Then
                        '    If iCount > 1 Then
                        '        oArInovice.Lines.Add()
                        '    End If
                        '    sItemDesc = "eservice_amount" & "-" & sServiceType
                        '    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        '    If dtItemCode.DefaultView.Count = 0 Then

                        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                        '        dtItemCode.DefaultView.RowFilter = Nothing
                        '        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'ESERVICE_AMOUNT'"
                        '        If dtItemCode.DefaultView.Count = 0 Then
                        '            sErrDesc = "ItemCode ::''eservice_amount'' provided does not exist in SAP(Mapping Table)."
                        '            Call WriteToLogFile(sErrDesc, sFuncName)
                        '            Throw New ArgumentException(sErrDesc)
                        '        Else
                        '            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        '        End If
                        '    Else
                        '        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        '    End If


                        '    oArInovice.Lines.ItemCode = sItemCode
                        '    oArInovice.Lines.Quantity = 1
                        '    oArInovice.Lines.UnitPrice = CDbl(2)
                        '    If sEservice_Taxcode <> "" Then
                        '        oArInovice.Lines.VatGroup = sEservice_Taxcode
                        '    End If
                        '    If Not (sCostCenter = String.Empty) Then
                        '        oArInovice.Lines.CostingCode = sCostCenter
                        '    End If
                        '    If Not (sCostCenter2 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode2 = sCostCenter2
                        '    End If
                        '    If Not (sCostCenter3 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode3 = sCostCenter3
                        '    End If
                        '    If Not (sCostCenter4 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode4 = sCostCenter4
                        '    End If
                        '    If Not (sCostCenter5 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode5 = sCostCenter5
                        '    End If
                        '    iCount = iCount + 1
                        'End If
                    ElseIf sAgency.ToUpper = "JPJ" And sServiceType.ToUpper = "JPN" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If

                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for Amount Column, JPJ and JPN", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1

                        'If dEservice = 0.0 Then
                        '    If iCount > 1 Then
                        '        oArInovice.Lines.Add()
                        '    End If
                        '    sItemDesc = "eservice_amount" & "-" & sServiceType
                        '    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        '    If dtItemCode.DefaultView.Count = 0 Then

                        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                        '        dtItemCode.DefaultView.RowFilter = Nothing
                        '        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'ESERVICE_AMOUNT'"
                        '        If dtItemCode.DefaultView.Count = 0 Then
                        '            sErrDesc = "ItemCode ::''eservice_amount'' provided does not exist in SAP(Mapping Table)."
                        '            Call WriteToLogFile(sErrDesc, sFuncName)
                        '            Throw New ArgumentException(sErrDesc)
                        '        Else
                        '            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        '        End If
                        '    Else
                        '        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        '    End If


                        '    oArInovice.Lines.ItemCode = sItemCode
                        '    oArInovice.Lines.Quantity = 1
                        '    oArInovice.Lines.UnitPrice = CDbl(2)
                        '    If sEservice_Taxcode <> "" Then
                        '        oArInovice.Lines.VatGroup = sEservice_Taxcode
                        '    End If
                        '    If Not (sCostCenter = String.Empty) Then
                        '        oArInovice.Lines.CostingCode = sCostCenter
                        '    End If
                        '    If Not (sCostCenter2 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode2 = sCostCenter2
                        '    End If
                        '    If Not (sCostCenter3 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode3 = sCostCenter3
                        '    End If
                        '    If Not (sCostCenter4 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode4 = sCostCenter4
                        '    End If
                        '    If Not (sCostCenter5 = String.Empty) Then
                        '        oArInovice.Lines.CostingCode5 = sCostCenter5
                        '    End If
                        '    iCount = iCount + 1
                        'End If
                    Else

                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for eservice_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for Amount Column Not JPJ", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(15).ToString.Trim)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                            If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                            End If
                        Else
                            If Not (sCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = sCostCenter
                            End If
                            If Not (sCostCenter2 = String.Empty) Then
                                oArInovice.Lines.CostingCode2 = sCostCenter2
                            End If
                            If Not (sCostCenter3 = String.Empty) Then
                                oArInovice.Lines.CostingCode3 = sCostCenter3
                            End If
                            If Not (sCostCenter4 = String.Empty) Then
                                oArInovice.Lines.CostingCode4 = sCostCenter4
                            End If
                            If Not (sCostCenter5 = String.Empty) Then
                                oArInovice.Lines.CostingCode5 = sCostCenter5
                            End If
                        End If
                        iCount = iCount + 1
                    End If
                End If
            End If
            '**************DELAMOUNT
            If Not (oDv(iLine)(16).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(16).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'DELAMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''delamount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If

                    sItemDesc = "delamount" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for DELAMOUNT ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for " & sItemDesc, sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for delamount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'DELAMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''delamount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for DELAMOUNT", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(16).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If
            '************LEVIFEE_AMOUNT
            If Not (oDv(iLine)(17).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(17).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    sItemDesc = "levifee_amount" & "-" & sServiceType
                    If sAgency = "IMMI" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PRocessing data for LEVIFEE Amount and IMMI Agency", sFuncName)

                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If
                        'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                        'If dtItemCode.DefaultView.Count = 0 Then
                        '    sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                        '    Call WriteToLogFile(sErrDesc, sFuncName)
                        '    Throw New ArgumentException(sErrDesc)
                        'Else
                        '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        'End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for " & sItemDesc, sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for LEVIFEE Amount and IMMI", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(17).ToString.Trim)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                            If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                            End If
                        Else
                            If Not (sCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = sCostCenter
                            End If
                            If Not (sCostCenter2 = String.Empty) Then
                                oArInovice.Lines.CostingCode2 = sCostCenter2
                            End If
                            If Not (sCostCenter3 = String.Empty) Then
                                oArInovice.Lines.CostingCode3 = sCostCenter3
                            End If
                            If Not (sCostCenter4 = String.Empty) Then
                                oArInovice.Lines.CostingCode4 = sCostCenter4
                            End If
                            If Not (sCostCenter5 = String.Empty) Then
                                oArInovice.Lines.CostingCode5 = sCostCenter5
                            End If
                        End If
                        iCount = iCount + 1

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for LEVIFEE Amount", sFuncName)

                        If iCount > 1 Then
                            oArInovice.Lines.Add()
                        End If
                        'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                        'If dtItemCode.DefaultView.Count = 0 Then
                        '    sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                        '    Call WriteToLogFile(sErrDesc, sFuncName)
                        '    Throw New ArgumentException(sErrDesc)
                        'Else
                        '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        'End If
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for LEVIFEE Amount", sFuncName)

                        oArInovice.Lines.ItemCode = sItemCode
                        oArInovice.Lines.Quantity = 1
                        oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(17).ToString.Trim)
                        If sVatGroup <> "" Then
                            oArInovice.Lines.VatGroup = sVatGroup
                        End If
                        If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                            If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                            End If
                        Else
                            If Not (sCostCenter = String.Empty) Then
                                oArInovice.Lines.CostingCode = sCostCenter
                            End If
                            If Not (sCostCenter2 = String.Empty) Then
                                oArInovice.Lines.CostingCode2 = sCostCenter2
                            End If
                            If Not (sCostCenter3 = String.Empty) Then
                                oArInovice.Lines.CostingCode3 = sCostCenter3
                            End If
                            If Not (sCostCenter4 = String.Empty) Then
                                oArInovice.Lines.CostingCode4 = sCostCenter4
                            End If
                            If Not (sCostCenter5 = String.Empty) Then
                                oArInovice.Lines.CostingCode5 = sCostCenter5
                            End If
                        End If
                        iCount = iCount + 1
                    End If
                End If
            End If
            '*************DELIVERYFEE
            If Not (oDv(iLine)(18).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(18).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing Data for DELIVERYFEE", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'DELIVERYFEE'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''deliveryfee'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for DELIVERYFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(18).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '**********PROCESSFEE
            If Not (oDv(iLine)(19).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(19).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Pocessing data for PROCESSFEE", sFuncName)

                    sItemDesc = "processfee" & "-" & sServiceType

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for PROCESSFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(19).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '**************PASSFEE
            If Not (oDv(iLine)(20).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(20).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    sItemDesc = "passfee" & "-" & sServiceType
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PASSFFEE", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for passfee", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for PASSFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(20).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '**************VISA FEE
            If Not (oDv(iLine)(21).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(21).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    sItemDesc = "visafee" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for VISAFEE", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for passfee", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting data for VISAFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(21).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '************FOMAFEE
            If Not (oDv(iLine)(22).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(22).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for FOMAFEE", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FOMAFEE'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''fomafee'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for FOMAFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(22).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '**********INSFEE
            If Not (oDv(iLine)(23).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(23).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for INSFEE", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''insfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If

                    sItemDesc = "insfee" & "-" & sServiceType

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for insfee", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''insfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for INSFEE", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(23).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '*************FIS AMOUNT
            If Not (oDv(iLine)(52).ToString.Trim = String.Empty) Then
                If (CDbl(oDv(iLine)(52).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for fis_amount", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FIS_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''fis_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If

                    sItemDesc = "fis_amount" & "-" & sServiceType

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for fis_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FIS_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''fis_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for FISAMOUNT ", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(52).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oArInovice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oArInovice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oArInovice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oArInovice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oArInovice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            '****************PHOTO AMOUNT
            If Not (oDv(iLine)(53).ToString.Trim = String.Empty) Then
                If (CDbl(oDv(iLine)(53).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oArInovice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PHOTO_AMOUNT", sFuncName)

                    sItemDesc = "photo_amount" & "-" & sServiceType

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for comptest_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PHOTO_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''photo_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PHOTO_AMOUNT", sFuncName)

                    oArInovice.Lines.ItemCode = sItemCode
                    oArInovice.Lines.Quantity = 1
                    oArInovice.Lines.UnitPrice = CDbl(oDv(iLine)(53).ToString.Trim)
                    If sVatGroup <> "" Then
                        oArInovice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oArInovice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oArInovice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oArInovice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oArInovice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oArInovice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Line Added " & bLineAdded, sFuncName)

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding A/R Invoice Document", sFuncName)
                If oArInovice.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription
                    sErrDesc = sErrDesc.Replace("'", " ")
                    Console.WriteLine("Error while adding A/R invoice document/ " & sErrDesc)
                    sErrDesc = sErrDesc & " in funct. " & sFuncName
                    Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("A/R invoice created successfully", sFuncName)

                    Dim iDocNo, iDocEntry As Integer
                    iDocEntry = p_oCompany.GetNewObjectKey()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oArInovice)

                    Dim sQuery As String

                    Dim oRecordSet As SAPbobsCOM.Recordset
                    sQuery = "SELECT ""DocNum"" FROM ""OINV"" WHERE ""DocEntry"" = '" & iDocEntry & "'"
                    oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("DocNum").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    sQuery = "UPDATE public.AB_REVENUEANDCOST SET ""A/R Invoice No"" = '" & iDocNo & "', " & _
                             " SyncDate = NOW(),Status = 'SUCCESS',""Error Message"" = NULL WHERE ID = '" & sIntegId & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                    If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc.Replace("'", "") & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoice = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create A/p Invoice"
    Public Function CreateAPInvoice(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateAPInvoice"
        Dim oAPInvoice As SAPbobsCOM.Documents
        Dim sCardCode As String
        Dim sNumAtCard As String
        Dim sIntegId As String
        Dim iCount As Integer
        Dim sMerChantid As String = String.Empty
        Dim sItemCode As String = String.Empty
        Dim sVatGroup As String = String.Empty
        Dim sServiceType As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemDesc As String = String.Empty
        Dim bLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Console.WriteLine("Creating AP Invoice")

            oAPInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            sIntegId = oDv(iLine)(0).ToString.Trim
            sCardCode = oDv(iLine)(2).ToString.Trim
            sNumAtCard = oDv(iLine)(4).ToString.Trim
            sServiceType = oDv(iLine)(3).ToString.Trim
            'sAGCode = oDv(iLine)(54).ToString.Trim
            sMerChantid = oDv(iLine)(24).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data based on id no " & sIntegId, sFuncName)

            If sServiceType.ToUpper() = "BOOKING" Then
                If sCostCenter5 = String.Empty Then
                    sErrDesc = "Cost center for dimension 5 is mandatory for booking type"
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            Dim dAmount As Double
            Try
                dAmount = CDbl(oDv(iLine)(15))
            Catch ex As Exception
                dAmount = 0.0
            End Try

            If sCardCode.ToUpper() = "JPJ" And dAmount = 0 Then
                sErrDesc = "agency_amount column Value is 0 for JPJ Service. No Ap Invoice will be Created"

                sSql = "UPDATE public.AB_REVENUEANDCOST SET Status = 'SUCCESS', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                         " WHERE ID = '" & sIntegId & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                If ExecuteNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                CreateAPInvoice = RTN_SUCCESS
                Exit Function
                'ElseIf sCardCode.ToUpper() = "JPJ" And dAmount = 17 Then
                '    sErrDesc = "agency_amount column Value is 17 for JPJ Service. No Ap Invoice will be Created"
                '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                '    CreateAPInvoice = RTN_SUCCESS
                '    Exit Function
            ElseIf sCardCode.ToUpper() = "INSURANCE" Then
                sErrDesc = "Agency is Insurance. No Ap Invoice will be Created"

                sSql = "UPDATE public.AB_REVENUEANDCOST SET Status = 'SUCCESS', ""Error Message"" = '" & sErrDesc & "',SyncDate = NOW() " & _
                         " WHERE ID = '" & sIntegId & "'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                If ExecuteNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                CreateAPInvoice = RTN_SUCCESS
                Exit Function
            End If

            oAPInvoice.CardCode = sCardCode
            oAPInvoice.NumAtCard = sNumAtCard
            If Not (oDv(iLine)(5).ToString.Trim = String.Empty) Then
                oAPInvoice.DocDate = CDate(oDv(iLine)(5).ToString.Trim)
            End If
            oAPInvoice.UserFields.Fields.Item("U_SERVICETYPE").Value = sServiceType
            oAPInvoice.DocDueDate = CDate(oDv(iLine)(5).ToString.Trim)
            oAPInvoice.BPL_IDAssignedToInvoice = "1"
            oAPInvoice.Comments = "From Integration database/Refer id No " & sIntegId

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning values to UDF", sFuncName)

            If Not (oDv(iLine)(24).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_MERCHANT_ID").Value = oDv(iLine)(24).ToString.Trim
            End If
            If Not (oDv(iLine)(25).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AE_PAYMENTTYPE").Value = oDv(iLine)(25).ToString.Trim
            End If
            If Not (oDv(iLine)(26).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONSID").Value = oDv(iLine)(26).ToString.Trim
            End If
            If Not (oDv(iLine)(27).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONTYPE").Value = oDv(iLine)(27).ToString.Trim
            End If
            If Not (oDv(iLine)(28).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENCEDATE").Value = oDv(iLine)(28).ToString.Trim
            End If
            If Not (oDv(iLine)(29).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERNAME").Value = oDv(iLine)(29).ToString.Trim
            End If
            If Not (oDv(iLine)(30).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERIC").Value = oDv(iLine)(30).ToString.Trim
            End If
            If Not (oDv(iLine)(31).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHICLENO").Value = oDv(iLine)(31).ToString.Trim
            End If
            If Not (oDv(iLine)(32).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE2").Value = oDv(iLine)(32).ToString.Trim
            End If
            If Not (oDv(iLine)(33).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE3").Value = oDv(iLine)(33).ToString.Trim
            End If
            If Not (oDv(iLine)(34).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_JPJREVCODE").Value = oDv(iLine)(34).ToString.Trim
            End If
            If Not (oDv(iLine)(35).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REPLACETYPE").Value = oDv(iLine)(35).ToString.Trim
            End If
            If Not (oDv(iLine)(36).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_USERID").Value = oDv(iLine)(36).ToString.Trim
            End If
            If Not (oDv(iLine)(37).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_IDNO").Value = oDv(iLine)(37).ToString.Trim
            End If
            If Not (oDv(iLine)(38).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COMPNO").Value = oDv(iLine)(38).ToString.Trim
            End If
            If Not (oDv(iLine)(39).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ACCOUNTNO").Value = oDv(iLine)(39).ToString.Trim
            End If
            If Not (oDv(iLine)(40).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BILLDATE").Value = oDv(iLine)(40).ToString.Trim
            End If
            If Not (oDv(iLine)(41).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CARREGNO").Value = oDv(iLine)(41).ToString.Trim
            End If
            If Not (oDv(iLine)(42).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PREPAIDACCTNO").Value = oDv(iLine)(42).ToString.Trim
            End If
            If Not (oDv(iLine)(43).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LICENSECLASS").Value = oDv(iLine)(43).ToString.Trim
            End If
            If Not (oDv(iLine)(44).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REVENUECODE").Value = oDv(iLine)(44).ToString.Trim
            End If
            If Not (oDv(iLine)(45).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHOWNERNAME").Value = oDv(iLine)(45).ToString.Trim
            End If
            If Not (oDv(iLine)(46).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPICNO").Value = oDv(iLine)(46).ToString.Trim
            End If
            If Not (oDv(iLine)(47).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPNAME").Value = oDv(iLine)(47).ToString.Trim
            End If
            If Not (oDv(iLine)(48).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PASSPORTNO").Value = oDv(iLine)(48).ToString.Trim
            End If
            If Not (oDv(iLine)(49).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_APPLICANTNAME").Value = oDv(iLine)(49).ToString.Trim
            End If
            If Not (oDv(iLine)(50).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTOR").Value = oDv(iLine)(50).ToString.Trim
            End If
            If Not (oDv(iLine)(51).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PRINTSTATUS").Value = oDv(iLine)(51).ToString.Trim
            End If
            If Not (oDv(iLine)(54).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AGCODE").Value = oDv(iLine)(54).ToString.Trim
            End If
            If Not (oDv(iLine)(55).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PAYMODE").Value = oDv(iLine)(55).ToString.Trim
            End If
            If Not (oDv(iLine)(56).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ICNO").Value = oDv(iLine)(56).ToString.Trim
            End If
            If Not (oDv(iLine)(57).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATID").Value = oDv(iLine)(57).ToString.Trim
            End If
            If Not (oDv(iLine)(58).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REQID").Value = oDv(iLine)(58).ToString.Trim
            End If
            If Not (oDv(iLine)(59).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CREDITCARDNO").Value = oDv(iLine)(59).ToString.Trim
            End If
            If Not (oDv(iLine)(60).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CONTACTNO").Value = oDv(iLine)(60).ToString.Trim
            End If
            If Not (oDv(iLine)(61).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATAGENCYID").Value = oDv(iLine)(61).ToString.Trim
            End If
            If Not (oDv(iLine)(62).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BOOKINGID").Value = oDv(iLine)(62).ToString.Trim
            End If
            If Not (oDv(iLine)(63).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COVERNOTENO").Value = oDv(iLine)(63).ToString.Trim
            End If
            If Not (oDv(iLine)(64).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMAIL").Value = oDv(iLine)(64).ToString.Trim
            End If
            If Not (oDv(iLine)(65).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHINSURANCE").Value = oDv(iLine)(65).ToString.Trim
            End If
            If Not (oDv(iLine)(66).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWORKERINSURANCE").Value = oDv(iLine)(66).ToString.Trim
            End If
            If Not (oDv(iLine)(69).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTIONCODE").Value = oDv(iLine)(69).ToString.Trim
            End If
            If Not (oDv(iLine)(70).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWID").Value = oDv(iLine)(70).ToString.Trim
            End If
            If Not (oDv(iLine)(71).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_TRANS_ID").Value = oDv(iLine)(71).ToString.Trim
            End If

            iCount = iCount + 1

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line Items", sFuncName)

            If Not (oDv(iLine)(17).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(17).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    sItemDesc = "levifee_amount" & "-" & sServiceType
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Process datas for LEVIFEE_AMOUNT", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for LEVIFEE_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(17).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            If Not (oDv(iLine)(10).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(10).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for SUMMONS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'SUMMONS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''summons_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for SUMMONS_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(10).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(11).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(11).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for PPZ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PPZ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''ppz_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PPZ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(11).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(12).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(12).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("processing data for JPJ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'JPJ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''jpj_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for JPJ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(12).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(14).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(14).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for INQ_AMT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INQ_AMT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''inq_amt'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for INQ_AMT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(14).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(15).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(15).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    Dim dTxAmount As Double
                    Try
                        dTxAmount = CDbl(oDv(iLine)(6).ToString.Trim())
                    Catch ex As Exception
                        dTxAmount = 0.0
                    End Try
                    Try
                        dAmount = CDbl(oDv(iLine)(15).ToString.Trim())
                    Catch ex As Exception
                        dAmount = 0.0
                    End Try
                    sItemDesc = "agency_amount" & "-" & sServiceType
                    If sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "CDL" Then
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT,JPJ and CDL", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and CDL", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dAmount)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                        iCount = iCount + 1
                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "JPJSUMMONS" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If
                        If dInvValue = 0.0 Then
                            sErrDesc = "Cannot create invoice.Doctotal is 0"
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT,JPJ and JPJSUMMONS", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and JPJSUMMONS", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                        iCount = iCount + 1
                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "LDL" Then
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT,JPJ and LDL", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for AMOUNT,JPJ and LDL", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dAmount)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                        iCount = iCount + 1

                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "JPN" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If
                        If dInvValue = 0.0 Then
                            sErrDesc = "Cannot create invoice.Doctotal is 0"
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT,JPJ and JPN", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and JPN", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AGENCY_AMOUNT", sFuncName)

                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for AGENCY_AMOUNT", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(15).ToString.Trim)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Then
                            If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                                oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                            End If
                        Else
                            If Not (sCostCenter = String.Empty) Then
                                oAPInvoice.Lines.CostingCode = sCostCenter
                            End If
                            If Not (sCostCenter2 = String.Empty) Then
                                oAPInvoice.Lines.CostingCode2 = sCostCenter2
                            End If
                            If Not (sCostCenter3 = String.Empty) Then
                                oAPInvoice.Lines.CostingCode3 = sCostCenter3
                            End If
                            If Not (sCostCenter4 = String.Empty) Then
                                oAPInvoice.Lines.CostingCode4 = sCostCenter4
                            End If
                            If Not (sCostCenter5 = String.Empty) Then
                                oAPInvoice.Lines.CostingCode5 = sCostCenter5
                            End If
                        End If
                        iCount = iCount + 1
                    End If

                End If
            End If
            If Not (oDv(iLine)(52).ToString.Trim = String.Empty) Then
                If (CDbl(oDv(iLine)(52).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for FIS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FIS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''fis_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for FIS_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(52).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If sServiceType.ToUpper() = "CDL" Or sServiceType.ToUpper() = "LDL" Or sServiceType.ToUpper() = "RTX" Or sServiceType.ToUpper() = "ETMS" Or sServiceType.ToUpper() = "STMS" Or sServiceType.ToUpper() = "JPJSUMMONS" Then
                        If Not (p_oCompDef.sBookingCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = p_oCompDef.sBookingCostCenter
                        End If
                    Else
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                    End If
                    iCount = iCount + 1
                End If
            End If

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding A/P Invoice Document", sFuncName)

                If oAPInvoice.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription
                    sErrDesc = sErrDesc.Replace("'", " ")
                    Console.WriteLine("Error while adding A/p invoice document/ " & sErrDesc)
                    sErrDesc = sErrDesc & " in funct. " & sFuncName

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding A/p invoice document", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("1 A/P invoice created successfully", sFuncName)

                    Dim iDocNo, iDocEntry As Integer
                    iDocEntry = p_oCompany.GetNewObjectKey()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAPInvoice)

                    Dim sQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset

                    sQuery = "SELECT ""DocNum"" FROM ""OPCH"" WHERE ""DocEntry"" = '" & iDocEntry & "'"
                    oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("DocNum").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    sQuery = "UPDATE public.AB_REVENUEANDCOST SET ""A/P Invoice No"" = '" & iDocNo & "', " & _
                             " SyncDate = NOW(),Status = 'SUCCESS',""Error Message"" = NULL WHERE ID = '" & sIntegId & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                    If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAPInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc.Replace("'", "") & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAPInvoice = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create A/p Invoice for IMMI"
    Public Function CreateAPInvoice_IMMI(ByVal oDv As DataView, ByVal iLine As Integer, ByVal sPrintStatus As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateAPInvoice_IMMI"
        Dim oAPInvoice As SAPbobsCOM.Documents
        Dim sCardCode As String
        Dim sNumAtCard As String
        Dim sIntegId As String
        Dim iCount As Integer
        Dim sMerChantid As String = String.Empty
        Dim sItemCode As String = String.Empty
        Dim sVatGroup As String = String.Empty
        Dim sServiceType As String = String.Empty
        Dim sSql As String = String.Empty
        Dim sItemDesc As String = String.Empty
        Dim sCardName As String = String.Empty
        Dim bLineAdded As Boolean = False
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Console.WriteLine("Creating AP Invoice")

            oAPInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            sIntegId = oDv(iLine)(0).ToString.Trim
            sCardCode = oDv(iLine)(46).ToString.Trim
            sCardName = oDv(iLine)(47).ToString.Trim
            sMerChantid = oDv(iLine)(24).ToString.Trim

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data based on id no " & sIntegId, sFuncName)

            If sCardCode <> "" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Checking cardcode " & sCardCode & " in BP Master", sFuncName)
                dtBP.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
                If dtBP.DefaultView.Count = 0 Then
                    sErrDesc = "Cardcode " & sCardCode & " not exists in SAP."
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                    Console.WriteLine("Cardcode " & sCardCode & " not exists in SAP.")
                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateBP()", sFuncName)
                    'If CreateBP(sCardCode, sCardName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    'Console.WriteLine("BP Master " & sCardCode & " created successfully")
                    Throw New ArgumentException(sErrDesc)
                End If
            Else
                sErrDesc = "CardCode cannot be null. Function " & sFuncName
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sNumAtCard = oDv(iLine)(4).ToString.Trim
            sServiceType = oDv(iLine)(3).ToString.Trim

            If sServiceType.ToUpper() = "BOOKING" Then
                If sCostCenter5 = String.Empty Then
                    sErrDesc = "Cost center for dimension 5 is mandatory for booking type"
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            Dim dAmount As Double
            Try
                dAmount = CDbl(oDv(iLine)(15))
            Catch ex As Exception
                dAmount = 0.0
            End Try

            oAPInvoice.CardCode = sCardCode
            oAPInvoice.NumAtCard = sNumAtCard
            If Not (oDv(iLine)(5).ToString.Trim = String.Empty) Then
                oAPInvoice.DocDate = CDate(oDv(iLine)(5).ToString.Trim)
            End If
            oAPInvoice.UserFields.Fields.Item("U_SERVICETYPE").Value = sServiceType
            oAPInvoice.DocDueDate = CDate(oDv(iLine)(5).ToString.Trim)
            oAPInvoice.BPL_IDAssignedToInvoice = "1"
            oAPInvoice.Comments = "From Integration database/Refer id No " & sIntegId

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning values to UDF", sFuncName)
            'sMerChantid = oDv(iLine)(24).ToString.Trim
            If Not (oDv(iLine)(24).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_MERCHANT_ID").Value = oDv(iLine)(24).ToString.Trim
            End If
            If Not (oDv(iLine)(25).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AE_PAYMENTTYPE").Value = oDv(iLine)(25).ToString.Trim
            End If
            If Not (oDv(iLine)(26).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONSID").Value = oDv(iLine)(26).ToString.Trim
            End If
            If Not (oDv(iLine)(27).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONTYPE").Value = oDv(iLine)(27).ToString.Trim
            End If
            If Not (oDv(iLine)(28).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENCEDATE").Value = oDv(iLine)(28).ToString.Trim
            End If
            If Not (oDv(iLine)(29).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERNAME").Value = oDv(iLine)(29).ToString.Trim
            End If
            If Not (oDv(iLine)(30).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERIC").Value = oDv(iLine)(30).ToString.Trim
            End If
            If Not (oDv(iLine)(31).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHICLENO").Value = oDv(iLine)(31).ToString.Trim
            End If
            If Not (oDv(iLine)(32).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE2").Value = oDv(iLine)(32).ToString.Trim
            End If
            If Not (oDv(iLine)(33).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE3").Value = oDv(iLine)(33).ToString.Trim
            End If
            If Not (oDv(iLine)(34).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_JPJREVCODE").Value = oDv(iLine)(34).ToString.Trim
            End If
            If Not (oDv(iLine)(35).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REPLACETYPE").Value = oDv(iLine)(35).ToString.Trim
            End If
            If Not (oDv(iLine)(36).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_USERID").Value = oDv(iLine)(36).ToString.Trim
            End If
            If Not (oDv(iLine)(37).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_IDNO").Value = oDv(iLine)(37).ToString.Trim
            End If
            If Not (oDv(iLine)(38).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COMPNO").Value = oDv(iLine)(38).ToString.Trim
            End If
            If Not (oDv(iLine)(39).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ACCOUNTNO").Value = oDv(iLine)(39).ToString.Trim
            End If
            If Not (oDv(iLine)(40).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BILLDATE").Value = oDv(iLine)(40).ToString.Trim
            End If
            If Not (oDv(iLine)(41).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CARREGNO").Value = oDv(iLine)(41).ToString.Trim
            End If
            If Not (oDv(iLine)(42).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PREPAIDACCTNO").Value = oDv(iLine)(42).ToString.Trim
            End If
            If Not (oDv(iLine)(43).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LICENSECLASS").Value = oDv(iLine)(43).ToString.Trim
            End If
            If Not (oDv(iLine)(44).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REVENUECODE").Value = oDv(iLine)(44).ToString.Trim
            End If
            If Not (oDv(iLine)(45).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHOWNERNAME").Value = oDv(iLine)(45).ToString.Trim
            End If
            If Not (oDv(iLine)(46).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPICNO").Value = oDv(iLine)(46).ToString.Trim
            End If
            If Not (oDv(iLine)(47).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPNAME").Value = oDv(iLine)(47).ToString.Trim
            End If
            If Not (oDv(iLine)(48).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PASSPORTNO").Value = oDv(iLine)(48).ToString.Trim
            End If
            If Not (oDv(iLine)(49).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_APPLICANTNAME").Value = oDv(iLine)(49).ToString.Trim
            End If
            If Not (oDv(iLine)(50).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTOR").Value = oDv(iLine)(50).ToString.Trim
            End If
            If Not (oDv(iLine)(51).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PRINTSTATUS").Value = oDv(iLine)(51).ToString.Trim
            End If
            If Not (oDv(iLine)(54).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AGCODE").Value = oDv(iLine)(54).ToString.Trim
            End If
            If Not (oDv(iLine)(55).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PAYMODE").Value = oDv(iLine)(55).ToString.Trim
            End If
            If Not (oDv(iLine)(56).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ICNO").Value = oDv(iLine)(56).ToString.Trim
            End If
            If Not (oDv(iLine)(57).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATID").Value = oDv(iLine)(57).ToString.Trim
            End If
            If Not (oDv(iLine)(58).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REQID").Value = oDv(iLine)(58).ToString.Trim
            End If
            If Not (oDv(iLine)(59).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CREDITCARDNO").Value = oDv(iLine)(59).ToString.Trim
            End If
            If Not (oDv(iLine)(60).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CONTACTNO").Value = oDv(iLine)(60).ToString.Trim
            End If
            If Not (oDv(iLine)(61).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATAGENCYID").Value = oDv(iLine)(61).ToString.Trim
            End If
            If Not (oDv(iLine)(62).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BOOKINGID").Value = oDv(iLine)(62).ToString.Trim
            End If
            If Not (oDv(iLine)(63).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COVERNOTENO").Value = oDv(iLine)(63).ToString.Trim
            End If
            If Not (oDv(iLine)(64).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMAIL").Value = oDv(iLine)(64).ToString.Trim
            End If
            If Not (oDv(iLine)(65).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHINSURANCE").Value = oDv(iLine)(65).ToString.Trim
            End If
            If Not (oDv(iLine)(66).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWORKERINSURANCE").Value = oDv(iLine)(66).ToString.Trim
            End If
            If Not (oDv(iLine)(69).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTIONCODE").Value = oDv(iLine)(69).ToString.Trim
            End If
            If Not (oDv(iLine)(70).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWID").Value = oDv(iLine)(70).ToString.Trim
            End If
            If Not (oDv(iLine)(71).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_TRANS_ID").Value = oDv(iLine)(71).ToString.Trim
            End If

            iCount = iCount + 1

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line Items", sFuncName)

            If Not (oDv(iLine)(17).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(17).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    sItemDesc = "levifee_amount" & "-" & sServiceType
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for LEVIFEE_AMOUNT", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for LEVIFEE_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(17).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(19).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(19).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PROCESSFEE", sFuncName)

                    sItemDesc = "processfee" & "-" & sServiceType

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PROCESSFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(19).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(20).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(20).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    sItemDesc = "passfee" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PASSFEE", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for passfee", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for PASSFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(20).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(21).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(21).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for VISAFEE", sFuncName)

                    sItemDesc = "visafee" & "-" & sServiceType

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for VISAFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(21).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(10).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(10).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for SUMMONS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'SUMMONS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''summons_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for SUMMONS_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(10).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(11).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(11).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PPZ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PPZ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''ppz_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PPZ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(11).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(12).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(12).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for JPJ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'JPJ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''jpj_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for JPJ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(12).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(13).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(13).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for EHAK_AMOUNT", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'EHAK_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''ehak_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    '    End

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'COMPTEST_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''comptest_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for EHAK_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(13).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(14).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(14).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for INQ_AMT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INQ_AMT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''inq_amt'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for INQ_AMT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(14).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(15).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(15).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    Dim dTxAmount As Double
                    Try
                        dTxAmount = CDbl(oDv(iLine)(6).ToString.Trim())
                    Catch ex As Exception
                        dTxAmount = 0.0
                    End Try
                    Try
                        dAmount = CDbl(oDv(iLine)(15).ToString.Trim())
                    Catch ex As Exception
                        dAmount = 0.0
                    End Try
                    sItemDesc = "agency_amount" & "-" & sServiceType
                    If sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "BOOKING" Then 'And dAmount = 27
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for AMOUNT,JPJ and BOOKING", sFuncName)

                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and BOOKING", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(10)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "CDL" Then
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data AMOUNT,JPJ and CDL", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and CDL", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dAmount)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "JPJSUMMONS" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If
                        If dInvValue = 0.0 Then
                            sErrDesc = "Cannot create invoice.Doctotal is 0"
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data AMOUNT,JPJ and JPJSUMMONS", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for AMOUNT,JPJ and JPJSUMMONS", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "LDL" Then
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data AMOUNT,JPJ and LDL", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item AMOUNT,JPJ and LDL", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dAmount)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1

                    ElseIf sCardCode.ToUpper = "JPJ" And sServiceType.ToUpper = "JPN" Then
                        Dim dInvValue As Double
                        If dAmount > 0 Then
                            dInvValue = dAmount
                        ElseIf dTxAmount > 0 Then
                            dInvValue = dTxAmount
                        End If
                        If dInvValue = 0.0 Then
                            sErrDesc = "Cannot create invoice.Doctotal is 0"
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data AMOUNT,JPJ and JPN", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT,JPJ and JPN", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(dInvValue)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    Else
                        If iCount > 1 Then
                            oAPInvoice.Lines.Add()
                        End If
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data AMOUNT", sFuncName)

                        dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                            dtItemCode.DefaultView.RowFilter = Nothing
                            dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                            If dtItemCode.DefaultView.Count = 0 Then
                                sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                            End If
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If

                        dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                        If dtVatGroup.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                        End If

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for AMOUNT", sFuncName)

                        oAPInvoice.Lines.ItemCode = sItemCode
                        oAPInvoice.Lines.Quantity = 1
                        oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(15).ToString.Trim)
                        If sVatGroup <> "" Then
                            oAPInvoice.Lines.VatGroup = sVatGroup
                        End If
                        If Not (sCostCenter = String.Empty) Then
                            oAPInvoice.Lines.CostingCode = sCostCenter
                        End If
                        If Not (sCostCenter2 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode2 = sCostCenter2
                        End If
                        If Not (sCostCenter3 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode3 = sCostCenter3
                        End If
                        If Not (sCostCenter4 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode4 = sCostCenter4
                        End If
                        If Not (sCostCenter5 = String.Empty) Then
                            oAPInvoice.Lines.CostingCode5 = sCostCenter5
                        End If
                        iCount = iCount + 1
                    End If

                End If
            End If

            If Not (oDv(iLine)(52).ToString.Trim = String.Empty) Then
                If (CDbl(oDv(iLine)(52).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for FIS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FIS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''fis_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If


                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for FIS_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(23).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding A/P Invoice Document", sFuncName)

                If oAPInvoice.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription
                    sErrDesc = sErrDesc.Replace("'", " ")
                    Console.WriteLine("Error while adding A/p invoice document/ " & sErrDesc)
                    sErrDesc = sErrDesc & " in funct. " & sFuncName

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding A/p invoice document", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("A/P invoice -IMMI created successfully", sFuncName)

                    Dim iDocNo, iDocEntry As Integer
                    iDocEntry = p_oCompany.GetNewObjectKey()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAPInvoice)

                    Dim sQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset

                    sQuery = "SELECT ""DocNum"" FROM ""OPCH"" WHERE ""DocEntry"" = '" & iDocEntry & "'"
                    oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("DocNum").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    sQuery = "UPDATE public.AB_REVENUEANDCOST SET ""A/P Invoice No"" = '" & iDocNo & "', " & _
                             " SyncDate = NOW(),Status = 'SUCCESS',""Error Message"" = NULL WHERE ID = '" & sIntegId & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                    If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                    
                End If
            End If
            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAPInvoice_IMMI = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc.Replace("'", "") & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAPInvoice_IMMI = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create Second A/p Invoice"
    Public Function CreateAPInvoice_Second(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateAPInvoice_Second"
        Dim oAPInvoice As SAPbobsCOM.Documents
        Dim sCardCode As String
        Dim sNumAtCard As String
        Dim sIntegId As String
        Dim iCount As Integer
        Dim sItemCode As String = String.Empty
        Dim sVatGroup As String = String.Empty
        Dim sServiceType As String = String.Empty
        'Dim sAGCode As String = String.Empty
        'Dim sCostCenter5, sCostCenter4, sCostCenter3, sCostCenter2, sCostCenter As String
        Dim sSql As String = String.Empty
        Dim sItemDesc As String = String.Empty
        Dim sPassPortNo As String = String.Empty
        Dim bLineAdded As Boolean = False
        Dim sMerChantid As String = String.Empty
        
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            Console.WriteLine("Creating AP Invoice")

            oAPInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            sIntegId = oDv(iLine)(0).ToString.Trim
            sCardCode = oDv(iLine)(2).ToString.Trim
            sNumAtCard = oDv(iLine)(4).ToString.Trim
            sServiceType = oDv(iLine)(3).ToString.Trim
            'sAGCode = oDv(iLine)(54).ToString.Trim
            sMerChantid = oDv(iLine)(24).ToString.Trim
            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data based on id no " & sIntegId, sFuncName)

            sPassPortNo = oDv(iLine)(48).ToString.Trim
            If sPassPortNo = "" Then
                sErrDesc = "Passport No. column value should not be null"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            If sServiceType.ToUpper() = "BOOKING" Then
                If sCostCenter5 = String.Empty Then
                    sErrDesc = "Cost center for dimension 5 is mandatory for booking type"
                    Throw New ArgumentException(sErrDesc)
                End If
            End If

            Dim dAmount As Double
            Try
                dAmount = CDbl(oDv(iLine)(15))
            Catch ex As Exception
                dAmount = 0.0
            End Try

            oAPInvoice.CardCode = sCardCode
            oAPInvoice.NumAtCard = sNumAtCard
            If Not (oDv(iLine)(5).ToString.Trim = String.Empty) Then
                oAPInvoice.DocDate = CDate(oDv(iLine)(5).ToString.Trim)
            End If
            oAPInvoice.UserFields.Fields.Item("U_SERVICETYPE").Value = sServiceType
            oAPInvoice.DocDueDate = CDate(oDv(iLine)(5).ToString.Trim)
            oAPInvoice.BPL_IDAssignedToInvoice = "1"
            oAPInvoice.Comments = "From Integration database/Refer id No " & sIntegId

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assinging values to UDF", sFuncName)

            If Not (oDv(iLine)(24).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_MERCHANT_ID").Value = oDv(iLine)(24).ToString.Trim
            End If
            If Not (oDv(iLine)(25).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AE_PAYMENTTYPE").Value = oDv(iLine)(25).ToString.Trim
            End If
            If Not (oDv(iLine)(26).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONSID").Value = oDv(iLine)(26).ToString.Trim
            End If
            If Not (oDv(iLine)(27).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SUMMONTYPE").Value = oDv(iLine)(27).ToString.Trim
            End If
            If Not (oDv(iLine)(28).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENCEDATE").Value = oDv(iLine)(28).ToString.Trim
            End If
            If Not (oDv(iLine)(29).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERNAME").Value = oDv(iLine)(29).ToString.Trim
            End If
            If Not (oDv(iLine)(30).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_OFFENDERIC").Value = oDv(iLine)(30).ToString.Trim
            End If
            If Not (oDv(iLine)(31).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHICLENO").Value = oDv(iLine)(31).ToString.Trim
            End If
            If Not (oDv(iLine)(32).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE2").Value = oDv(iLine)(32).ToString.Trim
            End If
            If Not (oDv(iLine)(33).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LAWCODE3").Value = oDv(iLine)(33).ToString.Trim
            End If
            If Not (oDv(iLine)(34).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_JPJREVCODE").Value = oDv(iLine)(34).ToString.Trim
            End If
            If Not (oDv(iLine)(35).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REPLACETYPE").Value = oDv(iLine)(35).ToString.Trim
            End If
            If Not (oDv(iLine)(36).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_USERID").Value = oDv(iLine)(36).ToString.Trim
            End If
            If Not (oDv(iLine)(37).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_IDNO").Value = oDv(iLine)(37).ToString.Trim
            End If
            If Not (oDv(iLine)(38).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COMPNO").Value = oDv(iLine)(38).ToString.Trim
            End If
            If Not (oDv(iLine)(39).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ACCOUNTNO").Value = oDv(iLine)(39).ToString.Trim
            End If
            If Not (oDv(iLine)(40).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BILLDATE").Value = oDv(iLine)(40).ToString.Trim
            End If
            If Not (oDv(iLine)(41).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CARREGNO").Value = oDv(iLine)(41).ToString.Trim
            End If
            If Not (oDv(iLine)(42).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PREPAIDACCTNO").Value = oDv(iLine)(42).ToString.Trim
            End If
            If Not (oDv(iLine)(43).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_LICENSECLASS").Value = oDv(iLine)(43).ToString.Trim
            End If
            If Not (oDv(iLine)(44).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REVENUECODE").Value = oDv(iLine)(44).ToString.Trim
            End If
            If Not (oDv(iLine)(45).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHOWNERNAME").Value = oDv(iLine)(45).ToString.Trim
            End If
            If Not (oDv(iLine)(46).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPICNO").Value = oDv(iLine)(46).ToString.Trim
            End If
            If Not (oDv(iLine)(47).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMPNAME").Value = oDv(iLine)(47).ToString.Trim
            End If
            If Not (oDv(iLine)(48).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PASSPORTNO").Value = oDv(iLine)(48).ToString.Trim
            End If
            If Not (oDv(iLine)(49).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_APPLICANTNAME").Value = oDv(iLine)(49).ToString.Trim
            End If
            If Not (oDv(iLine)(50).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTOR").Value = oDv(iLine)(50).ToString.Trim
            End If
            If Not (oDv(iLine)(51).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PRINTSTATUS").Value = oDv(iLine)(51).ToString.Trim
            End If
            If Not (oDv(iLine)(54).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_AGCODE").Value = oDv(iLine)(54).ToString.Trim
            End If
            If Not (oDv(iLine)(55).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_PAYMODE").Value = oDv(iLine)(55).ToString.Trim
            End If
            If Not (oDv(iLine)(56).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ICNO").Value = oDv(iLine)(56).ToString.Trim
            End If
            If Not (oDv(iLine)(57).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATID").Value = oDv(iLine)(57).ToString.Trim
            End If
            If Not (oDv(iLine)(58).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_REQID").Value = oDv(iLine)(58).ToString.Trim
            End If
            If Not (oDv(iLine)(59).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CREDITCARDNO").Value = oDv(iLine)(59).ToString.Trim
            End If
            If Not (oDv(iLine)(60).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_CONTACTNO").Value = oDv(iLine)(60).ToString.Trim
            End If
            If Not (oDv(iLine)(61).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_ZAKATAGENCYID").Value = oDv(iLine)(61).ToString.Trim
            End If
            If Not (oDv(iLine)(62).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_BOOKINGID").Value = oDv(iLine)(62).ToString.Trim
            End If
            If Not (oDv(iLine)(63).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_COVERNOTENO").Value = oDv(iLine)(63).ToString.Trim
            End If
            If Not (oDv(iLine)(64).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_EMAIL").Value = oDv(iLine)(64).ToString.Trim
            End If
            If Not (oDv(iLine)(65).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_VEHINSURANCE").Value = oDv(iLine)(65).ToString.Trim
            End If
            If Not (oDv(iLine)(66).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWORKERINSURANCE").Value = oDv(iLine)(66).ToString.Trim
            End If
            If Not (oDv(iLine)(67).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_NEW_PASSPORT_NO").Value = oDv(iLine)(67).ToString.Trim
            End If
            If Not (oDv(iLine)(69).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_SECTIONCODE").Value = oDv(iLine)(69).ToString.Trim
            End If
            If Not (oDv(iLine)(70).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_FWID").Value = oDv(iLine)(70).ToString.Trim
            End If
            If Not (oDv(iLine)(71).ToString.Trim = String.Empty) Then
                oAPInvoice.UserFields.Fields.Item("U_TRANS_ID").Value = oDv(iLine)(71).ToString.Trim
            End If

            iCount = iCount + 1

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Line Items", sFuncName)

            If Not (oDv(iLine)(19).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(19).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PROCESSFEE", sFuncName)

                    sItemDesc = "processfee" & "-" & sServiceType

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PROCESSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''processfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PROCESSFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(19).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(20).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(20).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    sItemDesc = "passfee" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for PASSFEE", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for passfee", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PASSFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''passfee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for PASSFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(20).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(21).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(21).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for VISAFEE", sFuncName)

                    sItemDesc = "visafee" & "-" & sServiceType

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'VISAFEE'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''visafee'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for VISAFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(21).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(17).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(17).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    sItemDesc = "levifee_amount" & "-" & sServiceType

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for LEVIFEE_AMOUNT", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode in item mapping table for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for levifee_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'LEVIFEE_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''levifee_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for LEVIFEE_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(17).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If

            If Not (oDv(iLine)(10).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(10).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for SUMMONS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'SUMMONS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''summons_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for SUMMONS_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(10).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(11).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(11).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for PPZ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'PPZ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''ppz_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for PPZ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(11).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(12).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(12).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for JPJ_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'JPJ_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''jpj_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for JPJ_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(12).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(13).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(13).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for EHAK_AMOUNT", sFuncName)

                    'dtItemCode.DefaultView.RowFilter = "RevCostCode = 'EHAK_AMOUNT'"
                    'If dtItemCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "ItemCode ::''ehak_amount'' provided does not exist in SAP(Mapping Table)."
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'Else
                    '    sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    'End If
                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'COMPTEST_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''comptest_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting item for EHAK_AMOUNT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(13).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(14).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(14).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for INQ_AMT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'INQ_AMT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''inq_amt'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting for INQ_AMT", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(14).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If
            If Not (oDv(iLine)(15).ToString = String.Empty) Then
                If (CDbl(oDv(iLine)(15).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    sItemDesc = "agency_amount" & "-" & sServiceType
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas for " & sItemDesc, sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = '" & sItemDesc.ToUpper() & "'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No itemcode for " & sItemDesc, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting item code for agency_amount", sFuncName)

                        dtItemCode.DefaultView.RowFilter = Nothing
                        dtItemCode.DefaultView.RowFilter = "RevCostCode = 'AGENCY_AMOUNT'"
                        If dtItemCode.DefaultView.Count = 0 Then
                            sErrDesc = "ItemCode ::''agency_amount'' provided does not exist in SAP(Mapping Table)."
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        Else
                            sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                        End If
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for " & sItemDesc, sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(15).ToString.Trim)
                    If Not (p_oCompDef.sImmiGlAccount = String.Empty) Then
                        oAPInvoice.Lines.AccountCode = p_oCompDef.sImmiGlAccount
                    End If
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1

                End If
            End If

            If Not (oDv(iLine)(52).ToString.Trim = String.Empty) Then
                If (CDbl(oDv(iLine)(52).ToString.Trim() <> 0)) Then
                    bLineAdded = True
                    If iCount > 1 Then
                        oAPInvoice.Lines.Add()
                    End If
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing data for FIS_AMOUNT", sFuncName)

                    dtItemCode.DefaultView.RowFilter = "RevCostCode = 'FIS_AMOUNT'"
                    If dtItemCode.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode ::''fis_amount'' provided does not exist in SAP(Mapping Table)."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sItemCode = dtItemCode.DefaultView.Item(0)(0).ToString().Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(2).ToString().Trim()
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Inserting Item for INSFEE", sFuncName)

                    oAPInvoice.Lines.ItemCode = sItemCode
                    oAPInvoice.Lines.Quantity = 1
                    oAPInvoice.Lines.UnitPrice = CDbl(oDv(iLine)(23).ToString.Trim)
                    If sVatGroup <> "" Then
                        oAPInvoice.Lines.VatGroup = sVatGroup
                    End If
                    If Not (sCostCenter = String.Empty) Then
                        oAPInvoice.Lines.CostingCode = sCostCenter
                    End If
                    If Not (sCostCenter2 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode2 = sCostCenter2
                    End If
                    If Not (sCostCenter3 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode3 = sCostCenter3
                    End If
                    If Not (sCostCenter4 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode4 = sCostCenter4
                    End If
                    If Not (sCostCenter5 = String.Empty) Then
                        oAPInvoice.Lines.CostingCode5 = sCostCenter5
                    End If
                    iCount = iCount + 1
                End If
            End If

            If bLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding A/P Invoice Document", sFuncName)

                If oAPInvoice.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription
                    sErrDesc = sErrDesc.Replace("'", " ")
                    Console.WriteLine("Error while adding A/p invoice document/ " & sErrDesc)
                    sErrDesc = sErrDesc & " in funct. " & sFuncName

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding A/p invoice document", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("A/P invoice 2 created successfully", sFuncName)

                    Dim iDocNo, iDocEntry As Integer
                    iDocEntry = p_oCompany.GetNewObjectKey()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAPInvoice)

                    Dim sQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset

                    sQuery = "SELECT ""DocNum"" FROM ""OPCH"" WHERE ""DocEntry"" = '" & iDocEntry & "'"
                    oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("DocNum").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    sQuery = "UPDATE public.AB_REVENUEANDCOST SET ""A/P Invoice No2"" = '" & iDocNo & "', " & _
                             " SyncDate = NOW(),Status = 'SUCCESS',""Error Message"" = NULL WHERE ID = '" & sIntegId & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                    If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                End If
            End If
            
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAPInvoice_Second = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc.Replace("'", "") & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAPInvoice_Second = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create BP Code"
    Public Function CreateBP(ByVal sBPCode As String, ByRef sBPName As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "CreateBP()"
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim lRetCode, lErrCode As Long
        Dim sGroupCode As String = String.Empty
        Dim sPayTerms As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oBP = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Customer Code : " & sBPCode & ". Customer Name : " & sBPName, sFuncName)

            If oBP.GetByKey(sBPCode) = False Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("BP not exist in SAP", sFuncName)

                oBP.CardCode = sBPCode.ToUpper()
                oBP.CardName = sBPName
                oBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP.", sFuncName)
                lRetCode = oBP.Add

                If lRetCode <> 0 Then
                    p_oCompany.GetLastError(lErrCode, sErrDesc)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding BP failed.", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    p_oCompany.GetNewObjectCode(sBPCode)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
            CreateBP = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile("Function completed with ERROR:" & sErrDesc, sFuncName)
            CreateBP = RTN_ERROR
        Finally
            oBP = Nothing
        End Try

    End Function
#End Region
#Region "Create Credit Memo"
    Private Function CreateCreditNote(ByVal oDv As DataView, ByVal iLine As Integer, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateCreditNote"
        Dim sIntegId As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim sApInvNo As String = String.Empty
        Dim sApInvEntry As String = String.Empty
        Dim sPassportNo As String = String.Empty
        Dim oDs As New DataSet
        Dim oCreditNote As SAPbobsCOM.Documents
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim iCount As Integer = 1

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sIntegId = oDv(iLine)(0).ToString.Trim
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing datas based on id no " & sIntegId, sFuncName)

            sSQL = "SELECT ""A/P Invoice No"",passportno,new_passport_no FROM public.AB_REVENUEANDCOST WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL into Dataset" & sSQL, sFuncName)
            oDs = GetDataSet(sSQL)
            If oDs.Tables(0).Rows.Count > 0 Then
                sApInvNo = oDs.Tables(0).Rows(0).Item(0).ToString
                sPassportNo = oDs.Tables(0).Rows(0).Item(1).ToString
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating A/p Credit Memo Object", sFuncName)
            oCreditNote = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

            'sSQL = "SELECT * FROM ""OPCH"" WHERE ""DocNum"" = '" & sApInvNo & "' AND ""U_PASSPORTNO"" = '" & sPassportNo & "'"
            sSQL = "SELECT * FROM ""OPCH"" A INNER JOIN ""PCH1"" B ON B.""DocEntry"" = A.""DocEntry"" WHERE A.""DocNum"" = '" & sApInvNo & "' AND ""U_PASSPORTNO"" = '" & sPassportNo & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)

            oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sSQL)
            If Not (oRecordSet.BoF And oRecordSet.EoF) Then
                oRecordSet.MoveFirst()

                oCreditNote.CardCode = oRecordSet.Fields.Item("CardCode").Value
                oCreditNote.NumAtCard = oRecordSet.Fields.Item("NumAtCard").Value

                oCreditNote.UserFields.Fields.Item("U_MERCHANT_ID").Value = oRecordSet.Fields.Item("U_MERCHANT_ID").Value
                oCreditNote.UserFields.Fields.Item("U_AE_PAYMENTTYPE").Value = oRecordSet.Fields.Item("U_AE_PAYMENTTYPE").Value
                oCreditNote.UserFields.Fields.Item("U_SUMMONSID").Value = oRecordSet.Fields.Item("U_SUMMONSID").Value
                oCreditNote.UserFields.Fields.Item("U_SUMMONTYPE").Value = oRecordSet.Fields.Item("U_SUMMONTYPE").Value
                oCreditNote.UserFields.Fields.Item("U_OFFENCEDATE").Value = oRecordSet.Fields.Item("U_OFFENCEDATE").Value
                oCreditNote.UserFields.Fields.Item("U_OFFENDERNAME").Value = oRecordSet.Fields.Item("U_OFFENDERNAME").Value
                oCreditNote.UserFields.Fields.Item("U_OFFENDERIC").Value = oRecordSet.Fields.Item("U_OFFENDERIC").Value
                oCreditNote.UserFields.Fields.Item("U_VEHICLENO").Value = oRecordSet.Fields.Item("U_VEHICLENO").Value
                oCreditNote.UserFields.Fields.Item("U_LAWCODE2").Value = oRecordSet.Fields.Item("U_LAWCODE2").Value
                oCreditNote.UserFields.Fields.Item("U_LAWCODE3").Value = oRecordSet.Fields.Item("U_LAWCODE3").Value
                oCreditNote.UserFields.Fields.Item("U_JPJREVCODE").Value = oRecordSet.Fields.Item("U_JPJREVCODE").Value
                oCreditNote.UserFields.Fields.Item("U_REPLACETYPE").Value = oRecordSet.Fields.Item("U_REPLACETYPE").Value
                oCreditNote.UserFields.Fields.Item("U_USERID").Value = oRecordSet.Fields.Item("U_USERID").Value
                oCreditNote.UserFields.Fields.Item("U_IDNO").Value = oRecordSet.Fields.Item("U_IDNO").Value
                oCreditNote.UserFields.Fields.Item("U_COMPNO").Value = oRecordSet.Fields.Item("U_COMPNO").Value
                oCreditNote.UserFields.Fields.Item("U_ACCOUNTNO").Value = oRecordSet.Fields.Item("U_ACCOUNTNO").Value
                oCreditNote.UserFields.Fields.Item("U_BILLDATE").Value = oRecordSet.Fields.Item("U_BILLDATE").Value
                oCreditNote.UserFields.Fields.Item("U_CARREGNO").Value = oRecordSet.Fields.Item("U_CARREGNO").Value
                oCreditNote.UserFields.Fields.Item("U_PREPAIDACCTNO").Value = oRecordSet.Fields.Item("U_PREPAIDACCTNO").Value
                oCreditNote.UserFields.Fields.Item("U_LICENSECLASS").Value = oRecordSet.Fields.Item("U_LICENSECLASS").Value
                oCreditNote.UserFields.Fields.Item("U_REVENUECODE").Value = oRecordSet.Fields.Item("U_REVENUECODE").Value
                oCreditNote.UserFields.Fields.Item("U_VEHOWNERNAME").Value = oRecordSet.Fields.Item("U_VEHOWNERNAME").Value
                oCreditNote.UserFields.Fields.Item("U_EMPICNO").Value = oRecordSet.Fields.Item("U_EMPICNO").Value
                oCreditNote.UserFields.Fields.Item("U_EMPNAME").Value = oRecordSet.Fields.Item("U_EMPNAME").Value
                oCreditNote.UserFields.Fields.Item("U_PASSPORTNO").Value = oRecordSet.Fields.Item("U_PASSPORTNO").Value
                oCreditNote.UserFields.Fields.Item("U_APPLICANTNAME").Value = oRecordSet.Fields.Item("U_APPLICANTNAME").Value
                oCreditNote.UserFields.Fields.Item("U_SECTOR").Value = oRecordSet.Fields.Item("U_SECTOR").Value
                oCreditNote.UserFields.Fields.Item("U_PRINTSTATUS").Value = oRecordSet.Fields.Item("U_PRINTSTATUS").Value
                oCreditNote.UserFields.Fields.Item("U_AGCODE").Value = oRecordSet.Fields.Item("U_AGCODE").Value
                oCreditNote.UserFields.Fields.Item("U_PAYMODE").Value = oRecordSet.Fields.Item("U_PAYMODE").Value
                oCreditNote.UserFields.Fields.Item("U_ICNO").Value = oRecordSet.Fields.Item("U_ICNO").Value
                oCreditNote.UserFields.Fields.Item("U_ZAKATID").Value = oRecordSet.Fields.Item("U_ZAKATID").Value
                oCreditNote.UserFields.Fields.Item("U_REQID").Value = oRecordSet.Fields.Item("U_REQID").Value
                oCreditNote.UserFields.Fields.Item("U_CREDITCARDNO").Value = oRecordSet.Fields.Item("U_CREDITCARDNO").Value
                oCreditNote.UserFields.Fields.Item("U_CONTACTNO").Value = oRecordSet.Fields.Item("U_CONTACTNO").Value
                oCreditNote.UserFields.Fields.Item("U_ZAKATAGENCYID").Value = oRecordSet.Fields.Item("U_ZAKATAGENCYID").Value
                oCreditNote.UserFields.Fields.Item("U_BOOKINGID").Value = oRecordSet.Fields.Item("U_BOOKINGID").Value
                oCreditNote.UserFields.Fields.Item("U_COVERNOTENO").Value = oRecordSet.Fields.Item("U_COVERNOTENO").Value
                oCreditNote.UserFields.Fields.Item("U_EMAIL").Value = oRecordSet.Fields.Item("U_EMAIL").Value
                oCreditNote.UserFields.Fields.Item("U_VEHINSURANCE").Value = oRecordSet.Fields.Item("U_VEHINSURANCE").Value
                oCreditNote.UserFields.Fields.Item("U_FWORKERINSURANCE").Value = oRecordSet.Fields.Item("U_FWORKERINSURANCE").Value
                oCreditNote.UserFields.Fields.Item("U_NEW_PASSPORT_NO").Value = oRecordSet.Fields.Item("U_NEW_PASSPORT_NO").Value
                oCreditNote.UserFields.Fields.Item("U_SECTIONCODE").Value = oRecordSet.Fields.Item("U_SECTIONCODE").Value
                oCreditNote.UserFields.Fields.Item("U_FWID").Value = oRecordSet.Fields.Item("U_FWID").Value
                oCreditNote.UserFields.Fields.Item("U_TRANS_ID").Value = oRecordSet.Fields.Item("U_TRANS_ID").Value

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning line item", sFuncName)

                Do Until oRecordSet.EoF
                    If iCount > 1 Then
                        oCreditNote.Lines.Add()
                    End If

                    oCreditNote.Lines.ItemCode = oRecordSet.Fields.Item("ItemCode").Value
                    oCreditNote.Lines.Quantity = oRecordSet.Fields.Item("Quantity").Value
                    oCreditNote.Lines.UnitPrice = oRecordSet.Fields.Item("Price").Value
                    oCreditNote.Lines.VatGroup = oRecordSet.Fields.Item("VatGroup").Value
                    oCreditNote.Lines.CostingCode = oRecordSet.Fields.Item("OcrCode").Value
                    oCreditNote.Lines.CostingCode2 = oRecordSet.Fields.Item("OcrCode2").Value
                    oCreditNote.Lines.CostingCode3 = oRecordSet.Fields.Item("OcrCode3").Value
                    oCreditNote.Lines.CostingCode4 = oRecordSet.Fields.Item("OcrCode4").Value
                    oCreditNote.Lines.CostingCode5 = oRecordSet.Fields.Item("OcrCode5").Value

                    oCreditNote.Lines.BaseType = "18"
                    oCreditNote.Lines.BaseEntry = oRecordSet.Fields.Item("DocEntry").Value
                    oCreditNote.Lines.BaseLine = oRecordSet.Fields.Item("LineNum").Value

                    iCount = iCount + 1
                    oRecordSet.MoveNext()
                Loop

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Credit Memo", sFuncName)

                If oCreditNote.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription()
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while adding Credit Note", sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Note created successfully", sFuncName)

                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String
                    sQuery = "SELECT ""DocNum"" FROM ""ORPC"" WHERE ""DocEntry"" = '" & iDocEntry & "'"
                    oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        iDocNo = oRs.Fields.Item("DocNum").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

                    Console.WriteLine("Credit Note Created successfully :: " & iDocNo)

                    sQuery = "UPDATE public.AB_REVENUEANDCOST SET ""Credit Note No"" = '" & iDocNo & "', " & _
                             " SyncDate = NOW(),Status = 'SUCCESS',""Error Message"" = NULL WHERE ID = '" & sIntegId & "'"

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
                    If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                End If

            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateCreditNote = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            Dim sQuery As String
            sQuery = "UPDATE public.AB_REVENUEANDCOST SET Status = 'FAIL', ""Error Message"" = '" & sErrDesc.Replace("'", "") & "',SyncDate = NOW() " & _
                     " WHERE ID = '" & sIntegId & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            If ExecuteNonQuery(sQuery, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateCreditNote = RTN_ERROR
        End Try
    End Function
#End Region

End Module
