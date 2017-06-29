Public Class clJE
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
#Region "Insert into Table"
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

    

#End Region

#Region "Cancel Bad Debpt"
    
    
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

#Region "Create Bad Debt Entry"
    Public Function CreateBadDebt(oForm As SAPbouiCOM.Form, InvoiceType As String) As String
        Try
            Dim OneJE As String = "N"
            Dim dt As DataTable = Functions.DoQueryReturnDT("select U_Value from [@GSTSETUP] where code='OneJE'")
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
                                Ret = CreateJE(InvoiceEntry, RefDate, DebAct, CreAct, Amount, TaxCode, InvoiceType)
                                If Ret <> "" Then
                                    Return Ret
                                End If
                            Else
                                dtHeader = InsertIntoOJDT(dtHeader, RefDate, "Y", "", "")
                                dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceEntry)
                            End If

                        End If
                    End If
                End If
            Next
            If OneJE = "Y" Then
                If dtHeader.Rows.Count > 0 Then
                    Ret = CreateOneJE(dtHeader, dtLine, InvoiceType)
                    Return Ret
                End If
            End If
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
    Private Function CreateOneJE(dtHeader As DataTable, dtLine As DataTable, InvoiceType As String) As String
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
            oinvoice = PublicVariable.oCompany.GetBusinessObject(InvoiceType)

            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            PublicVariable.oCompany.StartTransaction()
            ret = oXML.CreateMarketingDocument(xmlstr, "30")
            If ret = "" Then
                Dim JENo As Integer = PublicVariable.oCompany.GetNewObjectKey
                For i As Integer = 0 To dtLine.Rows.Count - 1
                    If InvoiceType = "13" Then
                        If oinvoice.GetByKey(ReturnInvEntryFromInvNum(dtLine.Rows(i).Item("U_InvoiceEntry").ToString)) Then
                            oinvoice.UserFields.Fields.Item("U_BadDebt").Value = "Y"
                            oinvoice.UserFields.Fields.Item("U_BadDebtJE").Value = CStr(JENo)
                            ret = oinvoice.Update()
                            If ret <> "0" Then
                                ret = PublicVariable.oCompany.GetLastErrorDescription
                            Else
                                ret = ""
                            End If
                        End If
                    Else
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
                    End If
                    
                Next
            End If
            If PublicVariable.oCompany.InTransaction Then
                PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            Return ret
        Catch ex As Exception
            PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return ex.ToString
        End Try
    End Function
    Private Function CreateJE(InvoiceEntry As Integer, RefDate As String, DebAct As String, CreAct As String, Amount As String, TaxCode As String, InvoiceType As String) As String
        Try
            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""

            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim oinvoice As SAPbobsCOM.Documents
            oinvoice = PublicVariable.oCompany.GetBusinessObject(InvoiceType)

            Dim ObjType As String = "30"
            dtHeader = InsertIntoOJDT(dtHeader, RefDate, "", "", "")
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
                If InvoiceType = "13" Then
                    If oinvoice.GetByKey(ReturnInvEntryFromInvNum(InvoiceEntry)) Then
                        oinvoice.UserFields.Fields.Item("U_BadDebt").Value = "Y"
                        oinvoice.UserFields.Fields.Item("U_BadDebtJE").Value = CStr(JENo)
                        ret = oinvoice.Update()
                        If ret <> "0" Then
                            ret = PublicVariable.oCompany.GetLastErrorDescription
                        Else
                            ret = ""
                        End If
                    End If
                Else
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
                
            End If
            PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Return ret
        Catch ex As Exception
            PublicVariable.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return ex.ToString
        End Try
    End Function

#End Region
    Private Function ReturnInvEntryFromInvNum(DocNum As String) As String
        Try
            Dim dt As DataTable = Functions.DoQueryReturnDT("select DocEntry from OINV with(nolock) where DocNum='" + DocNum + "'")
            Return dt.Rows(0).Item("DocEntry").ToString
        Catch ex As Exception
            Return ""
        End Try
       
    End Function
    Private Function ReturnInvEntryFromInvNum_AP(DocNum As String) As String
        Try
            Dim dt As DataTable = Functions.DoQueryReturnDT("select DocEntry from OPCH with(nolock) where DocNum='" + DocNum + "'")
            Return dt.Rows(0).Item("DocEntry").ToString
        Catch ex As Exception
            Return ""
        End Try

    End Function
#Region "Create JE for Incoming to reserve Bad Debt"
    Public Function CreateJEIncoming(InvoiceNum As String, RefDate As String, IncomingNum As String, PaidAmount As Double) As String
        'Cr OUTPUT tax, DR Relief recovery
        Try
            Dim DebAct As String, CreAct As String, Amount As String, TaxCode As String
            Dim dt As DataTable
            Dim str As String = "exec sp_B1Addon_BadDebtReverse '" + InvoiceNum + "'," + CStr(PaidAmount)
            dt = Functions.DoQueryReturnDT(str)
            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    DebAct = dt.Rows(0).Item("DrAct").ToString
                    CreAct = dt.Rows(0).Item("CrAct").ToString
                    TaxCode = dt.Rows(0).Item("TaxCode").ToString
                    Amount = dt.Rows(0).Item("Debit")
                Else
                    Return "Record not found!"
                End If
            Else
                Return "Record not found!"
            End If


            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""

            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim ObjType As String = "30"
            dtHeader = InsertIntoOJDT(dtHeader, RefDate, "", "", "")
            dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceNum, True, IncomingNum)
            dtHeader.TableName = "OJDT"
            dtLine.TableName = "JDT1"

            ds.Tables.Add(dtHeader.Copy)
            ds.Tables.Add(dtLine.Copy)
            xmlstr = oXML.ToXMLStringFromDS(ObjType, ds)

            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)
            Return ret
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
    

    
    Public Function CancelJEFromOutgoing(OutgoingNum As String) As String
        Try
            Dim ret As String = ""
            Dim dt As DataTable = Functions.DoQueryReturnDT("Select distinct TransID from JDT1 T0 join OVTG T1 on T0.VatGroup=T1.Code where T1.Category='I' and T0.TransType='30' and isnull(T0.Ref3Line,'')='" + OutgoingNum + "'")
            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        ret = CancelJE(dt.Rows(i).Item("TransID"))
                        If ret <> "" Then
                            Return ret
                        End If
                    Next
                End If
            End If
            Return ret
        Catch ex As Exception
            Return ex.ToString
        End Try
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
            Dim dtsetup As DataTable = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] where code='ContraAct'")
            If dtsetup.Rows.Count > 0 Then
                ContraAct = dtsetup.Rows(0).Item("U_Value").ToString
            End If
            Dim dtInput As DataTable
            Dim dtOutput As DataTable

            dtInput = dt.Copy.Select("Category='Input'").CopyToDataTable
            dtOutput = dt.Copy.Select("Category='Output'").CopyToDataTable
            TotalInput = Convert.ToDecimal(dtInput.Compute("Sum(Balance)", String.Empty))
            TotalOutput = Convert.ToDecimal(dtOutput.Compute("Sum(Balance)", String.Empty))
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

            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)
            Return ret
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
#End Region
  End Class
