Imports System.Threading

Public Class fIncomingPaymentEvents
    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
    End Sub

    Private Sub Handle_SBO_DataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.BeforeAction = False Then
                Select Case BusinessObjectInfo.Type

                    Case "24" 'Incoming
                        Dim pmnt As SAPbobsCOM.Payments
                        pmnt = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                        Dim oinvoice As SAPbobsCOM.Documents
                        oinvoice = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                        Dim ReverseMechanism As String = "N"
                        Dim dt As DataTable
                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            dt = Functions.Hana_RunQuery("Select * from ""@GSTSETUP"" where ""Code""='ReverseMechanism'")
                        Else
                            dt = Functions.DoQueryReturnDT("Select * from [@GSTSETUP] where Code='ReverseMechanism'")
                        End If

                        If Not IsNothing(dt) Then
                            If dt.Rows.Count > 0 Then
                                ReverseMechanism = dt.Rows(0).Item("U_Value").ToString
                            End If
                        End If
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            If pmnt.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                For i As Integer = 0 To pmnt.Invoices.Count - 1
                                    pmnt.Invoices.SetCurrentLine(i)
                                    If pmnt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice Then
                                        If oinvoice.GetByKey(pmnt.Invoices.DocEntry) Then
                                            If oinvoice.UserFields.Fields.Item("U_BadDebt").Value.ToString = "Y" Then
                                                Dim ret As String

                                                ret = CreateJEIncoming(oinvoice.DocNum, pmnt.DocDate.ToString("yyyyMMdd"), CStr(pmnt.DocNum), pmnt.Invoices.SumApplied)
                                                If ret <> "" Then
                                                    SBO_Application.MessageBox("Bad Debt Reserve Error: " + ret)
                                                Else
                                                    SBO_Application.MessageBox("Bad Debt Reserved!")
                                                End If
                                            End If
                                            If ReverseMechanism = "Y" Then
                                                For j As Integer = 0 To oinvoice.Lines.Count - 1
                                                    oinvoice.Lines.SetCurrentLine(j)
                                                    If oinvoice.Lines.VatGroup = "RC" Then
                                                        ' oinvoice.Lines.NetTaxAmount
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            If pmnt.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                If pmnt.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    Dim ret As String

                                    ret = CancelJEFromIncoming(pmnt.DocNum)
                                    If ret <> "" Then
                                        SBO_Application.MessageBox("Bad Debt Reserved Error: " + ret)
                                    Else
                                        SBO_Application.MessageBox("Bad Debt Reserved is cancelled!")
                                    End If
                                End If
                            End If
                        End If
                End Select

            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DataEvent Event: " + ex.ToString, , True)
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
        dt.Columns.Add("BaseSum")
        Return dt
    End Function

#End Region
#Region "Insert into Table"
    Private Function InsertIntoOJDT(dt As DataTable, RefDate As String, IsBadDebt As String, IsContraPayment As String, Remark As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("RefDate") = RefDate
        drNew("Memo") = "Bad Debt Reverse Entry"
        'drNew("U_BadDebt") = "Y"
        'drNew("U_ContraPayment") = IsContraPayment
        dt.Rows.Add(drNew)
        Return dt
    End Function
    Private Function InsertIntoJDT1_1side(dt As DataTable, DrAmt As Decimal, CrAmt As Decimal, Account As String, BaseEntry As String, BaseSum As Decimal) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("Account") = Account
        drNew("Debit") = DrAmt
        drNew("Credit") = CrAmt
        drNew("U_InvoiceEntry") = BaseEntry
        drNew("BaseSum") = BaseSum
        dt.Rows.Add(drNew)

        Return dt
    End Function
    Private Function InsertIntoJDT1(dt As DataTable, DebAct As String, CreAct As String, Amount As String, TaxCode As String, InvoiceEntry As String, BaseSum As Decimal, Optional CrTax As Boolean = False, Optional Ref3Line As String = "") As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("Account") = DebAct
        'drNew("ShortName") = DebAct
        drNew("Debit") = Amount
        drNew("Credit") = 0
        drNew("Ref3Line") = Ref3Line
        If Not CrTax Then
            drNew("VatGroup") = TaxCode
            drNew("BaseSum") = BaseSum
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
            drNew("BaseSum") = BaseSum
        End If
        dt.Rows.Add(drNew)

        Return dt
    End Function



#End Region

    Public Function CreateJEIncoming(InvoiceNum As String, RefDate As String, IncomingNum As String, PaidAmount As Double) As String
        'Cr OUTPUT tax, DR Relief recovery
        Try
            Dim DebAct As String, CreAct As String, Amount As String, TaxCode As String, BaseSum As Decimal
            Dim dt As DataTable
            Dim str As String
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "call sp_B1Addon_BadDebtReverse ('" + InvoiceNum + "'," + CStr(PaidAmount) + ")"
                dt = Functions.Hana_RunQuery(str)
            Else
                str = "exec sp_B1Addon_BadDebtReverse '" + InvoiceNum + "'," + CStr(PaidAmount)
                dt = Functions.DoQueryReturnDT(str)
            End If
            
            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    DebAct = dt.Rows(0).Item("DrAct").ToString
                    CreAct = dt.Rows(0).Item("CrAct").ToString
                    TaxCode = dt.Rows(0).Item("TaxCode").ToString
                    Amount = dt.Rows(0).Item("Debit")
                    BaseSum = dt.Rows(0).Item("BaseSum")
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
            dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceNum, BaseSum, True, IncomingNum)
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

    Public Function CancelJEFromIncoming(IncomingNum As String) As String
        Try
            Dim ret As String = ""
            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.Hana_RunQuery("Select distinct ""TransId"" from ""JDT1"" T0 join ""OVTG"" T1 on T0.""VatGroup""=T1.""Code"" where T1.""Category""='O' and T0.""TransType""='30' and ifnull(T0.""Ref3Line"",'')='" + IncomingNum + "'")
            Else
                dt = Functions.DoQueryReturnDT("Select distinct TransID from JDT1 T0 join OVTG T1 on T0.VatGroup=T1.Code where T1.Category='O' and T0.TransType='30' and isnull(T0.Ref3Line,'')='" + IncomingNum + "'")
            End If

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
End Class
