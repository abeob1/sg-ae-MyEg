Imports System.Threading

Public Class fOutgoingPaymentEvents
    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
    End Sub

    Private Sub Handle_SBO_DataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.BeforeAction = False Then
                Select Case BusinessObjectInfo.Type
                    Case "46" 'Outgoing
                        Dim pmnt As SAPbobsCOM.Payments
                        pmnt = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                        Dim oinvoice As SAPbobsCOM.Documents
                        oinvoice = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            If pmnt.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                For i As Integer = 0 To pmnt.Invoices.Count - 1
                                    pmnt.Invoices.SetCurrentLine(i)
                                    If pmnt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice Then
                                        If oinvoice.GetByKey(pmnt.Invoices.DocEntry) Then
                                            If oinvoice.UserFields.Fields.Item("U_BadDebt").Value.ToString = "Y" Then
                                                Dim ret As String

                                                ret = CreateJEOutgoing(oinvoice.DocNum, pmnt.DocDate.ToString("yyyyMMdd"), CStr(pmnt.DocNum), pmnt.Invoices.SumApplied)
                                                If ret <> "" Then
                                                    SBO_Application.MessageBox("Bad Debt Reverse Error: " + ret)
                                                Else
                                                    SBO_Application.MessageBox("Bad Debt Reversed!")
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            If pmnt.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                If pmnt.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    Dim cl As New clJE
                                    Dim ret As String

                                    ret = cl.CancelJEFromOutgoing(pmnt.DocNum)
                                    If ret <> "" Then
                                        SBO_Application.MessageBox("Bad Debt Reversed Error: " + ret)
                                    Else
                                        SBO_Application.MessageBox("Bad Debt Reversed is cancelled!")
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

    Public Function CreateJEOutgoing(InvoiceNum As String, RefDate As String, OutGoingNum As String, PaidAmount As Double) As String
        'Cr OUTPUT tax, DR Relief recovery
        Try
            Dim DebAct As String, CreAct As String, Amount As String, TaxCode As String
            Dim dt As DataTable
            Dim str As String = ""
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "Call sp_B1Addon_BadDebtReverse_AP ('" + InvoiceNum + "'," + CStr(PaidAmount) + ")"
                dt = Functions.Hana_RunQuery(str)
            Else
                str = "exec sp_B1Addon_BadDebtReverse_AP '" + InvoiceNum + "'," + CStr(PaidAmount)
                dt = Functions.DoQueryReturnDT(str)
            End If


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
            dtHeader = InsertIntoOJDT(dtHeader, RefDate, "Y", "", "")
            dtLine = InsertIntoJDT1(dtLine, DebAct, CreAct, Amount, TaxCode, InvoiceNum, True, OutGoingNum)
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
    Private Function InsertIntoOJDT(dt As DataTable, RefDate As String, IsBadDebt As String, IsContraPayment As String, Remark As String) As DataTable
        Dim drNew As DataRow = dt.NewRow
        drNew("RefDate") = RefDate
        drNew("Memo") = Remark
        drNew("U_BadDebt") = IsBadDebt
        drNew("U_ContraPayment") = IsContraPayment
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
End Class
