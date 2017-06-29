Imports System.Threading

Public Class fAREvents
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    
    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
    End Sub
    
    Private Sub Handle_SBO_DataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If BusinessObjectInfo.BeforeAction = False Then
                Select Case BusinessObjectInfo.Type
                    Case "13" 'AR Invoice
                        Dim oinvoice As SAPbobsCOM.Documents
                        oinvoice = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            If oinvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                For i As Integer = 0 To oinvoice.Lines.Count - 1
                                    oinvoice.Lines.SetCurrentLine(i)
                                    If oinvoice.Lines.BaseType = 15 Then
                                        Dim oDeliveryNote As SAPbobsCOM.Documents
                                        oDeliveryNote = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                        If oDeliveryNote.GetByKey(oinvoice.Lines.BaseEntry) Then
                                            If oDeliveryNote.UserFields.Fields.Item("U_21Day").Value.ToString = "Y" Then
                                                Dim ret As String = ""
                                                ret = CreateJE21Day_Reverse(oDeliveryNote.UserFields.Fields.Item("U_21DayJE").Value.ToString, oinvoice.DocDate.ToString("yyyyMMdd"), oDeliveryNote.DocNum.ToString)
                                                If ret <> "" Then
                                                    SBO_Application.MessageBox("Error (Outstanding DO Posting Reverse): " + ret)
                                                Else
                                                    SBO_Application.MessageBox("Outstanding DO GST Posting is reversed!")
                                                End If
                                                Exit For
                                            End If

                                        End If
                                    End If
                                Next
                            End If
                        End If
                End Select

            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DataEvent Event: " + ex.ToString, , True)
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

    Public Function CreateJE21Day_Reverse(JE As String, PostingDate As String, DODocNum As String) As String
        Try


            Dim xmlstr As String = ""
            Dim ds As New DataSet
            Dim ret As String = ""


            Dim dtHeader As DataTable = BuildTableOJDT()
            Dim dtLine As DataTable = BuildTableJDT1()
            Dim ObjType As String = "30"

            dtHeader = InsertIntoOJDT(dtHeader, PostingDate, "Y", "Outstanding DO GST Posting(Cancelled)")
            Dim Total As Decimal = 0

            Dim dt As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.DoQueryReturnDT("Select ""Debit"",""Credit"",""Account"",""U_InvoiceEntry"" from ""JDT1"" Where ""TransId""=" + JE + " AND ""U_InvoiceEntry""='" + DODocNum + "';")
            Else
                dt = Functions.DoQueryReturnDT("Select Debit,Credit,Account,U_InvoiceEntry from JDT1 Where TransID=" + JE + " AND U_InvoiceEntry='" + DODocNum + "'")
            End If

            For i As Integer = 0 To dt.Rows.Count - 1
                If dt.Rows(i).Item("Credit") > 0 Then
                    dtLine = InsertIntoJDT1_1side(dtLine, dt.Rows(i).Item("Credit"), 0, dt.Rows(i).Item("Account").ToString, dt.Rows(i).Item("U_InvoiceEntry").ToString)
                End If
                Total = Total + dt.Rows(i).Item("Credit")
            Next

            Dim ContraAct As String = ""
            Dim dtsetup As DataTable
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dtsetup = Functions.DoQueryReturnDT("Select * from ""@GSTSETUP"" where ""Code""='DOAct';")
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

            dtLine = InsertIntoJDT1_1side(dtLine, 0, Total, ContraAct, "")

            dtHeader.TableName = "OJDT"
            dtLine.TableName = "JDT1"

            ds.Tables.Add(dtHeader.Copy)
            ds.Tables.Add(dtLine.Copy)
            xmlstr = oXML.ToXMLStringFromDS(ObjType, ds)
            ret = oXML.CreateMarketingDocument(xmlstr, ObjType)

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
                dt = Functions.DoQueryReturnDT("select ""DocEntry"" from ""ODLN"" where ""DocNum""='" + DocNum + "';")
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
