Imports System.Text
Imports System.Xml
Imports System.Globalization

Public Class oXML
    Public Shared Function ToMultiXMLStringFromDS(NoOfDoc As Integer, ObjType() As String, ds() As DataSet, Optional ByVal GetFieldEmpty As Boolean = False) As String
        Try
            If ObjType.Length <> NoOfDoc Or ds.Length <> NoOfDoc Then
                Return "No. of Doc. different in length of array"
            End If
            'Dim gf As New GeneralFunctions()
            Dim XmlString As New StringBuilder()
            Dim writer As XmlWriter = XmlWriter.Create(XmlString)
            writer.WriteStartDocument()
            If True Then
                writer.WriteStartElement("BOM")
                If True Then
                    For i As Integer = 0 To NoOfDoc - 1
                        writer.WriteStartElement("BO")
                        If True Then
                            '#Region "write ADMINFO_ELEMENT"
                            writer.WriteStartElement("AdmInfo")
                            If True Then
                                writer.WriteStartElement("Object")
                                If True Then
                                    writer.WriteString(ObjType(i))
                                End If
                                writer.WriteEndElement()
                            End If
                            writer.WriteEndElement()
                            '#End Region

                            '#Region "Header&Line XML"
                            For Each dt As DataTable In ds(i).Tables
                                If dt.Rows.Count > 0 Then
                                    writer.WriteStartElement(dt.TableName.ToString(CultureInfo.InvariantCulture))
                                    If True Then
                                        For Each row As DataRow In dt.Rows
                                            writer.WriteStartElement("row")
                                            If True Then
                                                For Each column As DataColumn In dt.Columns
                                                    If column.DefaultValue.ToString() <> "xx_remove_xx" Then
                                                        'Force datetime format follow SQL 
                                                        If column.DataType Is GetType(DateTime) Then
                                                            Dim dateTime As DateTime
                                                            dateTime = row(column)
                                                            If GetFieldEmpty Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(IIf(IsDBNull(dateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), "", dateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                                                                writer.WriteEndElement()
                                                            Else
                                                                If row(column).ToString() <> "" Then
                                                                    writer.WriteStartElement(column.ColumnName)
                                                                    writer.WriteString(row(column).ToString)
                                                                    writer.WriteEndElement()
                                                                End If
                                                            End If
                                                        ElseIf column.DataType Is GetType(Date) Then
                                                            Dim d As Date
                                                            d = row(column)

                                                            If GetFieldEmpty Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(IIf(IsDBNull(d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), "", d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                                                                writer.WriteEndElement()
                                                            Else
                                                                If row(column).ToString() <> "" Then
                                                                    writer.WriteStartElement(column.ColumnName)
                                                                    writer.WriteString(row(column).ToString)
                                                                    writer.WriteEndElement()
                                                                End If
                                                            End If
                                                        Else
                                                            'Write Tag
                                                            If GetFieldEmpty Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(IIf(IsDBNull(row(column)), "", row(column)))
                                                                writer.WriteEndElement()
                                                            Else
                                                                If row(column).ToString() <> "" Then
                                                                    writer.WriteStartElement(column.ColumnName)
                                                                    writer.WriteString(row(column).ToString)
                                                                    writer.WriteEndElement()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If
                                            writer.WriteEndElement()
                                        Next
                                    End If
                                    writer.WriteEndElement()
                                End If
                            Next
                            '#End Region
                        End If
                        writer.WriteEndElement()
                    Next
                End If
                writer.WriteEndElement()
            End If
            writer.WriteEndDocument()

            writer.Flush()

            Return XmlString.ToString()
        Catch ex As Exception
            Return ex.ToString()
        End Try
    End Function
    Public Shared Function ToXMLStringFromDS(ObjType As String, ds As DataSet, Optional ByVal GetFieldEmpty As Boolean = False) As String
        Try
            'Dim gf As New GeneralFunctions()
            Dim XmlString As New StringBuilder()
            Dim writer As XmlWriter = XmlWriter.Create(XmlString)
            writer.WriteStartDocument()
            If True Then
                writer.WriteStartElement("BOM")
                If True Then
                    writer.WriteStartElement("BO")
                    If True Then
                        '#Region "write ADMINFO_ELEMENT"
                        writer.WriteStartElement("AdmInfo")
                        If True Then
                            writer.WriteStartElement("Object")
                            If True Then
                                writer.WriteString(ObjType)
                            End If
                            writer.WriteEndElement()
                        End If
                        writer.WriteEndElement()
                        '#End Region

                        '#Region "Header&Line XML"
                        For Each dt As DataTable In ds.Tables
                            If dt.Rows.Count > 0 Then
                                writer.WriteStartElement(dt.TableName.ToString(CultureInfo.InvariantCulture))
                                If True Then
                                    For Each row As DataRow In dt.Rows
                                        writer.WriteStartElement("row")
                                        If True Then
                                            For Each column As DataColumn In dt.Columns
                                                If column.DefaultValue.ToString() <> "xx_remove_xx" Then
                                                    'Force datetime format follow SQL 
                                                    If column.DataType Is GetType(DateTime) And Not IsDBNull(row(column)) Then
                                                        Dim dateTime As DateTime
                                                        dateTime = row(column)
                                                        If GetFieldEmpty Then
                                                            writer.WriteStartElement(column.ColumnName)
                                                            writer.WriteString(IIf(IsDBNull(dateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), "", dateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                                                            writer.WriteEndElement()
                                                        Else
                                                            If row(column).ToString() <> "" Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(row(column).ToString)
                                                                writer.WriteEndElement()
                                                            End If
                                                        End If
                                                    ElseIf column.DataType Is GetType(Date) And Not IsDBNull(row(column)) Then
                                                        Dim d As Date
                                                        d = row(column)

                                                        If GetFieldEmpty Then
                                                            writer.WriteStartElement(column.ColumnName)
                                                            writer.WriteString(IIf(IsDBNull(d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), "", d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                                                            writer.WriteEndElement()
                                                        Else
                                                            If row(column).ToString() <> "" Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(row(column).ToString)
                                                                writer.WriteEndElement()
                                                            End If
                                                        End If
                                                    Else
                                                        'Write Tag
                                                        If GetFieldEmpty Then
                                                            writer.WriteStartElement(column.ColumnName)
                                                            writer.WriteString(IIf(IsDBNull(row(column)), "", row(column)))
                                                            writer.WriteEndElement()
                                                        Else
                                                            If row(column).ToString() <> "" Then
                                                                writer.WriteStartElement(column.ColumnName)
                                                                writer.WriteString(row(column).ToString)
                                                                writer.WriteEndElement()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                        writer.WriteEndElement()
                                    Next
                                End If
                                writer.WriteEndElement()
                            End If
                        Next
                        '#End Region
                    End If
                    writer.WriteEndElement()
                End If
                writer.WriteEndElement()
            End If
            writer.WriteEndDocument()

            writer.Flush()

            Return XmlString.ToString()
        Catch ex As Exception
            Return ex.ToString()
        End Try
    End Function
    Public Shared Function CreateMarketingDocument(ByVal strXml As String, DocType As String, Optional DocEntry As String = "") As String
        Try
            Dim sStr As String = ""
            Dim lErrCode As Integer
            Dim sErrMsg As String
            Dim oDocment
            'Select Case DocType
            '    Case "30"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.JournalEntries)
            '    Case "97"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.SalesOpportunities)
            '    Case "191"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.ServiceCalls)
            '    Case "33"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.Contacts)
            '    Case "221"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.Attachments2)
            '    Case "2"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.BusinessPartners)
            '    Case "53"
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.SalesPersons)
            '    Case Else
            '        oDocment = DirectCast(oDocment, SAPbobsCOM.Documents)
            'End Select

            PublicVariable.oCompany.XMLAsString = True
            oDocment = PublicVariable.oCompany.GetBusinessObjectFromXML(strXml, 0)

            If DocEntry <> "" Then
                If oDocment.GetByKey(DocEntry) Then
                    oDocment.Browser.ReadXML(strXml, 0)
                    lErrCode = oDocment.Update()
                Else
                    lErrCode = oDocment.Add()
                End If
            Else
                lErrCode = oDocment.Add()
            End If


            If lErrCode <> 0 Then
                PublicVariable.oCompany.GetLastError(lErrCode, sErrMsg)
                Return sErrMsg
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Public Shared Function GetMarketingDocument(DocType As String, DocEntry As String) As String
        Try
            Dim sStr As String = ""
            Dim lErrCode As Integer
            Dim sErrMsg As String
            Dim oDocment
            Select Case DocType
                Case "30"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.JournalEntries)
                Case "97"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.SalesOpportunities)
                Case "191"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.ServiceCalls)
                Case "33"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Contacts)
                Case "221"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Attachments2)
                Case "2"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.BusinessPartners)
                Case "171" 'Employee
                    oDocment = DirectCast(oDocment, SAPbobsCOM.ContactEmployees)
                Case "206" '
                    oDocment = DirectCast(oDocment, SAPbobsCOM.UserObjectsMD)
                Case "28"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.IJournalVouchers)
                Case "25"
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Deposit)
                Case Else
                    oDocment = DirectCast(oDocment, SAPbobsCOM.Documents)

            End Select
            If PublicVariable.oCompany.Connected = False Then
                Dim MyArr As Array
                MyArr = System.Configuration.ConfigurationSettings.AppSettings.Get("LocalConnection").ToString.Split(";")

                PublicVariable.oCompany.CompanyDB = MyArr(0).ToString()
                PublicVariable.oCompany.UserName = MyArr(1).ToString()
                PublicVariable.oCompany.Password = MyArr(2).ToString()
                PublicVariable.oCompany.Server = MyArr(3).ToString()
                PublicVariable.oCompany.DbUserName = MyArr(4).ToString()
                PublicVariable.oCompany.DbPassword = MyArr(5).ToString()
                PublicVariable.oCompany.LicenseServer = MyArr(6)
                PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008

                lErrCode = PublicVariable.oCompany.Connect
                If lErrCode <> 0 Then
                    PublicVariable.oCompany.GetLastError(lErrCode, sErrMsg)
                    Functions.WriteLog("SystemInitial:" + sErrMsg)
                    Return "♠Error♠ : " + sErrMsg
                End If
            End If
            PublicVariable.oCompany.XMLAsString = True
            PublicVariable.oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ValidNodesOnly
           
            If DocType = "25" Then
                Dim sDeposit As SAPbobsCOM.DepositsService
                sDeposit = PublicVariable.oCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService)
                Dim oDepositsParams As SAPbobsCOM.DepositsParams
                oDepositsParams = sDeposit.GetDepositList()
                Dim oDeposit As SAPbobsCOM.Deposit
                Dim index As Integer = DocEntry - 1
                oDeposit = sDeposit.GetDeposit(oDepositsParams.Item(index))
                sStr = oDeposit.ToXMLString()
                Return sStr
            Else
                oDocment = PublicVariable.oCompany.GetBusinessObject(DocType)
                If oDocment.GetByKey(DocEntry) Then
                    oDocment.SaveXML(sStr)
                    Return sStr
                Else
                    Return "♠Error♠: docentry not found"
                End If
            End If

            
        Catch ex As Exception
            Return "♠Error♠ : " + ex.ToString
        End Try
    End Function

    

    Public Shared Function DeleteColumn(KeepColumns As String, dt As DataTable) As DataTable
        Dim dt1 As DataTable = dt.Copy
        For Each col As DataColumn In dt1.Columns
            If Not KeepColumns.Contains(col.ColumnName) Then
                dt.Columns.Remove(col.ColumnName)
            End If
        Next
        Return dt
    End Function


    


End Class
