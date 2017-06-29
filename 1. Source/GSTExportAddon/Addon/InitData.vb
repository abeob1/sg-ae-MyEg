Imports System.IO

Public Class InitData
    
    Public Function CreateAllTaxCode()
        Dim str As String = ""
        Return ""
        '------------------PURCHASE TAX CODE ----------------------------
        str = TaxCode("TX", "Purchase with GST incurred at 6% and Directly attributable to taxable supplies", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("IM", "Import of goods with GST incurred", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("IS", "Imports under special scheme with no GST incurred (e.g. Approved Trader Scheme, ATMS Scheme)", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("BL", "Purchases with GST incurred but not claimable (Disallowance of Input Tax) (e.g. medical expenses for staff)", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("NR", "Purchase from non GST-registered supplier with no GST incurred", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcInputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        '-----------------SALES TAX CODE---------------------------------------------
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)
        str = TaxCode("", "", SAPbobsCOM.BoVatCategoryEnum.bovcOutputTax, Now.Date, 6)
        If str <> "" Then Functions.WriteLog(str)

        Return str
    End Function
    Public Function TaxCode(ByVal Code As String, ByVal Name As String, ByVal category As SAPbobsCOM.BoVatCategoryEnum, ByVal effectiveDate As Date, ByVal Rate As Decimal) As String
        Dim otx As SAPbobsCOM.VatGroups
        otx = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups)
        If Not otx.GetByKey(Code) Then
            otx.Code = Code
            otx.Name = Name
            otx.Category = category
            otx.VatGroups_Lines.Add()
            otx.VatGroups_Lines.Effectivefrom = effectiveDate
            otx.VatGroups_Lines.Rate = Rate

            Dim lRetCode As Integer
            Dim sErrMsg As String = ""

            otx.Add()
            If lRetCode <> 0 Then
                Return PublicVariable.oCompany.GetLastErrorDescription()
            End If
        End If
        Return ""
    End Function
    Public Function CreateUDT(ByVal tableName As String, ByVal tableDesc As String, ByVal tableType As SAPbobsCOM.BoUTBTableType) As String
        Dim oUdtMD As SAPbobsCOM.UserTablesMD = Nothing
        Try
            oUdtMD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            If oUdtMD.GetByKey(tableName) = False Then
                oUdtMD.TableName = tableName
                oUdtMD.TableDescription = tableDesc
                oUdtMD.TableType = tableType
                Dim lRetCode As Integer
                lRetCode = oUdtMD.Add
                If (lRetCode <> 0) Then
                    If (lRetCode = -2035) Then
                        Return "-2035"
                    End If
                    Return PublicVariable.oCompany.GetLastErrorDescription()

                End If

                'PublicVariable.oCompany.Disconnect()
                'GC.Collect()
                'Dim sCookie As String = PublicVariable.oCompany.GetContextCookie
                'Dim sConnectionContext As String
                'sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)
                'PublicVariable.oCompany.SetSboLoginContext(sConnectionContext)
                'If PublicVariable.oCompany.Connect() <> 0 Then
                '    Return PublicVariable.oCompany.GetLastErrorDescription()
                'End If

                Return ""
            Else
                Return ""
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdtMD)
            oUdtMD = Nothing
            GC.Collect()
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function CreateUDF(ByVal tableName As String, ByVal fieldName As String, ByVal desc As String, _
                              ByVal fieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal LinkTab As String, _
                              Optional SubType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None) As String
        Try
            Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
            oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUserFieldsMD.TableName = tableName
            oUserFieldsMD.Name = fieldName
            oUserFieldsMD.Description = desc
            oUserFieldsMD.Type = fieldType
            If Size <> 0 Then
                oUserFieldsMD.EditSize = Size
            End If

            oUserFieldsMD.SubType = SubType
            Dim lRetCode As Integer
            Dim sErrMsg As String = ""
            lRetCode = oUserFieldsMD.Add()
            If lRetCode <> 0 Then
                If (lRetCode = -2035 Or lRetCode = -1120) Then
                    Return CStr(lRetCode)
                End If
                Return PublicVariable.oCompany.GetLastErrorDescription()
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect()
            oUserFieldsMD = Nothing


            Return ""
        Catch ex As Exception
            Functions.WriteLog(ex.ToString)
            Return ex.Message

        End Try

    End Function

    Public Function CheckTableExists(TableName As String) As Boolean
        Dim oUdtMD As SAPbobsCOM.UserTablesMD = Nothing
        Dim ret As Boolean = False
        Try
            TableName = TableName.Replace("@", "")
            oUdtMD = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            If oUdtMD.GetByKey(TableName) Then
                ret = True
            Else
                ret = False
            End If

        Catch ex As Exception
            ret = False
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdtMD)
            oUdtMD = Nothing
            GC.Collect()
        End Try
        Return ret
        'Dim dt As DataTable
        'dt = Functions.DoQueryReturnDT("SELECT count(*) CountN FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" + TableName + "'")
        'If Not IsNothing(dt) Then
        '    If dt.Rows.Count > 0 Then
        '        If dt.Rows(0).Item("CountN") = 1 Then
        '            Return True
        '        End If
        '    End If
        'End If
        'Return False
    End Function
    Public Function CheckFieldExists(TableName As String, FieldName As String) As Boolean
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        Dim ret As Boolean = False
        Try

            FieldName = FieldName.Replace("U_", "")
            oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            Dim FieldID As Integer = getFieldidByName(TableName, FieldName)
            'TableName = TableName.Replace("@", "")
            If oUserFieldsMD.GetByKey(TableName, FieldID) Then
                ret = True
            Else
                ret = False
            End If

        Catch ex As Exception
            ret = False
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            oUserFieldsMD = Nothing
            GC.Collect()

        End Try

        Return ret
        'Dim dt As DataTable
        'dt = Functions.DoQueryReturnDT("SELECT count(*) CountN FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" + TableName + "' and COLUMN_NAME='" + FieldName + "'")
        'If Not IsNothing(dt) Then
        '    If dt.Rows.Count > 0 Then
        '        If dt.Rows(0).Item("CountN") = 1 Then
        '            Return True
        '        End If
        '    End If
        'End If
        'Return False
    End Function
    Private Function getFieldidByName(TableName As String, FieldName As String) As Integer
        Dim index As Integer = -1
        Dim ors As SAPbobsCOM.Recordset
        Try

            ors = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                ors.DoQuery("select ""FieldID"" from ""CUFD"" where ""TableID"" = '" + TableName + "' and ""AliasID"" = '" + FieldName + "';")
            Else
                ors.DoQuery("select FieldID from CUFD where TableID = '" + TableName + "' and AliasID = '" + FieldName + "'")
            End If

            If Not ors.EoF Then
                index = ors.Fields.Item("FieldID").Value
            End If
        Catch ex As Exception
            Return Nothing
        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors)
            ors = Nothing
            GC.Collect()
        End Try
        Return index
    End Function
    Public Function CheckStoreProcedureExists(spname As String) As String
        Try
            Dim dt As DataTable
            Dim result As String = ""
            '0: same
            '1: alter
            '2: create
            Dim st As String = ""
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                st = "SELECT count(*)  as ""CountN"" FROM SYS.PROCEDURES where SCHEMA_NAME='" + PublicVariable.oCompany.CompanyDB + "' AND PROCEDURE_NAME='" + spname.ToUpper + "'"
                dt = Functions.Hana_RunQuery(st)
            Else
                st = "SELECT count(*) CountN FROM sys.objects where type in (N'P',N'PC') and  object_id=object_ID(N'" + spname + "')"
                dt = Functions.DoQueryReturnDT(st)
            End If


            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("CountN") = 1 Then
                        Dim GetVersion As String = ""
                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            st = "call get_object_definition('" + PublicVariable.oCompany.CompanyDB + "','" + spname.ToUpper + "');"
                            ''dt = Functions.DoQueryReturnDT(st)
                            dt = Functions.Hana_RunQuery(st)
                            If dt.Rows(0).Item("OBJECT_CREATION_STATEMENT").ToString.Length >= 14 + spname.Length + 16 Then
                                GetVersion = dt.Rows(0).Item("OBJECT_CREATION_STATEMENT").ToString.Substring(14 + spname.Length + 8, 8)
                                'Functions.WriteLog("SP Version " + spname + ":" + GetVersion)
                                If PublicVariable.Version <> GetVersion Then
                                    result = " alter "
                                Else
                                    result = ""
                                End If
                            End If
                        Else
                            dt = Functions.DoQueryReturnDT("SELECT SUBSTRING(OBJECT_DEFINITION(OBJECT_ID('" + spname + "')),CHARINDEX('--**',OBJECT_DEFINITION(OBJECT_ID('" + spname + "')))+4,8)")
                            result = dt.Rows(0).Item(0).ToString
                        End If
                    Else
                        result = ""
                    End If
                Else
                    result = ""
                End If
            Else
                result = ""
            End If

            Return result
        Catch ex As Exception
            Functions.WriteLog("CheckStoreProcedureExists:" + ex.ToString)
            Return " create "
        End Try
    End Function
    Public Function CheckFunctionExists(fnname As String) As String
        Try
            Dim dt As DataTable
            Dim result As String = ""
            '0: same
            '1: alter
            '2: create
            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                dt = Functions.Hana_RunQuery("SELECT count(*) as ""CountN"" FROM sys.FUNCTIONS where  SCHEMA_NAME='" + PublicVariable.oCompany.CompanyDB + "' AND FUNCTION_NAME='" + fnname.ToUpper + "'")
            Else
                dt = Functions.DoQueryReturnDT("SELECT count(*) CountN FROM sys.objects where type in (N'FN') and  object_id=object_ID(N'" + fnname + "')")
            End If

            If Not IsNothing(dt) Then
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("CountN") = 1 Then
                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            dt = Functions.Hana_RunQuery("call get_object_definition('" + PublicVariable.oCompany.CompanyDB + "','" + fnname.ToUpper + "');")
                            If dt.Rows(0).Item("OBJECT_CREATION_STATEMENT").ToString.Length >= 14 + fnname.Length + 16 Then
                                Dim GetVersion As String = ""
                                GetVersion = dt.Rows(0).Item("OBJECT_CREATION_STATEMENT").ToString.Substring(14 + fnname.Length + 6, 8)
                                'Functions.WriteLog("FN Version " + fnname + ":" + GetVersion)
                                If PublicVariable.Version <> GetVersion Then
                                    result = " alter "
                                End If
                            End If
                        Else
                            dt = Functions.DoQueryReturnDT("sp_helptext '" + fnname + "'")
                            If dt.Rows(0).Item("Text").ToString.Length >= 16 Then
                                If PublicVariable.Version <> dt.Rows(0).Item("Text").ToString.Substring(4, 8) Then
                                    result = " alter "
                                End If
                            Else
                                result = " alter "
                            End If
                        End If



                    Else
                        result = " create "
                    End If
                Else
                    result = " create "
                End If
            Else
                result = " create "
            End If
            Return result
        Catch ex As Exception
            Functions.WriteLog("CheckFunctionExists:" + ex.ToString)
            Return " create "
        End Try
    End Function

    Public Function GetCrystalReportFile(ByVal RDOCCode As String, ByVal outFileName As String) As String
        Try
            If ocompany.Connected = True Then
                Dim oBlobParams As SAPbobsCOM.BlobParams = ocompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = RDOCCode
                Dim oBlob As SAPbobsCOM.Blob = ocompany.GetCompanyService().GetBlob(oBlobParams)
                Dim sContent As String = oBlob.Content
                Dim buf() As Byte = Convert.FromBase64String(sContent)
                Using oFile As New System.IO.FileStream(outFileName, System.IO.FileMode.Create)
                    oFile.Write(buf, 0, buf.Length)
                    oFile.Close()
                End Using
            Else
                Return "Not connected!"
            End If
        Catch ex As Exception
            Return ex.ToString
        End Try
        Return ""
    End Function
End Class
