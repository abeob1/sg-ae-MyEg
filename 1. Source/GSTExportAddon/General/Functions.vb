Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports System.Data.Odbc
Imports Sap.Data.Hana
Imports SAPbobsCOM

Public Class Functions
    Public Shared Sub WriteLog(ByVal Str As String)
        Dim oWrite As IO.StreamWriter
        Dim FilePath As String
        FilePath = Application.StartupPath + "\logfile.txt"

        If IO.File.Exists(FilePath) Then
            oWrite = IO.File.AppendText(FilePath)
        Else
            oWrite = IO.File.CreateText(FilePath)
        End If
        oWrite.Write(Now.ToString() + ":" + Str + vbCrLf)
        oWrite.Close()
    End Sub
    Public Shared Function ConvertRSSAPDT(RS As SAPbobsCOM.Recordset) As SAPbouiCOM.DataTable
        Dim returndt As SAPbouiCOM.DataTable
        Dim ColCount As Integer
        'add column
        For ColCount = 0 To RS.Fields.Count - 1
            returndt.Columns.Add(RS.Fields.Item(ColCount).Name, RS.Fields.Item(ColCount).Type, 1)
        Next
        returndt.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_DataOnly, RS.GetFixedXML(RecordsetXMLModeEnum.rxmData))

        ''add row
        'Do Until RS.EoF

        '    returndt.Rows.Add()
        '    'populate each column in the row we're creating
        '    For ColCount = 0 To RS.Fields.Count - 1

        '        NewRow.Item(RS.Fields.Item(ColCount).Name) = RS.Fields.Item(ColCount).Value

        '    Next

        '    'Add the row to the datatable
        '    dtTable.Rows.Add(NewRow)

        '    RS.MoveNext()
        'Loop

        Return returndt
    End Function
    Public Shared Function ConvertDataTable(dt As SAPbouiCOM.DataTable) As DataTable
        Dim dtreturn As DataTable = New DataTable

        'add columns to ado from sap
        For i As Integer = 0 To dt.Columns.Count - 1
            Dim dataType As String = "System."
            Select Case dt.Columns.Item(i).Type
                Case SAPbobsCOM.BoFieldTypes.db_Alpha
                    dataType = dataType & "String"
                Case SAPbobsCOM.BoFieldTypes.db_Date
                    dataType = dataType & "DateTime"
                Case SAPbobsCOM.BoFieldTypes.db_Float
                    dataType = dataType & "Double"
                Case SAPbobsCOM.BoFieldTypes.db_Memo
                    dataType = dataType & "String"
                Case SAPbobsCOM.BoFieldTypes.db_Numeric
                    dataType = dataType & "Decimal"
                Case Else
                    dataType = dataType & "String"
            End Select

            dtreturn.Columns.Add(dt.Columns.Item(i).Name, System.Type.GetType(dataType))
        Next
        'looping row in sap table
        For i As Integer = 0 To dt.Rows.Count - 1
            'looping column in sap table
            Dim dr As DataRow = dtreturn.NewRow
            For j As Integer = 0 To dt.Columns.Count - 1
                dr(dt.Columns.Item(j).Name) = dt.GetValue(dt.Columns.Item(j).Name, i)
            Next
            dtreturn.Rows.Add(dr)
        Next
        Return dtreturn
    End Function

    
    Public Shared Function DoQueryReturnDT(ByVal query As String) As DataTable
        Dim dt As DataTable
        Dim oRecordSet As SAPbobsCOM.Recordset

        'query = "select * from oitm"
        oRecordSet = PublicVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oRecordSet.DoQuery(query)
            If oRecordSet.RecordCount > 0 Then
                dt = ConvertRS2DT(oRecordSet)
                Return dt
            Else
                Return Nothing
            End If
        Catch ex As Exception
            WriteLog("query: " + query + "Error: " + ex.ToString)
            Return Nothing
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try

    End Function
    Public Shared Function ConvertRS2DT(ByVal RS As SAPbobsCOM.Recordset) As DataTable
        Dim dtTable As New DataTable
        Dim NewCol As DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer
        Try
            For ColCount = 0 To RS.Fields.Count - 1
                Dim dataType As String = "System."
                Select Case RS.Fields.Item(ColCount).Type
                    Case SAPbobsCOM.BoFieldTypes.db_Alpha
                        dataType = dataType & "String"
                    Case SAPbobsCOM.BoFieldTypes.db_Date
                        dataType = dataType & "DateTime"
                    Case SAPbobsCOM.BoFieldTypes.db_Float
                        dataType = dataType & "Double"
                    Case SAPbobsCOM.BoFieldTypes.db_Memo
                        dataType = dataType & "String"
                    Case SAPbobsCOM.BoFieldTypes.db_Numeric
                        dataType = dataType & "Decimal"
                    Case Else
                        dataType = dataType & "String"
                End Select

                NewCol = New DataColumn(RS.Fields.Item(ColCount).Name, System.Type.GetType(dataType))
                dtTable.Columns.Add(NewCol)
            Next

            Do Until RS.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To RS.Fields.Count - 1

                    NewRow.Item(RS.Fields.Item(ColCount).Name) = RS.Fields.Item(ColCount).Value

                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)

                RS.MoveNext()
            Loop
            Return dtTable
        Catch ex As Exception
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Return Nothing
        End Try
    End Function
    Public Shared Function SaveTextToFile(ByVal strData As String, _
     ByVal FullPath As String, _
       Optional ByVal ErrInfo As String = "") As Boolean

        Dim Contents As String
        Dim bAns As Boolean = False
        Dim objReader As StreamWriter
        Try


            objReader = New StreamWriter(FullPath)
            objReader.Write(strData)
            objReader.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try
        Return bAns
    End Function
    Public Shared Function GetFileContents(ByVal FullPath As String, _
       Optional ByRef ErrInfo As String = "") As String

        If Not File.Exists(FullPath) Then
            Return ""
        End If
        Dim strContents As String
        Dim objReader As StreamReader
        Try

            objReader = New StreamReader(FullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
            Return strContents
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
    End Function

#Region "ADO"
    Public Shared Function Hana_OpenSQLConnection() As HanaConnection
        Dim SAPConnection As HanaConnection = New HanaConnection
        SAPConnection.ConnectionString = "Server=" + PublicVariable.oCompany.Server + ";UserID=SYSTEM;Password=" + PublicVariable.SAPPass + ";Current schema=" + PublicVariable.oCompany.CompanyDB
        Try
            SAPConnection.Open()
            Return SAPConnection
        Catch ex As Exception
            Functions.WriteLog(ex.Message)
            Return Nothing
        End Try

    End Function

    Public Shared Function ODBC_OpenSQLConnection() As OdbcConnection
        Const _strLoginName As String = "SYSTEM"

        Dim strConnectionString As String = String.Empty
        'Does NOT require to create an odbc connection in windows system

        If IntPtr.Size = 8 Then
            ' Do 64-bit stuff
            strConnectionString = String.Concat(strConnectionString, "Driver={HDBODBC};")
        Else
            ' Do 32-bit
            strConnectionString = String.Concat(strConnectionString, "Driver={HDBODBC32};") '
        End If

        strConnectionString = String.Concat(strConnectionString, "ServerNode=", PublicVariable.oCompany.Server, ";")
        strConnectionString = String.Concat(strConnectionString, "UID=", _strLoginName, ";")
        strConnectionString = String.Concat(strConnectionString, "PWD=", PublicVariable.SAPPass, ";")
        strConnectionString = String.Concat(strConnectionString, "DATABASENAME=", PublicVariable.oCompany.Server + ";CS=" + PublicVariable.oCompany.CompanyDB, ";")

        Dim SAPConnection As OdbcConnection = New OdbcConnection
        SAPConnection.ConnectionString = strConnectionString
        Try
            SAPConnection.Open()
            Return SAPConnection
        Catch ex As Exception
            Functions.WriteLog(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Shared Function ADO_OpenSQLConnection() As SqlConnection


        Dim SAPConnection As SqlConnection = New SqlConnection
        SAPConnection.ConnectionString = "server= " + PublicVariable.oCompany.Server + ";database=" + PublicVariable.oCompany.CompanyDB + " ;uid=" + PublicVariable.oCompany.DbUserName + "; pwd=" + PublicVariable.SAPPass + ";"
        Try
            SAPConnection.Open()
            Return SAPConnection
        Catch ex As Exception
            Functions.WriteLog(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Shared Function ADO_RunQuery(ByVal querystr As String) As DataTable
        Dim SAPConnection As SqlConnection = New SqlConnection
        Try
            SAPConnection = ADO_OpenSQLConnection()
            If Not IsNothing(SAPConnection) Then
                Dim MyCommand As SqlCommand = New SqlCommand(querystr, SAPConnection)
                MyCommand.CommandType = CommandType.Text
                Dim da As SqlDataAdapter = New SqlDataAdapter()
                Dim mytable As DataTable = New DataTable()
                da.SelectCommand = MyCommand
                da.SelectCommand.CommandTimeout = 240
                da.Fill(mytable)

                If mytable Is Nothing Then Return New DataTable
                Return mytable
            Else
                Return New DataTable
            End If
        Catch ex As Exception
            WriteLog(ex.Message + vbCrLf + querystr)
            Return New DataTable
        Finally
            If Not SAPConnection Is Nothing Then
                SAPConnection.Close()
            End If
        End Try
    End Function
    Public Shared Function ODBC_RunQuery(ByVal querystr As String) As DataTable
        Dim SAPConnection As OdbcConnection = New OdbcConnection
        Try
            SAPConnection = ODBC_OpenSQLConnection()
            If Not IsNothing(SAPConnection) Then
                Dim MyCommand As New OdbcCommand(querystr, SAPConnection)
                MyCommand.CommandType = CommandType.Text
                Dim da As OdbcDataAdapter = New OdbcDataAdapter()
                Dim mytable As DataTable = New DataTable()
                da.SelectCommand = MyCommand
                da.SelectCommand.CommandTimeout = 240
                da.Fill(mytable)

                If mytable Is Nothing Then Return New DataTable
                Return mytable
            Else
                Return New DataTable
            End If
        Catch ex As Exception
            WriteLog(ex.Message + vbCrLf + querystr)
            Return New DataTable
        Finally
            If Not SAPConnection Is Nothing Then
                SAPConnection.Close()
            End If
        End Try
    End Function
    Public Shared Function Hana_RunQuery(ByVal querystr As String) As DataTable
        Dim SAPConnection As HanaConnection = New HanaConnection
        Try
            SAPConnection = Hana_OpenSQLConnection()
            If Not IsNothing(SAPConnection) Then
                Dim MyCommand As New HanaCommand(querystr, SAPConnection)
                MyCommand.CommandType = CommandType.Text
                Dim da As HanaDataAdapter = New HanaDataAdapter()
                Dim mytable As DataTable = New DataTable()
                da.SelectCommand = MyCommand
                da.SelectCommand.CommandTimeout = 240
                da.Fill(mytable)

                If mytable Is Nothing Then Return New DataTable
                Return mytable
            Else
                Return New DataTable
            End If
        Catch ex As Exception
            WriteLog(ex.Message + vbCrLf + querystr)
            Return New DataTable
        Finally
            If Not SAPConnection Is Nothing Then
                SAPConnection.Close()
            End If
        End Try
    End Function
#End Region
End Class
