Imports System.Data.Common
Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Data.Common.DbProviderFactories
Imports System.Data.SqlClient
Public Class ExcelDrilling
    Dim con As System.Data.OleDb.OleDbConnection
    ' Dim ExcelConnectionStr As String

    Public Sub New(ByVal dbPath As String)
        Dim ExcelConnectionStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath + ";Extended Properties=""Excel 12.0 Xml;"""
        con = New System.Data.OleDb.OleDbConnection(ExcelConnectionStr)
    End Sub

    Public Function GetDataSQL(ByVal sqlQuery As String) As DataTable
        Try
            con.Open()
            Dim ExcelCommand As New System.Data.OleDb.OleDbCommand(sqlQuery, con)
            Dim Reader As System.Data.OleDb.OleDbDataReader
            Reader = ExcelCommand.ExecuteReader
            Dim dt As New DataTable
            dt.Load(Reader)
            con.Close()
            Return dt
        Catch ex As Exception
            con.Close()
            Return Nothing
        End Try
    End Function
    Public Function GetSheets() As DataTable
        Try
            con.Open()
            Dim dt As DataTable
            dt = con.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
            con.Close()
            Return dt
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function
End Class
