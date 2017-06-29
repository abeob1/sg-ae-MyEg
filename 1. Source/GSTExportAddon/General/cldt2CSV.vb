Public Class cldt2CSV
    Public Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String)
        DataTable2CSV(table, filename, vbTab)
    End Sub
    Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String, _
        ByVal sepChar As String)
        Dim writer As System.IO.StreamWriter
        Try
            writer = New System.IO.StreamWriter(filename)

            ' first write a line with the columns name
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            For Each col As DataColumn In table.Columns
                builder.Append(sep).Append(col.ColumnName)
                sep = sepChar
            Next
            writer.WriteLine(builder.ToString())

            ' then write all the rows
            For Each row As DataRow In table.Rows
                sep = ""
                builder = New System.Text.StringBuilder

                For Each col As DataColumn In table.Columns
                    builder.Append(sep).Append(row(col.ColumnName))
                    sep = sepChar
                Next
                writer.WriteLine(builder.ToString())
            Next
        Finally
            If Not writer Is Nothing Then writer.Close()
        End Try
    End Sub
End Class
