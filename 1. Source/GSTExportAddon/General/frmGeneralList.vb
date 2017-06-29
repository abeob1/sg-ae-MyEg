Public Class frmGeneralList
    Public Direction As Integer = 1
    '1: LIVE DB
    '2: DMS DB
    '3: HQ DB

    Public QueryStr As String
    Private Sub frmBPList_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim dt As DataTable
        dt = Functions.SAP_Local_RunQuery(QueryStr)

        GridStep1.DataSource = dt
        GridStep1.RetrieveStructure()
        GridStep1.ColumnAutoResize = True



        'GridStep1.RootTable.Columns.Item(1).FilterRowComparison = Janus.Windows.GridEX.ConditionOperator.Contains
        For i As Integer = 0 To GridStep1.RootTable.Columns.Count - 1
            GridStep1.RootTable.Columns(i).FilterRowComparison = Janus.Windows.GridEX.ConditionOperator.Contains
            If GridStep1.RootTable.Columns(i).FormatString = "c" Then
                GridStep1.RootTable.Columns(i).FormatString = "n0"
                GridStep1.RootTable.Columns(i).CellStyle.TextAlignment = Janus.Windows.GridEX.TextAlignment.Far
            End If
            GridStep1.RootTable.Columns(i).CellStyle.Font = New Font("Myanmar3", 10)

        Next
        GridStep1.RootTable.Columns(0).FilterRowComparison = Janus.Windows.GridEX.ConditionOperator.BeginsWith

    End Sub

    Private Sub GridStep1_DoubleClick(sender As Object, e As System.EventArgs) Handles GridStep1.DoubleClick
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub btnExport_Click(sender As System.Object, e As System.EventArgs) Handles btnExport.Click
        Functions.ExportGridEx(GridStep1)
    End Sub
End Class
