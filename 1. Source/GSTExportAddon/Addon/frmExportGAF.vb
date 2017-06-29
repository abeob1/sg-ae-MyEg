Public Class frmExportGAF

    Private Sub frmExportGAF_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Application.DoEvents()
        Me.BringToFront()
        Me.TopMost = True
    End Sub

    Private Sub frmBPList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.lblCaption.Text = "Export GAF file"

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If cbExportType.Text = "GST Audit File" Then
            SaveFileDialog1.Filter = "Xml files|*.xml"
            SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
            SaveFileDialog1.FileName = "GAF Export File.xml"
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor


                Dim dt As DataTable = Functions.DoQueryReturnDT("exec sp_B1Addon_GAF '" + cbFromDate.Value.ToString("yyyyMMdd") + "','" + cbToDate.Value.ToString("yyyyMMdd") + "'")
                If dt.Rows.Count > 0 Then
                    Dim xmlstr As String = dt.Rows(0).Item("xmlGAF").ToString
                    'If IO.File.Exists(SaveFileDialog1.FileName) Then
                    '    If MessageBox.Show("Do you want to overwrite the file?", "File Exist", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                    '        Me.Cursor = Cursors.Default
                    '        Return
                    '    End If
                    'End If
                    WriteFile(xmlstr, SaveFileDialog1.FileName)
                    If xmlstr = "" Then
                        MessageBox.Show("There's no data")
                    End If
                End If
                Me.Cursor = Cursors.Default
            End If
        Else
            SaveFileDialog1.Filter = "Online Submission files|*.txt"
            SaveFileDialog1.InitialDirectory = Environment.SpecialFolder.Desktop
            SaveFileDialog1.FileName = "GST Tap Return File.txt"
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor


                Dim dt As DataTable = Functions.DoQueryReturnDT("exec sp_B1Addon_GSTReturn '" + cbFromDate.Value.ToString("yyyyMMdd") + "','" + cbToDate.Value.ToString("yyyyMMdd") + "'")
                If dt.Rows.Count > 0 Then
                    Dim xmlstr As String = dt.Rows(0).Item("txtFilestr").ToString
                    'If IO.File.Exists(SaveFileDialog1.FileName) Then
                    '    If MessageBox.Show("Do you want to overwrite the file?", "File Exist", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                    '        Me.Cursor = Cursors.Default
                    '        Return
                    '    End If
                    'End If
                    WriteFile(xmlstr, SaveFileDialog1.FileName)
                    If xmlstr = "" Then
                        MessageBox.Show("There's no data")
                    End If
                End If
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub WriteFile(ByVal Str As String, ByVal FileName As String)
        Dim oWrite As IO.StreamWriter
        oWrite = IO.File.CreateText(FileName)
        oWrite.Write(Str)
        oWrite.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class
