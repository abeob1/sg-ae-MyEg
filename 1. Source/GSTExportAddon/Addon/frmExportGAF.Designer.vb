<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExportGAF
    Inherits GSTAddon.SAP_FakeForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.cbFromDate = New System.Windows.Forms.DateTimePicker()
        Me.MyLabel5 = New GSTAddon.myLabel()
        Me.cbToDate = New System.Windows.Forms.DateTimePicker()
        Me.MyLabel1 = New GSTAddon.myLabel()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.MyLabel2 = New GSTAddon.myLabel()
        Me.cbExportType = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(139, 148)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(69, 26)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(64, 148)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(69, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "Export"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'cbFromDate
        '
        Me.cbFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.cbFromDate.Location = New System.Drawing.Point(92, 37)
        Me.cbFromDate.Name = "cbFromDate"
        Me.cbFromDate.Size = New System.Drawing.Size(132, 23)
        Me.cbFromDate.TabIndex = 56
        '
        'MyLabel5
        '
        Me.MyLabel5.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel5.Caption = "From Date"
        Me.MyLabel5.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(6, 43)
        Me.MyLabel5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel5.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel5.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(127, 17)
        Me.MyLabel5.TabIndex = 55
        Me.MyLabel5.TabStop = False
        '
        'cbToDate
        '
        Me.cbToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.cbToDate.Location = New System.Drawing.Point(92, 67)
        Me.cbToDate.Name = "cbToDate"
        Me.cbToDate.Size = New System.Drawing.Size(132, 23)
        Me.cbToDate.TabIndex = 58
        '
        'MyLabel1
        '
        Me.MyLabel1.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel1.Caption = "To Date"
        Me.MyLabel1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(6, 73)
        Me.MyLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel1.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel1.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(127, 17)
        Me.MyLabel1.TabIndex = 57
        Me.MyLabel1.TabStop = False
        '
        'MyLabel2
        '
        Me.MyLabel2.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel2.Caption = "Export File"
        Me.MyLabel2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(6, 107)
        Me.MyLabel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel2.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel2.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(127, 17)
        Me.MyLabel2.TabIndex = 59
        Me.MyLabel2.TabStop = False
        '
        'cbExportType
        '
        Me.cbExportType.FormattingEnabled = True
        Me.cbExportType.Items.AddRange(New Object() {"GST Audit File", "GST Tap Return File"})
        Me.cbExportType.Location = New System.Drawing.Point(92, 100)
        Me.cbExportType.Name = "cbExportType"
        Me.cbExportType.Size = New System.Drawing.Size(132, 24)
        Me.cbExportType.TabIndex = 60
        Me.cbExportType.Text = "GST Audit File"
        '
        'frmExportGAF
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(268, 201)
        Me.Controls.Add(Me.cbExportType)
        Me.Controls.Add(Me.MyLabel2)
        Me.Controls.Add(Me.cbToDate)
        Me.Controls.Add(Me.MyLabel1)
        Me.Controls.Add(Me.cbFromDate)
        Me.Controls.Add(Me.MyLabel5)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Name = "frmExportGAF"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "GAF Export"
        Me.TopMost = True
        Me.Controls.SetChildIndex(Me.btnOK, 0)
        Me.Controls.SetChildIndex(Me.btnCancel, 0)
        Me.Controls.SetChildIndex(Me.MyLabel5, 0)
        Me.Controls.SetChildIndex(Me.cbFromDate, 0)
        Me.Controls.SetChildIndex(Me.MyLabel1, 0)
        Me.Controls.SetChildIndex(Me.cbToDate, 0)
        Me.Controls.SetChildIndex(Me.MyLabel2, 0)
        Me.Controls.SetChildIndex(Me.cbExportType, 0)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents cbFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents MyLabel5 As GSTAddon.myLabel
    Friend WithEvents cbToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents MyLabel1 As GSTAddon.myLabel
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MyLabel2 As GSTAddon.myLabel
    Friend WithEvents cbExportType As System.Windows.Forms.ComboBox

End Class
