<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGeneralList
    Inherits GSTExport.SAP_FakeForm

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
        Dim GridStep1_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGeneralList))
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.GridStep1 = New Janus.Windows.GridEX.GridEX()
        Me.Panel5.SuspendLayout()
        CType(Me.GridStep1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.btnExport)
        Me.Panel5.Controls.Add(Me.btnCancel)
        Me.Panel5.Controls.Add(Me.btnOK)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(1, 441)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(978, 40)
        Me.Panel5.TabIndex = 4
        '
        'btnExport
        '
        Me.btnExport.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnExport.Location = New System.Drawing.Point(898, 6)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(69, 26)
        Me.btnExport.TabIndex = 5
        Me.btnExport.Text = "Export"
        Me.btnExport.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(86, 6)
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
        Me.btnOK.Location = New System.Drawing.Point(11, 6)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(69, 26)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "Choose"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'GridStep1
        '
        Me.GridStep1.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.[False]
        Me.GridStep1.AlternatingColors = True
        GridStep1_DesignTimeLayout.LayoutString = resources.GetString("GridStep1_DesignTimeLayout.LayoutString")
        Me.GridStep1.DesignTimeLayout = GridStep1_DesignTimeLayout
        Me.GridStep1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridStep1.EnterKeyBehavior = Janus.Windows.GridEX.EnterKeyBehavior.None
        Me.GridStep1.FilterMode = Janus.Windows.GridEX.FilterMode.Automatic
        Me.GridStep1.FilterRowFormatStyle.Font = New System.Drawing.Font("Myanmar3", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridStep1.FilterRowUpdateMode = Janus.Windows.GridEX.FilterRowUpdateMode.WhenValueChanges
        Me.GridStep1.Font = New System.Drawing.Font("Myanmar3", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GridStep1.GridLineColor = System.Drawing.Color.Black
        Me.GridStep1.GridLineStyle = Janus.Windows.GridEX.GridLineStyle.Solid
        Me.GridStep1.GroupByBoxVisible = False
        Me.GridStep1.Location = New System.Drawing.Point(1, 31)
        Me.GridStep1.Name = "GridStep1"
        Me.GridStep1.NewRowPosition = Janus.Windows.GridEX.NewRowPosition.BottomRow
        Me.GridStep1.OfficeColorScheme = Janus.Windows.GridEX.OfficeColorScheme.Custom
        Me.GridStep1.OfficeCustomColor = System.Drawing.SystemColors.ButtonShadow
        Me.GridStep1.RowHeaders = Janus.Windows.GridEX.InheritableBoolean.[True]
        Me.GridStep1.SelectionMode = Janus.Windows.GridEX.SelectionMode.MultipleSelection
        Me.GridStep1.Size = New System.Drawing.Size(978, 410)
        Me.GridStep1.TabIndex = 29
        Me.GridStep1.UseCompatibleTextRendering = False
        '
        'frmGeneralList
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(980, 482)
        Me.Controls.Add(Me.GridStep1)
        Me.Controls.Add(Me.Panel5)
        Me.Name = "frmGeneralList"
        Me.Controls.SetChildIndex(Me.Panel5, 0)
        Me.Controls.SetChildIndex(Me.GridStep1, 0)
        Me.Panel5.ResumeLayout(False)
        CType(Me.GridStep1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents GridStep1 As Janus.Windows.GridEX.GridEX
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExport As System.Windows.Forms.Button

End Class
