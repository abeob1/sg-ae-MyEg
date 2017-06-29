<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
    Inherits GSTExport.SAP_FakeForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbSQLType = New System.Windows.Forms.ComboBox()
        Me.MyLabel8 = New GSTExport.myLabel()
        Me.btnTestLocal = New System.Windows.Forms.Button()
        Me.txt_SQLServer = New System.Windows.Forms.TextBox()
        Me.txt_UserName = New System.Windows.Forms.TextBox()
        Me.btn_Ok = New System.Windows.Forms.Button()
        Me.txt_SAPPass = New System.Windows.Forms.TextBox()
        Me.txt_Password = New System.Windows.Forms.TextBox()
        Me.txt_SAPUser = New System.Windows.Forms.TextBox()
        Me.cmb_Database = New System.Windows.Forms.ComboBox()
        Me.txt_LicenseServer = New System.Windows.Forms.TextBox()
        Me.MyLabel7 = New GSTExport.myLabel()
        Me.MyLabel4 = New GSTExport.myLabel()
        Me.MyLabel6 = New GSTExport.myLabel()
        Me.MyLabel5 = New GSTExport.myLabel()
        Me.MyLabel3 = New GSTExport.myLabel()
        Me.MyLabel2 = New GSTExport.myLabel()
        Me.MyLabel1 = New GSTExport.myLabel()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.txtServiceStatus = New System.Windows.Forms.TextBox()
        Me.MyLabel10 = New GSTExport.myLabel()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnunregister = New System.Windows.Forms.Button()
        Me.btnRegister = New System.Windows.Forms.Button()
        Me.txtServiceName = New System.Windows.Forms.TextBox()
        Me.MyLabel9 = New GSTExport.myLabel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(1, 31)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(380, 381)
        Me.TabControl1.TabIndex = 48
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(372, 352)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Connection"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.cbSQLType)
        Me.GroupBox1.Controls.Add(Me.MyLabel8)
        Me.GroupBox1.Controls.Add(Me.btnTestLocal)
        Me.GroupBox1.Controls.Add(Me.txt_SQLServer)
        Me.GroupBox1.Controls.Add(Me.txt_UserName)
        Me.GroupBox1.Controls.Add(Me.btn_Ok)
        Me.GroupBox1.Controls.Add(Me.txt_SAPPass)
        Me.GroupBox1.Controls.Add(Me.txt_Password)
        Me.GroupBox1.Controls.Add(Me.txt_SAPUser)
        Me.GroupBox1.Controls.Add(Me.cmb_Database)
        Me.GroupBox1.Controls.Add(Me.txt_LicenseServer)
        Me.GroupBox1.Controls.Add(Me.MyLabel7)
        Me.GroupBox1.Controls.Add(Me.MyLabel4)
        Me.GroupBox1.Controls.Add(Me.MyLabel6)
        Me.GroupBox1.Controls.Add(Me.MyLabel5)
        Me.GroupBox1.Controls.Add(Me.MyLabel3)
        Me.GroupBox1.Controls.Add(Me.MyLabel2)
        Me.GroupBox1.Controls.Add(Me.MyLabel1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(358, 346)
        Me.GroupBox1.TabIndex = 50
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Local Connection"
        '
        'cbSQLType
        '
        Me.cbSQLType.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSQLType.FormattingEnabled = True
        Me.cbSQLType.Items.AddRange(New Object() {"2005", "2008", "2012"})
        Me.cbSQLType.Location = New System.Drawing.Point(143, 22)
        Me.cbSQLType.Name = "cbSQLType"
        Me.cbSQLType.Size = New System.Drawing.Size(195, 23)
        Me.cbSQLType.TabIndex = 46
        Me.cbSQLType.Text = "2008"
        '
        'MyLabel8
        '
        Me.MyLabel8.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel8.Caption = "SQL Type"
        Me.MyLabel8.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(15, 28)
        Me.MyLabel8.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel8.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel8.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel8.TabIndex = 47
        Me.MyLabel8.TabStop = False
        '
        'btnTestLocal
        '
        Me.btnTestLocal.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btnTestLocal.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnTestLocal.Location = New System.Drawing.Point(227, 268)
        Me.btnTestLocal.Name = "btnTestLocal"
        Me.btnTestLocal.Size = New System.Drawing.Size(111, 29)
        Me.btnTestLocal.TabIndex = 45
        Me.btnTestLocal.Text = "Test Connection"
        Me.btnTestLocal.UseVisualStyleBackColor = False
        '
        'txt_SQLServer
        '
        Me.txt_SQLServer.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_SQLServer.Location = New System.Drawing.Point(142, 51)
        Me.txt_SQLServer.Name = "txt_SQLServer"
        Me.txt_SQLServer.Size = New System.Drawing.Size(196, 21)
        Me.txt_SQLServer.TabIndex = 43
        '
        'txt_UserName
        '
        Me.txt_UserName.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_UserName.Location = New System.Drawing.Point(143, 78)
        Me.txt_UserName.Name = "txt_UserName"
        Me.txt_UserName.Size = New System.Drawing.Size(195, 21)
        Me.txt_UserName.TabIndex = 18
        '
        'btn_Ok
        '
        Me.btn_Ok.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btn_Ok.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btn_Ok.Location = New System.Drawing.Point(142, 268)
        Me.btn_Ok.Name = "btn_Ok"
        Me.btn_Ok.Size = New System.Drawing.Size(79, 29)
        Me.btn_Ok.TabIndex = 36
        Me.btn_Ok.Text = "Save"
        Me.btn_Ok.UseVisualStyleBackColor = False
        '
        'txt_SAPPass
        '
        Me.txt_SAPPass.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_SAPPass.Location = New System.Drawing.Point(143, 234)
        Me.txt_SAPPass.Name = "txt_SAPPass"
        Me.txt_SAPPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(8226)
        Me.txt_SAPPass.Size = New System.Drawing.Size(195, 21)
        Me.txt_SAPPass.TabIndex = 28
        '
        'txt_Password
        '
        Me.txt_Password.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_Password.Location = New System.Drawing.Point(143, 106)
        Me.txt_Password.Name = "txt_Password"
        Me.txt_Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(8226)
        Me.txt_Password.Size = New System.Drawing.Size(195, 21)
        Me.txt_Password.TabIndex = 20
        '
        'txt_SAPUser
        '
        Me.txt_SAPUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_SAPUser.Location = New System.Drawing.Point(143, 206)
        Me.txt_SAPUser.Name = "txt_SAPUser"
        Me.txt_SAPUser.Size = New System.Drawing.Size(195, 21)
        Me.txt_SAPUser.TabIndex = 26
        '
        'cmb_Database
        '
        Me.cmb_Database.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmb_Database.FormattingEnabled = True
        Me.cmb_Database.Location = New System.Drawing.Point(143, 134)
        Me.cmb_Database.Name = "cmb_Database"
        Me.cmb_Database.Size = New System.Drawing.Size(195, 23)
        Me.cmb_Database.TabIndex = 22
        '
        'txt_LicenseServer
        '
        Me.txt_LicenseServer.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_LicenseServer.Location = New System.Drawing.Point(143, 178)
        Me.txt_LicenseServer.Name = "txt_LicenseServer"
        Me.txt_LicenseServer.Size = New System.Drawing.Size(195, 21)
        Me.txt_LicenseServer.TabIndex = 24
        '
        'MyLabel7
        '
        Me.MyLabel7.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel7.Caption = "SQL Server"
        Me.MyLabel7.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(14, 55)
        Me.MyLabel7.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel7.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel7.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel7.TabIndex = 44
        Me.MyLabel7.TabStop = False
        '
        'MyLabel4
        '
        Me.MyLabel4.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel4.Caption = "SAP License Server"
        Me.MyLabel4.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(12, 182)
        Me.MyLabel4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel4.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel4.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel4.TabIndex = 40
        Me.MyLabel4.TabStop = False
        '
        'MyLabel6
        '
        Me.MyLabel6.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel6.Caption = "SAP Password"
        Me.MyLabel6.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(13, 238)
        Me.MyLabel6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel6.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel6.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel6.TabIndex = 42
        Me.MyLabel6.TabStop = False
        '
        'MyLabel5
        '
        Me.MyLabel5.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel5.Caption = "SAP User"
        Me.MyLabel5.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(13, 210)
        Me.MyLabel5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel5.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel5.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel5.TabIndex = 41
        Me.MyLabel5.TabStop = False
        '
        'MyLabel3
        '
        Me.MyLabel3.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel3.Caption = "Database"
        Me.MyLabel3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel3.Location = New System.Drawing.Point(15, 140)
        Me.MyLabel3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel3.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel3.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel3.TabIndex = 39
        Me.MyLabel3.TabStop = False
        '
        'MyLabel2
        '
        Me.MyLabel2.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel2.Caption = "SQL Password"
        Me.MyLabel2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(15, 110)
        Me.MyLabel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel2.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel2.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel2.TabIndex = 38
        Me.MyLabel2.TabStop = False
        '
        'MyLabel1
        '
        Me.MyLabel1.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel1.Caption = "SQL User"
        Me.MyLabel1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(15, 82)
        Me.MyLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel1.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel1.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel1.TabIndex = 37
        Me.MyLabel1.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtServiceStatus)
        Me.TabPage2.Controls.Add(Me.MyLabel10)
        Me.TabPage2.Controls.Add(Me.btnStop)
        Me.TabPage2.Controls.Add(Me.btnStart)
        Me.TabPage2.Controls.Add(Me.btnunregister)
        Me.TabPage2.Controls.Add(Me.btnRegister)
        Me.TabPage2.Controls.Add(Me.txtServiceName)
        Me.TabPage2.Controls.Add(Me.MyLabel9)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(372, 352)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Alert Service"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'txtServiceStatus
        '
        Me.txtServiceStatus.Enabled = False
        Me.txtServiceStatus.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServiceStatus.Location = New System.Drawing.Point(113, 34)
        Me.txtServiceStatus.Name = "txtServiceStatus"
        Me.txtServiceStatus.Size = New System.Drawing.Size(196, 21)
        Me.txtServiceStatus.TabIndex = 51
        Me.txtServiceStatus.Text = "Running"
        '
        'MyLabel10
        '
        Me.MyLabel10.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel10.Caption = "Service Status"
        Me.MyLabel10.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(11, 38)
        Me.MyLabel10.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel10.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel10.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel10.TabIndex = 52
        Me.MyLabel10.TabStop = False
        '
        'btnStop
        '
        Me.btnStop.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btnStop.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnStop.Location = New System.Drawing.Point(217, 106)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(92, 29)
        Me.btnStop.TabIndex = 50
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = False
        '
        'btnStart
        '
        Me.btnStart.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btnStart.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnStart.Location = New System.Drawing.Point(113, 106)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(92, 29)
        Me.btnStart.TabIndex = 49
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = False
        '
        'btnunregister
        '
        Me.btnunregister.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btnunregister.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnunregister.Location = New System.Drawing.Point(217, 71)
        Me.btnunregister.Name = "btnunregister"
        Me.btnunregister.Size = New System.Drawing.Size(92, 29)
        Me.btnunregister.TabIndex = 48
        Me.btnunregister.Text = "Unregister"
        Me.btnunregister.UseVisualStyleBackColor = False
        '
        'btnRegister
        '
        Me.btnRegister.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btnRegister.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.btnRegister.Location = New System.Drawing.Point(113, 71)
        Me.btnRegister.Name = "btnRegister"
        Me.btnRegister.Size = New System.Drawing.Size(92, 29)
        Me.btnRegister.TabIndex = 47
        Me.btnRegister.Text = "Register"
        Me.btnRegister.UseVisualStyleBackColor = False
        '
        'txtServiceName
        '
        Me.txtServiceName.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServiceName.Location = New System.Drawing.Point(113, 6)
        Me.txtServiceName.Name = "txtServiceName"
        Me.txtServiceName.Size = New System.Drawing.Size(196, 21)
        Me.txtServiceName.TabIndex = 45
        Me.txtServiceName.Text = "SAPB1Addon_Alert"
        '
        'MyLabel9
        '
        Me.MyLabel9.BackColor = System.Drawing.Color.Transparent
        Me.MyLabel9.Caption = "Service Name"
        Me.MyLabel9.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(11, 10)
        Me.MyLabel9.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MyLabel9.MaximumSize = New System.Drawing.Size(500, 18)
        Me.MyLabel9.MinimumSize = New System.Drawing.Size(0, 17)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(135, 17)
        Me.MyLabel9.TabIndex = 46
        Me.MyLabel9.TabStop = False
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(159, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(7, 10)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(87, 29)
        Me.Button1.TabIndex = 37
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.Transparent
        Me.Panel5.Controls.Add(Me.Button1)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel5.Location = New System.Drawing.Point(1, 362)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(380, 50)
        Me.Panel5.TabIndex = 49
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(382, 413)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.TabControl1)
        Me.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.MaximizeBox = False
        Me.Name = "frmSettings"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Settings"
        Me.TopMost = True
        Me.Controls.SetChildIndex(Me.TabControl1, 0)
        Me.Controls.SetChildIndex(Me.Panel5, 0)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents cbSQLType As System.Windows.Forms.ComboBox
    Friend WithEvents MyLabel8 As GSTExport.myLabel
    Friend WithEvents btnTestLocal As System.Windows.Forms.Button
    Private WithEvents txt_SQLServer As System.Windows.Forms.TextBox
    Private WithEvents txt_UserName As System.Windows.Forms.TextBox
    Friend WithEvents btn_Ok As System.Windows.Forms.Button
    Private WithEvents txt_SAPPass As System.Windows.Forms.TextBox
    Private WithEvents txt_Password As System.Windows.Forms.TextBox
    Private WithEvents txt_SAPUser As System.Windows.Forms.TextBox
    Private WithEvents cmb_Database As System.Windows.Forms.ComboBox
    Private WithEvents txt_LicenseServer As System.Windows.Forms.TextBox
    Friend WithEvents MyLabel7 As GSTExport.myLabel
    Friend WithEvents MyLabel4 As GSTExport.myLabel
    Friend WithEvents MyLabel6 As GSTExport.myLabel
    Friend WithEvents MyLabel5 As GSTExport.myLabel
    Friend WithEvents MyLabel3 As GSTExport.myLabel
    Friend WithEvents MyLabel2 As GSTExport.myLabel
    Friend WithEvents MyLabel1 As GSTExport.myLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Panel5 As System.Windows.Forms.Panel


    Friend WithEvents btnRegister As System.Windows.Forms.Button
    Private WithEvents txtServiceName As System.Windows.Forms.TextBox
    Friend WithEvents MyLabel9 As GSTExport.myLabel
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnunregister As System.Windows.Forms.Button
    Private WithEvents txtServiceStatus As System.Windows.Forms.TextBox
    Friend WithEvents MyLabel10 As GSTExport.myLabel
End Class
