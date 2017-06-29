Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Reflection

Public Class frmSettings

    Private Sub frmSettings_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Application.DoEvents()
        Me.BringToFront()
        Me.TopMost = True
    End Sub
    Private Sub Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'IF not EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID('_OCNT')) CREATE TABLE _OCNT(CnnStr [nvarchar](max) NOT NULL)
        If PublicVariable.oCompany.Server <> System.Net.Dns.GetHostName() Then
            TabControl1.TabPages.RemoveByKey("TabPage2")
        End If
        RefreshStatus()

        Dim dt As DataTable = Functions.DoQueryReturnDT(" Select * from _OCNT")
        If IsNothing(dt) Then
            cmb_Database.Text = PublicVariable.oCompany.CompanyDB
            txt_SAPUser.Text = PublicVariable.oCompany.UserName
            txt_SAPPass.Text = ""
            txt_SQLServer.Text = PublicVariable.oCompany.Server
            txt_UserName.Text = "sa"
            txt_Password.Text = ""
            txt_LicenseServer.Text = PublicVariable.oCompany.LicenseServer
            cbSQLType.Text = "2008"
        Else
            If dt.Rows.Count > 0 Then
                'SAP_YGN;manager;4321;PC\WIN764;sa;win764;PC;2008
                Dim MyArr As Array = oMD5.DecryptPassword(dt.Rows(0).Item("CnnStr").ToString).Split(";")
                cmb_Database.Text = MyArr(0).ToString()
                txt_SAPUser.Text = MyArr(1).ToString()
                txt_SAPPass.Text = MyArr(2).ToString()
                txt_SQLServer.Text = MyArr(3).ToString()
                txt_UserName.Text = MyArr(4).ToString()
                txt_Password.Text = MyArr(5).ToString()
                txt_LicenseServer.Text = MyArr(6).ToString()
                cbSQLType.Text = MyArr(7).ToString()
            Else
                cmb_Database.Text = PublicVariable.oCompany.CompanyDB
                txt_SAPUser.Text = PublicVariable.oCompany.UserName
                txt_SAPPass.Text = ""
                txt_SQLServer.Text = PublicVariable.oCompany.Server
                txt_UserName.Text = "sa"
                txt_Password.Text = ""
                txt_LicenseServer.Text = PublicVariable.oCompany.LicenseServer
                cbSQLType.Text = "2008"
            End If
        End If
    End Sub
    Private Sub btn_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Ok.Click
        Try
            Dim str As String = ""
            str = cmb_Database.Text + ";" + txt_SAPUser.Text + ";" + txt_SAPPass.Text + ";" + txt_SQLServer.Text + ";" + txt_UserName.Text + ";" + txt_Password.Text + ";" + txt_LicenseServer.Text + ";" + cbSQLType.Text
            str = oMD5.EncryptPassword(str)
            str = str.Replace("'", "''")
            Dim ret As String = ""
            Functions.DoQueryReturnDT("if (select COUNT(*) from _OCNT)>0 update _OCNT set CnnStr='" + str + "' else insert into dbo._OCNT values ('" + str + "')")
            Functions.SaveTextToFile(str, Application.StartupPath + "\Connect.txt")
            MessageBox.Show("Operation Complete!")
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Private Sub cmb_Database_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb_Database.DropDown
        Dim table As DataTable = New DataTable("DBName")
        table.Columns.Add("Name")
        Dim str As String = String.Format("Data Source={0};User ID={1};Password={2}", txt_SQLServer.Text, txt_UserName.Text, txt_Password.Text)
        Dim oldValue As String = cmb_Database.Text
        Dim sqlConx As SqlConnection = New SqlConnection(str)
        Try

            sqlConx.Open()
            Dim tblDatabases As DataTable = sqlConx.GetSchema("Databases")
            sqlConx.Close()

            For Each row As DataRow In tblDatabases.Rows
                table.Rows.Add(row("database_name"))
            Next
            cmb_Database.DataSource = table
            cmb_Database.DisplayMember = "Name"
            cmb_Database.ValueMember = "Name"
            cmb_Database.Text = ""
            cmb_Database.SelectedText = oldValue
        Catch
            MessageBox.Show("Can not get database list", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try

    End Sub
    Private Sub btnTestLocal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestLocal.Click
        Me.Cursor = Cursors.WaitCursor
        Dim sStr As String = ""
        Dim lErrCode As Integer
        Dim sErrMsg As String
        Dim TestCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company
        TestCompany.CompanyDB = cmb_Database.Text
        TestCompany.UserName = txt_SAPUser.Text
        TestCompany.Password = txt_SAPPass.Text
        TestCompany.Server = txt_SQLServer.Text
        TestCompany.DbUserName = txt_UserName.Text
        TestCompany.DbPassword = txt_Password.Text
        TestCompany.LicenseServer = txt_LicenseServer.Text
        Select Case cbSQLType.Text
            Case "2008"
                TestCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            Case "2005"
                TestCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Case "2012"
                TestCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
        End Select

        lErrCode = TestCompany.Connect
        If lErrCode <> 0 Then
            TestCompany.GetLastError(lErrCode, sErrMsg)
            MessageBox.Show(sErrMsg)
        Else
            TestCompany.Disconnect()
            MessageBox.Show("Connect to SAP successful!")
        End If
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim a As New ServiceController(txtServiceName.Text)
        Dim str As String
        str = a.Start()
        RefreshStatus()
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        Dim a As New ServiceController(txtServiceName.Text)
        Dim str As String = a.Stop()
        RefreshStatus()
    End Sub

    Private Sub btnRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegister.Click
        Dim a As New ServiceController(txtServiceName.Text)
        a.Description = txtServiceName.Text
        a.DisplayName = txtServiceName.Text
        a.ServiceName = txtServiceName.Text
        a.StartupType = ServiceController.ServiceStartupType.Automatic



        Dim location = Assembly.GetExecutingAssembly().Location
        Dim appPath = Path.GetDirectoryName(location)       ' C:\Some\Directory
        Dim appName = Path.GetFileName(location)


        'copy file
        CopyDirectory(appPath, appPath + "\" + txtServiceName.Text)


        Dim sReturn As String
        sReturn = a.Register(appPath + "\" + txtServiceName.Text + "\" + appName + " -service")
        If sReturn = "" Then
            MessageBox.Show("Register Sucessfull!")
            'Application.Exit()

        Else
            MessageBox.Show("Error: " + sReturn)
        End If
        RefreshStatus()
    End Sub

    Private Sub btnunregister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnunregister.Click
        Dim a As New ServiceController(txtServiceName.Text)
        Dim sReturn As String
        sReturn = a.Unregister()
        If sReturn = "" Then
            Dim location = Assembly.GetExecutingAssembly().Location
            Dim appPath = Path.GetDirectoryName(location)       ' C:\Some\Directory
            Directory.Delete(appPath + "\" + txtServiceName.Text, True)
            MessageBox.Show("UnRegister Sucessfull!")
        Else
            MessageBox.Show("Error: " + sReturn)
        End If
        RefreshStatus()
    End Sub

    Private Sub RefreshStatus()
        Dim a As New ServiceController(txtServiceName.Text)
        If a.Status = "" Then
            btnRegister.Enabled = True
            btnunregister.Enabled = False
            btnStart.Enabled = False
            btnStop.Enabled = False
        ElseIf a.Status = "Stopped" Then
            btnStart.Enabled = True
            btnStop.Enabled = False
            btnRegister.Enabled = False
            btnunregister.Enabled = True
        ElseIf a.Status = "Running" Then
            btnStart.Enabled = False
            btnStop.Enabled = True
            btnRegister.Enabled = False
            btnunregister.Enabled = True
        End If
    End Sub

    Private Sub btnInit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim sqlstring As String

        'lblStatus.Text = lblStatus.Text + "Create sp QC notification" + vbCrLf
        'sqlstring = GetFileContents(Application.StartupPath + "\[sp_AI_TransactionNotification_QC].sql")
        'Functions.SAP_Local_RunQuery(sqlstring)

        'lblStatus.Text = lblStatus.Text + "Alter SAP sp notification " + vbCrLf
        'sqlstring = GetFileContents(Application.StartupPath + "\[SBO_SP_TransactionNotification].sql")
        'Functions.SAP_Local_RunQuery(sqlstring)

        'lblStatus.Text = lblStatus.Text + "Create QC Tables" + vbCrLf
        'sqlstring = GetFileContents(Application.StartupPath + "\[QC Tables].sql")
        'Functions.SAP_Local_RunQuery(sqlstring)

        'lblStatus.Text = lblStatus.Text + "Create sp BOM Structure" + vbCrLf
        'sqlstring = GetFileContents(Application.StartupPath + "\[sp_QC_BOMStructure].sql")
        'Functions.SAP_Local_RunQuery(sqlstring)

    End Sub
    Public Function GetFileContents(ByVal FullPath As String, _
       Optional ByRef ErrInfo As String = "") As String

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

    Private Sub CopyDirectory(ByVal sourcePath As String, ByVal destPath As String)
        If Not Directory.Exists(destPath) Then
            Directory.CreateDirectory(destPath)
        End If

        For Each file__1 As String In Directory.GetFiles(sourcePath)
            Dim dest As String = Path.Combine(destPath, Path.GetFileName(file__1))
            File.Copy(file__1, dest)
        Next


    End Sub
End Class