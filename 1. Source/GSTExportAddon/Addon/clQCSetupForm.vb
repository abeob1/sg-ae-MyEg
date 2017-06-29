Imports SAPbouiCOM
Imports GSTExport.clDrawItem

Module clQCSetupForm
    Public Sub DrawQCSetupForm()
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim fcp As SAPbouiCOM.FormCreationParams = DirectCast(oApp.CreateObject(BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
        fcp.FormType = "QCSETUP"
        fcp.BorderStyle = BoFormBorderStyle.fbs_Fixed
        fcp.UniqueID = String.Format("QCSETUP{0}", (New Random().[Next](1000).ToString()))

        oForm = oApp.Forms.AddEx(fcp)
        AddDataSourceToFormMaster(oForm)
        oForm.Mode = BoFormMode.fm_ADD_MODE
        oForm.Width = 500
        oForm.Height = 500
        oForm.ClientHeight = 500
        oForm.ClientWidth = 550

        oForm.EnableMenu(1290, True) ' First Data Record
        oForm.EnableMenu(1288, True) ' Next Record
        oForm.EnableMenu(1289, True) ' Previous Record
        oForm.EnableMenu(1292, True) ' Last Data Record

        'Code static text
        Dim layout As New ItemLayout(10, 10, 20, 60)
        Dim oItem As SAPbouiCOM.Item = FormItemCreator.CreateItem(oForm, "lblCode", BoFormItemTypes.it_STATIC, layout)
        Dim staticText As SAPbouiCOM.StaticText = DirectCast(oItem.Specific, SAPbouiCOM.StaticText)
        staticText.Caption = "Code"

        'Code edit text
        layout = New ItemLayout(70, 10, 20, 100)
        oItem = FormItemCreator.CreateItem(oForm, "txtCode", BoFormItemTypes.it_EDIT, layout)
        oEditText = oItem.Specific
        Try
            oEditText.DataBind.SetBound(True, "@QCSETUP", "Code")
        Catch ex As Exception
            MsgBoxWrapper(ex.Message + " DataBind.SetBound(True, @QCSETUP, Code) ", "", True)
        End Try

        ''Code maxtrix
        'layout = New ItemLayout(70, 10, 20, 100)
        'oItem = FormItemCreator.CreateItem(oForm, "mtSetup", BoFormItemTypes.it_MATRIX, layout)
        'oMatrix = oItem.Specific
        'Try
        '    Dim oColumn As SAPbouiCOM.Column
        '    oColumn = oMatrix.Columns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        '    oColumn.Editable = False
        '    oColumn.Width = 20
        '    oColumn.TitleObject.Caption = "#"

        '    oColumn = oMatrix.Columns.Add("clCode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        '    oColumn.Width = 40
        '    oColumn.TitleObject.Caption = "Code"
        '    oColumn.DataBind.Bind("@QCSETUP", "Code")

        '    oColumn = oMatrix.Columns.Add("clName", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        '    oColumn.Width = oForm.Width - 100
        '    oColumn.TitleObject.Caption = "Name"
        '    oColumn.Editable = False
        '    oColumn.DataBind.Bind("@QCSETUP", "Name")

        '    oMatrix.LoadFromDataSource()

        'Catch ex As Exception
        '    MsgBoxWrapper(ex.Message + " DataBind.SetBound(True, @QCSETUP, Code) ", "", True)
        'End Try

        Dim oButton As SAPbouiCOM.Button
        layout = New ItemLayout(10, oForm.Height - 60, 20, 40)
        oItem = FormItemCreator.CreateItem(oForm, "1", BoFormItemTypes.it_BUTTON, layout)
        oButton = DirectCast(oItem.Specific, SAPbouiCOM.Button)

        layout = New ItemLayout(60, oForm.Height - 60, 20, 40)
        oItem = FormItemCreator.CreateItem(oForm, "2", BoFormItemTypes.it_BUTTON, layout)
        oButton = DirectCast(oItem.Specific, SAPbouiCOM.Button)

        oForm.Title = "QC SETUP"


        oForm.Visible = True
    End Sub
    Public Sub AddDataSourceToFormMaster(ByVal ooForm As SAPbouiCOM.Form)
        Try
            oDBDataSource_MH = ooForm.DataSources.DBDataSources.Add("@QCSETUP")
        Catch ex As Exception
            MsgBoxWrapper(ex.Message + " AddDataSourceToFormMaster : 294", "", True)
        End Try

    End Sub
End Module
