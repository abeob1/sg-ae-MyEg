Public Class fBadDeptSetup
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Dim f As SAPbouiCOM.Form
    Private oDBDataSource As SAPbouiCOM.DBDataSource
    Private oUserDataSource As SAPbouiCOM.UserDataSource

    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        DrawForm()

    End Sub
    Private Sub DrawForm()
        Try

            Dim oItem As SAPbouiCOM.Item
            Dim oLabel As SAPbouiCOM.StaticText
            Dim cp As SAPbouiCOM.FormCreationParams
            Dim oEdit As SAPbouiCOM.EditText
            Dim oFld As SAPbouiCOM.Folder

            cp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            cp.FormType = "fBadDeptSetup"
            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 590
            f.ClientHeight = 350

            ' set the form title
            f.Title = "GST Setup (" + PublicVariable.Version + ")"

            f.DataSources.UserDataSources.Add("act1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act4", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act5", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act6", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act7", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act8", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act9", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act10", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("act12", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            f.DataSources.UserDataSources.Add("arintax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("arinname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("apintax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("apinname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("arouttax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("aroutname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("apouttax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("apoutname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            f.DataSources.UserDataSources.Add("sqlpass", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            f.DataSources.UserDataSources.Add("ck", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("GST03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("GST03Det", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("ContraAct", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("ckReverse", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("conactname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("revintax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("revouttax", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("revintaxn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("revouttaxn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("ckContraJV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("ckBadDJV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            f.DataSources.UserDataSources.Add("DOAct", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            f.DataSources.UserDataSources.Add("DOName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            AddChooseFromList(f)

            ' ------------------------------Adding Folder AR Bad Debt---------------------------------
            oItem = f.Items.Add("f1", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Left = 5
            oItem.Top = 10
            oItem.Width = 220
            oItem.Height = 300
            oFld = oItem.Specific
            oFld.Caption = "AR Bad Debt"

            AddingItemInARBadDebt(f)


            ' ------------------------------Adding Folder AP Bad Debt---------------------------------
            oItem = f.Items.Add("f2", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Left = 5
            oItem.Top = 10
            oItem.Width = 220
            oItem.Height = 300
            oFld = oItem.Specific
            oFld.Caption = "AP Bad Debt"
            oFld.GroupWith("f1")

            AddingItemInAPBadDebt(f)

            ' ------------------------------Adding Folder General---------------------------------
            oItem = f.Items.Add("f3", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Left = 100
            oItem.Top = 10
            oItem.Width = 220
            oItem.Height = 300
            oFld = oItem.Specific
            oFld.Caption = "General"
            oFld.GroupWith("f2")

            AddingItemInGeneral(f)

            f.PaneLevel = 1

            '---------------------------adding folder Reverse Mechanism--------------------
            oItem = f.Items.Add("f4", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Left = 200
            oItem.Top = 10
            oItem.Width = 220
            oItem.Height = 300
            oFld = oItem.Specific
            oFld.Caption = "Reverse Mechanism"
            oFld.GroupWith("f3")

            AddingItemInReverseMechanism(f)

            ' -------------------Add the OK button----------------------
            oItem = f.Items.Add("btnU", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 320
            Dim obtn As SAPbouiCOM.Button
            obtn = oItem.Specific
            obtn.Caption = "Update"
            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 320


            f.PaneLevel = 1
            f.Visible = True

            LoadSetup(f)

           
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DrawForm Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub AddChooseFromList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact3"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact4"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact5"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact6"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact7"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for account
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 1
            oCFLCreationParams.UniqueID = "clact8"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            'CFL for tax 1
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "I"
            oCFL.SetConditions(oCons)

            'CFL for tax 2
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "I"
            oCFL.SetConditions(oCons)

            'CFL for tax 3
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax3"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)

            'CFL for tax 4
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax4"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)

            'CFL for tax 5
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax5"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)

            'CFL for tax 6
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.ObjectType = 5
            oCFLCreationParams.UniqueID = "cltax6"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Category"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "I"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("AddChooseFromList Event: " + ex.ToString, , True)
        Finally
            System.GC.Collect() 'Release the handle to the table
        End Try
    End Sub
    Private Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction = False Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                End If
                Select Case pVal.FormTypeEx
                    Case "fBadDeptSetup"
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "f1"
                                        oForm.PaneLevel = 1
                                    Case "f2"
                                        oForm.PaneLevel = 2
                                    Case "f3"
                                        oForm.PaneLevel = 3
                                    Case "f4"
                                        oForm.PaneLevel = 4
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                                SBO_Application = Nothing
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                If pVal.ItemUID = "btnU" Then
                                    Dim str As String = SaveData(oForm)
                                    If str = "" Then
                                        SBO_Application.SetStatusBarMessage("Operation Successful!", , False)
                                    Else
                                        SBO_Application.SetStatusBarMessage(str, , True)
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvento = pVal
                                Dim sCFL_ID As String
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                Dim AcctCode, AcctName As String
                                Try
                                    AcctCode = oDataTable.GetValue(0, 0)
                                    AcctName = oDataTable.GetValue(1, 0)
                                Catch ex As Exception

                                End Try
                                Select Case pVal.ItemUID
                                    Case "txt1"
                                        oForm.DataSources.UserDataSources.Item("act1").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act2").ValueEx = AcctName
                                    Case "txt3"
                                        oForm.DataSources.UserDataSources.Item("act3").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act4").ValueEx = AcctName
                                    Case "txt5"
                                        oForm.DataSources.UserDataSources.Item("act5").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act6").ValueEx = AcctName
                                    Case "txt7"
                                        oForm.DataSources.UserDataSources.Item("act7").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act8").ValueEx = AcctName
                                    Case "txt9"
                                        oForm.DataSources.UserDataSources.Item("act9").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act10").ValueEx = AcctName
                                    Case "txt11"
                                        oForm.DataSources.UserDataSources.Item("act11").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("act12").ValueEx = AcctName
                                    Case "txt13"
                                        oForm.DataSources.UserDataSources.Item("arintax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("arinname").ValueEx = AcctName
                                    Case "txt15"
                                        oForm.DataSources.UserDataSources.Item("apintax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("apinname").ValueEx = AcctName
                                    Case "txt17"
                                        oForm.DataSources.UserDataSources.Item("arouttax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("aroutname").ValueEx = AcctName
                                    Case "txt19"
                                        oForm.DataSources.UserDataSources.Item("apouttax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("apoutname").ValueEx = AcctName
                                    Case "txt25"
                                        oForm.DataSources.UserDataSources.Item("ContraAct").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("conactname").ValueEx = AcctName

                                    Case "txt27"
                                        oForm.DataSources.UserDataSources.Item("revouttax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("revouttaxn").ValueEx = AcctName

                                    Case "txt29"
                                        oForm.DataSources.UserDataSources.Item("revintax").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("revintaxn").ValueEx = AcctName

                                    Case "txt31"
                                        oForm.DataSources.UserDataSources.Item("DOAct").ValueEx = AcctCode
                                        oForm.DataSources.UserDataSources.Item("DOName").ValueEx = AcctName
                                End Select
                        End Select
                End Select
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("ItemEvent Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub LoadSetup(ByVal oForm As SAPbouiCOM.Form)
        oForm.Freeze(True)
        Dim oedit As SAPbouiCOM.EditText
        Dim ock As SAPbouiCOM.CheckBox
        Dim oUserTable As SAPbobsCOM.UserTable
        oUserTable = PublicVariable.oCompany.UserTables.Item("GSTSETUP")
        If oUserTable.GetByKey("ACT1") Then
            oedit = oForm.Items.Item("txt1").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("ACT2") Then
            oedit = oForm.Items.Item("txt3").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If
                      
        If oUserTable.GetByKey("ACT3") Then
            oedit = oForm.Items.Item("txt5").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("ACT4") Then
            oedit = oForm.Items.Item("txt7").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("ACT5") Then
            oedit = oForm.Items.Item("txt9").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("ACT6") Then
            oedit = oForm.Items.Item("txt11").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("ARInTaxCode") Then
            oedit = oForm.Items.Item("txt13").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("APInTaxCode") Then
            oedit = oForm.Items.Item("txt15").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("AROutTaxCode") Then
            oedit = oForm.Items.Item("txt17").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("APOutTaxCode") Then
            oedit = oForm.Items.Item("txt19").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        If oUserTable.GetByKey("SQLPass") Then
            oedit = oForm.Items.Item("txt21").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If

        

        
       
        'If oUserTable.GetByKey("GSTRpt") Then
        '    oedit = oForm.Items.Item("txt22").Specific
        '    oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString()
        'End If

        'If oUserTable.GetByKey("GSTRptDet") Then
        '    oedit = oForm.Items.Item("txt23").Specific
        '    oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString()
        'End If

        If oUserTable.GetByKey("ContraAct") Then
            oedit = oForm.Items.Item("txt25").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString()
        End If

        If oUserTable.GetByKey("revOutTaxCode") Then
            oedit = oForm.Items.Item("txt27").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString()
        End If

        If oUserTable.GetByKey("revInTaxCode") Then
            oedit = oForm.Items.Item("txt29").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString()
        End If

        If oUserTable.GetByKey("DOAct") Then
            oedit = oForm.Items.Item("txt31").Specific
            oedit.Value = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
        End If
        

        If oUserTable.GetByKey("ReverseMechanism") Then
            ock = oForm.Items.Item("ck2").Specific
            If oUserTable.UserFields.Fields.Item("U_Value").Value.ToString = "Y" Then
                ock.Checked = True
            Else
                ock.Checked = False
            End If
        End If

        If oUserTable.GetByKey("OneJE") Then
            ock = oForm.Items.Item("ck1").Specific
            If oUserTable.UserFields.Fields.Item("U_Value").Value.ToString = "Y" Then
                ock.Checked = True
            Else
                ock.Checked = False
            End If
        End If

        If oUserTable.GetByKey("BadDebtJV") Then
            ock = oForm.Items.Item("ck3").Specific
            If oUserTable.UserFields.Fields.Item("U_Value").Value.ToString = "Y" Then
                ock.Checked = True
            Else
                ock.Checked = False
            End If
        End If

        oForm.Freeze(False)

    End Sub

    Private Function SaveData(oForm As SAPbouiCOM.Form) As String
        Try

            Dim act1, act2, act3, act4, act5, act6, ARINTaxCode, APINTaxCode, AROUTTaxCode, APOUTTaxCode, OneJE, SQLPass, ContraAct, ReverseMechanism, RevInputTax, RevOutputTax, DOAct, BadDebtJV As String

            Dim oedit As SAPbouiCOM.EditText
            oedit = oForm.Items.Item("txt1").Specific
            act1 = oedit.Value
            oedit = oForm.Items.Item("txt3").Specific
            act2 = oedit.Value
            oedit = oForm.Items.Item("txt5").Specific
            act3 = oedit.Value
            oedit = oForm.Items.Item("txt7").Specific
            act4 = oedit.Value
            oedit = oForm.Items.Item("txt9").Specific
            act5 = oedit.Value
            oedit = oForm.Items.Item("txt11").Specific
            act6 = oedit.Value
            oedit = oForm.Items.Item("txt13").Specific
            ARINTaxCode = oedit.Value
            oedit = oForm.Items.Item("txt15").Specific
            APINTaxCode = oedit.Value
            oedit = oForm.Items.Item("txt17").Specific
            AROUTTaxCode = oedit.Value
            oedit = oForm.Items.Item("txt19").Specific
            APOUTTaxCode = oedit.Value
            oedit = oForm.Items.Item("txt19").Specific
            APOUTTaxCode = oedit.Value
            oedit = oForm.Items.Item("txt21").Specific
            SQLPass = oedit.Value

            oedit = oForm.Items.Item("txt25").Specific
            ContraAct = oedit.Value

            oedit = oForm.Items.Item("txt27").Specific
            RevOutputTax = oedit.Value
            oedit = oForm.Items.Item("txt29").Specific
            RevInputTax = oedit.Value

            oedit = oForm.Items.Item("txt31").Specific
            DOAct = oedit.Value

            Dim ck As SAPbouiCOM.CheckBox
            ck = oForm.Items.Item("ck1").Specific
            If ck.Checked Then
                OneJE = "Y"
            Else
                OneJE = "N"
            End If

            ck = oForm.Items.Item("ck2").Specific
            If ck.Checked Then
                ReverseMechanism = "Y"
            Else
                ReverseMechanism = "N"
            End If

            ck = oForm.Items.Item("ck3").Specific
            If ck.Checked Then
                BadDebtJV = "Y"
            Else
                BadDebtJV = "N"
            End If

            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = PublicVariable.oCompany.UserTables.Item("GSTSETUP")
            Dim rest As Integer = 0

            '----------------------------------------ContraAct-------------------------------------
            If oUserTable.GetByKey("ContraAct") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = ContraAct
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ContraAct"
                oUserTable.Name = "Payment Contra Account"
                oUserTable.UserFields.Fields.Item("U_Value").Value = ContraAct
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------ReverseMechanism------------------------------------
            If oUserTable.GetByKey("ReverseMechanism") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = ReverseMechanism
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ReverseMechanism"
                oUserTable.Name = "Reverse Mechanism"
                oUserTable.UserFields.Fields.Item("U_Value").Value = ReverseMechanism
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If
            '-----------------------------------------ReverseMechanism Input Tax------------------------------------
            If oUserTable.GetByKey("revInTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = RevInputTax
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "revInTaxCode"
                oUserTable.Name = "Reverse Mechanism Input Tax"
                oUserTable.UserFields.Fields.Item("U_Value").Value = RevInputTax
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If
            '-----------------------------------------ReverseMechanism Output Tax------------------------------------
            If oUserTable.GetByKey("revOutTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = RevOutputTax
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "revOutTaxCode"
                oUserTable.Name = "Reverse Mechanism Output Tax"
                oUserTable.UserFields.Fields.Item("U_Value").Value = RevOutputTax
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AR Bad Debt Account------------------------------------
            If oUserTable.GetByKey("ACT1") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act1
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT1"
                oUserTable.Name = "AR Bad Debt Account"
                oUserTable.UserFields.Fields.Item("U_Value").Value = act1
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AR Bad Debt Relief Account------------------------------------
            If oUserTable.GetByKey("ACT2") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act2
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT2"
                oUserTable.Name = "AR Bad Debt Relief Account"
                oUserTable.UserFields.Fields.Item("U_Value").Value = act2
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AR Bad Debt Relief Re. Act.------------------------------------
            If oUserTable.GetByKey("ACT3") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act3
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT3"
                oUserTable.Name = "AR Bad Debt Relief Re. Act."
                oUserTable.UserFields.Fields.Item("U_Value").Value = act3
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AP Bad Debt Account------------------------------------
            If oUserTable.GetByKey("ACT4") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act4
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT4"
                oUserTable.Name = "AP Bad Debt Account"
                oUserTable.UserFields.Fields.Item("U_Value").Value = act4
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AP Bad Debt Relief Act------------------------------------
            If oUserTable.GetByKey("ACT5") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act5
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT5"
                oUserTable.Name = "AP Bad Debt Relief Act"
                oUserTable.UserFields.Fields.Item("U_Value").Value = act5
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AP Bad Debt Relief Re. Act.------------------------------------
            If oUserTable.GetByKey("ACT6") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = act6
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ACT6"
                oUserTable.Name = "AP Bad Debt Relief Re. Act."
                oUserTable.UserFields.Fields.Item("U_Value").Value = act6
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AR Input Tax Code------------------------------------
            If oUserTable.GetByKey("ARInTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = ARINTaxCode
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "ARInTaxCode"
                oUserTable.Name = "AR Input Tax Code"
                oUserTable.UserFields.Fields.Item("U_Value").Value = ARINTaxCode
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AP Input Tax Code------------------------------------
            If oUserTable.GetByKey("APInTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = APINTaxCode
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "APInTaxCode"
                oUserTable.Name = "AP Input Tax Code"
                oUserTable.UserFields.Fields.Item("U_Value").Value = APINTaxCode
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AR Output Tax Code------------------------------------
            If oUserTable.GetByKey("AROutTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = AROUTTaxCode
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "AROutTaxCode"
                oUserTable.Name = "AR Output Tax Code"
                oUserTable.UserFields.Fields.Item("U_Value").Value = AROUTTaxCode
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------AP Output Tax Codee------------------------------------
            If oUserTable.GetByKey("APOutTaxCode") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = APOUTTaxCode
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "APOutTaxCode"
                oUserTable.Name = "AP Output Tax Code"
                oUserTable.UserFields.Fields.Item("U_Value").Value = APOUTTaxCode
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------Consolidate JE------------------------------------
            If oUserTable.GetByKey("OneJE") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = OneJE
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "OneJE"
                oUserTable.Name = "Consolidate JE"
                oUserTable.UserFields.Fields.Item("U_Value").Value = OneJE
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If
            '-----------------------------------------Bad Debt JV------------------------------------
            If oUserTable.GetByKey("BadDebtJV") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = BadDebtJV
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "BadDebtJV"
                oUserTable.Name = "Bad Debt JV"
                oUserTable.UserFields.Fields.Item("U_Value").Value = BadDebtJV
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------SQLPass------------------------------------
            If oUserTable.GetByKey("SQLPass") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = SQLPass
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "SQLPass"
                oUserTable.Name = "SQLPass"
                oUserTable.UserFields.Fields.Item("U_Value").Value = SQLPass
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            '-----------------------------------------DOAct------------------------------------
            If oUserTable.GetByKey("DOAct") Then
                oUserTable.UserFields.Fields.Item("U_Value").Value = DOAct
                rest = oUserTable.Update()
            Else
                oUserTable.Code = "DOAct"
                oUserTable.Name = "DO Contra Act"
                oUserTable.UserFields.Fields.Item("U_Value").Value = DOAct
                oUserTable.Add()
            End If
            If rest <> 0 Then
                SBO_Application.SetStatusBarMessage(PublicVariable.oCompany.GetLastErrorDescription, , True)
            End If

            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function
#Region "Adding Items to Tab"
    Private Sub AddingItemInGeneral(oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim cp As SAPbouiCOM.FormCreationParams
        Dim oEdit As SAPbouiCOM.EditText
        Dim oFld As SAPbouiCOM.Folder

        oItem = f.Items.Add("ck1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Left = 10
        oItem.Top = 40
        oItem.DisplayDesc = True
        oItem.Width = 300
        oItem.FromPane = 3
        oItem.ToPane = 3
        Dim ock As SAPbouiCOM.CheckBox
        ock = oItem.Specific
        ock.DataBind.SetBound(True, "", "ck")
        ock.Caption = "Consolidate Bad Debt Relief"

        ' ------------------------------SQL Pass---------------------------------
        oItem = f.Items.Add("lbl21", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 60
        oItem.Width = 210
        oItem.FromPane = 3
        oItem.ToPane = 3
        oLabel = oItem.Specific
        oLabel.Caption = "SQL Password"
        oItem = f.Items.Add("txt21", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 60
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.FromPane = 3
        oItem.ToPane = 3
        oEdit = oItem.Specific
        oEdit.IsPassword = True
        oEdit.DataBind.SetBound(True, "", "sqlpass")

        '' ------------------------------GST 03 Report UID---------------------------------
        'oItem = f.Items.Add("lbl22", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = 10
        'oItem.Top = 480
        'oItem.Width = 210
        'oItem.FromPane = 3
        'oItem.ToPane = 3
        'oLabel = oItem.Specific
        'oLabel.Caption = "GST-03 Report MenuID"
        'oItem = f.Items.Add("txt22", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        'oItem.Left = 215
        'oItem.Top = 480
        'oItem.Width = 300
        'oItem.DisplayDesc = True
        'oItem.FromPane = 3
        'oItem.ToPane = 3
        'oEdit = oItem.Specific
        'oEdit.DataBind.SetBound(True, "", "GST03")
        '' ------------------------------GST 03 detail Report UID---------------------------------
        'oItem = f.Items.Add("lbl23", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        'oItem.Left = 10
        'oItem.Top = 500
        'oItem.Width = 210
        'oItem.FromPane = 3
        'oItem.ToPane = 3
        'oLabel = oItem.Specific
        'oLabel.Caption = "GST-03 Detail Report MenuID"
        'oItem = f.Items.Add("txt23", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        'oItem.Left = 215
        'oItem.Top = 500
        'oItem.Width = 300
        'oItem.FromPane = 3
        'oItem.ToPane = 3
        'oItem.DisplayDesc = True
        'oEdit = oItem.Specific
        'oEdit.DataBind.SetBound(True, "", "GST03Det")

        ' ------------------------------Payment Contra Account Code---------------------------------

        oItem = f.Items.Add("lbl25", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 80
        oItem.Width = 210
        oItem.FromPane = 3
        oItem.ToPane = 3
        oLabel = oItem.Specific
        oLabel.Caption = "Payment Contra Account"
        oItem = f.Items.Add("txt25", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 80
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 3
        oItem.ToPane = 3
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "ContraAct")
        oEdit.ChooseFromListUID = "clact7"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------Payment Contra Account name---------------------------------
        oItem = f.Items.Add("lbl26", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 100
        oItem.Width = 210
        oItem.FromPane = 3
        oItem.ToPane = 3
        oLabel = oItem.Specific
        oLabel.Caption = "Contra Account Name"
        oItem = f.Items.Add("txt26", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 100
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 3
        oItem.ToPane = 3
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "conactname")
        ' ------------------------------Bad Debt JV---------------------------------
        oItem = f.Items.Add("ck3", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Left = 10
        oItem.Top = 120
        oItem.DisplayDesc = True
        oItem.Width = 300
        oItem.FromPane = 3
        oItem.ToPane = 3
        ock = oItem.Specific
        ock.DataBind.SetBound(True, "", "ckBadDJV")
        ock.Caption = "Bad Debt Posting JV"

        '' ------------------------------Payment Contra JV---------------------------------
        'oItem = f.Items.Add("ck4", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        'oItem.Left = 10
        'oItem.Top = 140
        'oItem.DisplayDesc = True
        'oItem.Width = 300
        'oItem.FromPane = 3
        'oItem.ToPane = 3
        'ock = oItem.Specific
        'ock.DataBind.SetBound(True, "", "ckContraJV")
        'ock.Caption = "Payment Contra Posting JV"

        ' ------------------------------OutStanding DO Contra Act Code---------------------------------

        oItem = f.Items.Add("lbl31", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 140
        oItem.Width = 210
        oItem.FromPane = 3
        oItem.ToPane = 3
        oLabel = oItem.Specific
        oLabel.Caption = "Outstanding DO Contra Account"
        oItem = f.Items.Add("txt31", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 140
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 3
        oItem.ToPane = 3
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "DOAct")
        oEdit.ChooseFromListUID = "clact8"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------OutStanding DO Contra Account name---------------------------------
        oItem = f.Items.Add("lbl32", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 160
        oItem.Width = 210
        oItem.FromPane = 3
        oItem.ToPane = 3
        oLabel = oItem.Specific
        oLabel.Caption = "Outstanding DO Contra Account Name"
        oItem = f.Items.Add("txt32", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 160
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 3
        oItem.ToPane = 3
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "DOName")
    End Sub
    Private Sub AddingItemInReverseMechanism(oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim cp As SAPbouiCOM.FormCreationParams
        Dim oEdit As SAPbouiCOM.EditText
        Dim oFld As SAPbouiCOM.Folder

        ' ------------------------------Enable Reverse Mechanism---------------------------------
        oItem = f.Items.Add("ck2", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Left = 10
        oItem.Top = 40
        oItem.DisplayDesc = True
        oItem.Width = 300
        oItem.FromPane = 4
        oItem.ToPane = 4
        Dim ock2 As SAPbouiCOM.CheckBox
        ock2 = oItem.Specific
        ock2.DataBind.SetBound(True, "", "ckReverse")
        ock2.Caption = "Enable Reverse Mechanism"

        ' ------------------------------Reverse Mechanism OUT TAXCode---------------------------------
        oItem = f.Items.Add("lbl27", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 60
        oItem.Width = 210
        oItem.FromPane = 4
        oItem.ToPane = 4
        oLabel = oItem.Specific
        oLabel.Caption = "Reverse Output Tax Code"
        oItem = f.Items.Add("txt27", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 60
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 4
        oItem.ToPane = 4
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "revouttax")
        oEdit.ChooseFromListUID = "cltax5"
        oEdit.ChooseFromListAlias = "Code"
        ' ------------------------------Reverse Mechanism OUT TAXName---------------------------------
        oItem = f.Items.Add("lbl28", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 80
        oItem.Width = 210
        oItem.FromPane = 4
        oItem.ToPane = 4
        oLabel = oItem.Specific
        oLabel.Caption = "Reverse Ouput Tax Name"
        oItem = f.Items.Add("txt28", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 80
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 4
        oItem.ToPane = 4
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "revouttaxn")

        ' ------------------------------Reverse Mechanism IN TAXCode---------------------------------
        oItem = f.Items.Add("lbl29", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 100
        oItem.Width = 210
        oItem.FromPane = 4
        oItem.ToPane = 4
        oLabel = oItem.Specific
        oLabel.Caption = "Reverse Input Tax Code"
        oItem = f.Items.Add("txt29", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 100
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 4
        oItem.ToPane = 4
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "revintax")
        oEdit.ChooseFromListUID = "cltax6"
        oEdit.ChooseFromListAlias = "Code"

        ' ------------------------------Reverse Mechanism IN TAXName---------------------------------
        oItem = f.Items.Add("lbl30", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 120
        oItem.Width = 210
        oItem.FromPane = 4
        oItem.ToPane = 4
        oLabel = oItem.Specific
        oLabel.Caption = "Reverse Input Tax Name"
        oItem = f.Items.Add("txt30", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 120
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 4
        oItem.ToPane = 4
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "revintaxn")
    End Sub
    Private Sub AddingItemInAPBadDebt(oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim cp As SAPbouiCOM.FormCreationParams
        Dim oEdit As SAPbouiCOM.EditText
        Dim oFld As SAPbouiCOM.Folder

        ' ------------------------------AP Bad Debt Account Code---------------------------------
        oItem = f.Items.Add("lbl7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 40
        oItem.Width = 210
        oItem.FromPane = 2
        oItem.ToPane = 2
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Account"
        oItem = f.Items.Add("txt7", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 40
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 2
        oItem.ToPane = 2
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act7")
        oEdit.ChooseFromListUID = "clact4"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AP Bad Debt Account Name---------------------------------
        oItem = f.Items.Add("lbl8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 60
        oItem.Width = 210
        oItem.FromPane = 2
        oItem.ToPane = 2
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Account Name"
        oItem = f.Items.Add("txt8", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 60
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 2
        oItem.ToPane = 2
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act8")
        ' ------------------------------AP Bad Debt Relief Account---------------------------------
        oItem = f.Items.Add("lbl9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 80
        oItem.Width = 210
        oItem.FromPane = 2
        oItem.ToPane = 2
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Relief Account"
        oItem = f.Items.Add("txt9", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 80
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 2
        oItem.ToPane = 2
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act9")
        oEdit.ChooseFromListUID = "clact5"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AP Bad Debt Relief Account Name---------------------------------
        oItem = f.Items.Add("lbl10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 100
        oItem.Width = 210
        oItem.FromPane = 2
        oItem.ToPane = 2
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Relief Account Name"
        oItem = f.Items.Add("txt10", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 100
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act10")
        ' ------------------------------AP Bad Debt Relief Recovery Account---------------------------------
        oItem = f.Items.Add("lbl11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 120
        oItem.Width = 210
        oItem.FromPane = 2
        oItem.ToPane = 2
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Relief Recovery Account"
        oItem = f.Items.Add("txt11", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 120
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 100
        oItem.DisplayDesc = True
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act11")
        oEdit.ChooseFromListUID = "clact6"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AP Bad Debt Relief Recovery Account Name---------------------------------
        oItem = f.Items.Add("lbl12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 140
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AP Bad Debt Relief Recovery Account Name"
        oItem = f.Items.Add("txt12", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 140
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act12")

        ' ------------------------------AP in TAXCode---------------------------------
        oItem = f.Items.Add("lbl15", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 160
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AP Input Tax Code"
        oItem = f.Items.Add("txt15", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 160
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 100
        oItem.DisplayDesc = True
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "apintax")
        oEdit.ChooseFromListUID = "cltax2"
        oEdit.ChooseFromListAlias = "Code"

        ' ------------------------------AP in TAXName---------------------------------
        oItem = f.Items.Add("lbl16", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 180
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AP Input Tax Name"
        oItem = f.Items.Add("txt16", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 180
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "apinname")

        ' ------------------------------AP OUT TAXCode---------------------------------
        oItem = f.Items.Add("lbl19", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 200
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AP Output Tax Code"
        oItem = f.Items.Add("txt19", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 200
        oItem.Width = 100
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.DisplayDesc = True
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "apouttax")
        oEdit.ChooseFromListUID = "cltax4"
        oEdit.ChooseFromListAlias = "Code"

        ' ------------------------------AP OUT TAXName---------------------------------
        oItem = f.Items.Add("lbl20", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 220
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AP Output Tax Name"
        oItem = f.Items.Add("txt20", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 220
        oItem.FromPane = 2
        oItem.ToPane = 2
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "apoutname")
    End Sub
    Private Sub AddingItemInARBadDebt(oform As SAPbouiCOM.Form)
        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim cp As SAPbouiCOM.FormCreationParams
        Dim oEdit As SAPbouiCOM.EditText
        Dim oFld As SAPbouiCOM.Folder

        oItem = f.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 40
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Account"
        oItem = f.Items.Add("txt1", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 40
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 1
        oItem.ToPane = 1
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act1")
        oEdit.ChooseFromListUID = "clact1"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AR Bad Debt Account Name---------------------------------
        oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 60
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Account Name"
        oItem = f.Items.Add("txt2", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 60
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act2")
        ' ------------------------------AR Bad Debt Relief Account---------------------------------
        oItem = f.Items.Add("lbl3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 80
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Relief Account"
        oItem = f.Items.Add("txt3", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 80
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 1
        oItem.ToPane = 1
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act3")
        oEdit.ChooseFromListUID = "clact2"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AR Bad Debt Relief Account Name---------------------------------
        oItem = f.Items.Add("lbl4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 100
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Relief Account Name"
        oItem = f.Items.Add("txt4", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 100
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 1
        oItem.ToPane = 1
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act4")
        ' ------------------------------AR Bad Debt Relief Recovery Account---------------------------------
        oItem = f.Items.Add("lbl5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 120
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Relief Recovery Account"
        oItem = f.Items.Add("txt5", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 120
        oItem.Width = 100
        oItem.DisplayDesc = True
        oItem.FromPane = 1
        oItem.ToPane = 1
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act5")
        oEdit.ChooseFromListUID = "clact3"
        oEdit.ChooseFromListAlias = "AcctCode"
        ' ------------------------------AR Bad Debt Relief Recovery Account Name---------------------------------
        oItem = f.Items.Add("lbl6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 140
        oItem.Width = 210
        oItem.FromPane = 1
        oItem.ToPane = 1
        oLabel = oItem.Specific
        oLabel.Caption = "AR Bad Debt Relief Recovery Account Name"
        oItem = f.Items.Add("txt6", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 140
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oItem.FromPane = 1
        oItem.ToPane = 1
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "act6")

        ' ------------------------------AR IN TAXCode---------------------------------
        oItem = f.Items.Add("lbl13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 160
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AR Input Tax Code"
        oItem = f.Items.Add("txt13", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 160
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 100
        oItem.DisplayDesc = True
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "arintax")
        oEdit.ChooseFromListUID = "cltax1"
        oEdit.ChooseFromListAlias = "Code"
        ' ------------------------------AR in TAXName---------------------------------
        oItem = f.Items.Add("lbl14", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 180
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AR Input Tax Name"
        oItem = f.Items.Add("txt14", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 180
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "arinname")

        ' ------------------------------AR OUT TAXCode---------------------------------
        oItem = f.Items.Add("lbl17", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 200
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AR Output Tax Code"
        oItem = f.Items.Add("txt17", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 200
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 100
        oItem.DisplayDesc = True
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "arouttax")
        oEdit.ChooseFromListUID = "cltax3"
        oEdit.ChooseFromListAlias = "Code"
        ' ------------------------------AR OUT TAXName---------------------------------
        oItem = f.Items.Add("lbl18", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Left = 10
        oItem.Top = 220
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 210
        oLabel = oItem.Specific
        oLabel.Caption = "AR Ouput Tax Name"
        oItem = f.Items.Add("txt18", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Left = 215
        oItem.Top = 220
        oItem.FromPane = 1
        oItem.ToPane = 1
        oItem.Width = 300
        oItem.DisplayDesc = True
        oItem.Enabled = False
        oEdit = oItem.Specific
        oEdit.DataBind.SetBound(True, "", "aroutname")
    End Sub

#End Region
End Class
