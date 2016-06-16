Imports System.Data
Module modCustomerSelection

    Private objForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oRecordSet As SAPbobsCOM.Recordset

#Region "Load Customer Selection Form"
    Public Sub InitializeCustSelectionForm(ByVal sId As String, ByVal sLine As String, ByVal dAmt As Double, ByVal sInvRefNo As String, ByVal sRandomNo As String)
        Dim sFuncName As String = "InitializeCustSelectionForm"
        Dim sErrDesc As String = String.Empty
        Try
            LoadFromXML("Customer Selection.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("CSS")
            objForm.Visible = True
            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            objForm.EnableMenu("6913", False) 'User Defined windows
            objForm.EnableMenu("1290", False) 'Move First Record
            objForm.EnableMenu("1288", False) 'Move Next Record
            objForm.EnableMenu("1289", False) 'Move Previous Record
            objForm.EnableMenu("1291", False) 'Move Last Record
            objForm.EnableMenu("1281", True) 'Find Record
            objForm.EnableMenu("1282", False) 'Add New Record
            objForm.EnableMenu("1292", True) 'Add New Row

            AddUserDatasources(objForm)
            objForm.DataBrowser.BrowseBy = "14"

            GenerateDocNum(objForm)

            oMatrix = objForm.Items.Item("11").Specific
            oMatrix.AddRow(1)
            oMatrix.AutoResizeColumns()

            AddChooseFromList(objForm)
            CFLDataBinding(objForm)

            LoadDefValues(objForm, sId, sInvRefNo, sLine, dAmt, sRandomNo)

            objForm.Freeze(False)
            objForm.Update()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Load Customer Selection Form in FindMode"
    Public Sub CustSelectionFindForm(ByVal sDocNo As String, ByVal sPayDocNo As String)
        Dim sFuncName As String = "InitializeCustSelectionForm"
        Dim sErrDesc As String = String.Empty
        Dim sSQL As String = String.Empty

        Try
            LoadFromXML("Customer Selection.srf", p_oSBOApplication)
            objForm = p_oSBOApplication.Forms.Item("CSS")
            objForm.Visible = True
            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            objForm.EnableMenu("6913", False) 'User Defined windows
            objForm.EnableMenu("1290", False) 'Move First Record
            objForm.EnableMenu("1288", False) 'Move Next Record
            objForm.EnableMenu("1289", False) 'Move Previous Record
            objForm.EnableMenu("1291", False) 'Move Last Record
            objForm.EnableMenu("1281", False) 'Find Record
            objForm.EnableMenu("1282", False) 'Add New Record
            objForm.EnableMenu("1292", True) 'Add New Row

            AddUserDatasources(objForm)
            objForm.DataBrowser.BrowseBy = "14"

            AddChooseFromList(objForm)
            CFLDataBinding(objForm)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            objForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Items.Item("14").Enabled = False
            objForm.Items.Item("4").Enabled = False
            objForm.Items.Item("6").Enabled = False
            objForm.Items.Item("8").Enabled = False
            objForm.Items.Item("10").Enabled = False
            objForm.Items.Item("14").Enabled = False
            objForm.Items.Item("11").Enabled = False

            Dim oDt As New DataTable
            sSQL = "SELECT * FROM AB_SELECTEDCUSTOMER WHERE DOCNUM = '" & sDocNo & "'"
            oDt = ExecuteSQLQueryDataTable(sSQL, sErrDesc)

            If Not oDt Is Nothing Then
                If oDt.Rows.Count >= 1 Then
                    oMatrix = objForm.Items.Item("11").Specific
                    oMatrix.Clear()
                    For Each oDr As DataRow In oDt.Rows
                        oMatrix.AddRow(1)
                        If oMatrix.RowCount = 1 Then
                            oEdit = objForm.Items.Item("14").Specific
                            oEdit.Value = oDr("DOCNUM").ToString.Trim()
                            oEdit = objForm.Items.Item("4").Specific
                            oEdit.Value = oDr("ID").ToString.Trim()
                            oEdit = objForm.Items.Item("6").Specific
                            oEdit.Value = oDr("LINE").ToString.Trim()
                            oEdit = objForm.Items.Item("10").Specific
                            oEdit.Value = oDr("INVREFNO").ToString.Trim()
                            oEdit = objForm.Items.Item("8").Specific
                            oEdit.Value = oDr("AMOUNT").ToString.Trim()
                            oEdit = objForm.Items.Item("12").Specific
                            oEdit.Value = oDr("RANDOMNO").ToString.Trim()
                        End If
                        oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.value = oMatrix.RowCount
                        oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = oDr("CUSTCODE").ToString.Trim()
                        oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = oDr("CUSTNAME").ToString.Trim()
                        oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.value = oDr("INVDOCENTRY").ToString.Trim()
                        oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific.value = oDr("CUSTAMT").ToString.Trim()
                    Next
                End If
            End If

            If sPayDocNo = "" Then
                objForm.Items.Item("11").Enabled = True
            End If

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

            objForm.Freeze(False)
            objForm.Update()
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Add Datasources to form"
    Private Sub AddUserDatasources(ByVal objForm As SAPbouiCOM.Form)
        objForm.DataSources.UserDataSources.Add("uID", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit = objForm.Items.Item("4").Specific
        oEdit.DataBind.SetBound(True, "", "uID")

        objForm.DataSources.UserDataSources.Add("uItmLine", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
        oEdit = objForm.Items.Item("6").Specific
        oEdit.DataBind.SetBound(True, "", "uItmLine")

        objForm.DataSources.UserDataSources.Add("uAmount", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oEdit = objForm.Items.Item("8").Specific
        oEdit.DataBind.SetBound(True, "", "uAmount")

        objForm.DataSources.UserDataSources.Add("uInvRefNo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit = objForm.Items.Item("10").Specific
        oEdit.DataBind.SetBound(True, "", "uInvRefNo")

        objForm.DataSources.UserDataSources.Add("uRandNo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit = objForm.Items.Item("12").Specific
        oEdit.DataBind.SetBound(True, "", "uRandNo")

        objForm.DataSources.UserDataSources.Add("uDocNo", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
        oEdit = objForm.Items.Item("14").Specific
        oEdit.DataBind.SetBound(True, "", "uDocNo")

        oMatrix = objForm.Items.Item("11").Specific
        objForm.DataSources.UserDataSources.Add("uLineId", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
        oMatrix.Columns.Item("V_-1").DataBind.SetBound(True, "", "uLineId")

        objForm.DataSources.UserDataSources.Add("uCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_0").DataBind.SetBound(True, "", "uCardCode")

        objForm.DataSources.UserDataSources.Add("uCardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_1").DataBind.SetBound(True, "", "uCardName")

        objForm.DataSources.UserDataSources.Add("uInvEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
        oMatrix.Columns.Item("V_2").DataBind.SetBound(True, "", "uInvEntry")

        objForm.DataSources.UserDataSources.Add("uLineAmt", SAPbouiCOM.BoDataType.dt_PRICE, 20)
        oMatrix.Columns.Item("V_3").DataBind.SetBound(True, "", "uLineAmt")

    End Sub
#End Region
#Region "Generate Document number"
    Private Sub GenerateDocNum(ByVal objForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        Dim sSQL As String
        Dim oDs As New DataSet

        sSQL = "SELECT COALESCE(MAX(DOCNUM),0) + 1 ""DOCNUM"" FROM AB_SELECTEDCUSTOMER "
        oDs = ExecuteSQLQueryDataset(sSQL, sErrDesc)

        If oDs.Tables(0).Rows.Count > 0 Then
            oEdit = objForm.Items.Item("14").Specific
            oEdit.Value = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        objForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        objForm.Items.Item("14").Enabled = False

    End Sub
#End Region
#Region "Choose From list Code"
    Private Sub AddChooseFromList(ByRef objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Customer Code
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "frozenFor"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

            'Invoice DocEntry
            oCFLCreationParams.ObjectType = "13"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = "CANCELED"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CFLDataBinding(ByRef objForm As SAPbouiCOM.Form)
        'Customer Code
        oMatrix = objForm.Items.Item("11").Specific
        oMatrix.Columns.Item("V_0").ChooseFromListUID = "CFL1"
        oMatrix.Columns.Item("V_0").ChooseFromListAlias = "CardCode"

        'Invoice DocEntry
        oMatrix = objForm.Items.Item("11").Specific
        oMatrix.Columns.Item("V_2").ChooseFromListUID = "CFL2"
        oMatrix.Columns.Item("V_2").ChooseFromListAlias = "DocEntry"
    End Sub
#End Region
#Region "Load predefined values"
    Private Sub LoadDefValues(ByVal objForm As SAPbouiCOM.Form, ByVal sId As String, ByVal sInvRefNo As String, ByVal sLine As String, ByVal dAmt As Double, ByVal sRandomNo As String)
        oEdit = objForm.Items.Item("4").Specific
        oEdit.Value = sId
        oEdit = objForm.Items.Item("6").Specific
        oEdit.Value = sLine
        oEdit = objForm.Items.Item("10").Specific
        oEdit.Value = sInvRefNo
        oEdit = objForm.Items.Item("8").Specific
        oEdit.Value = dAmt
        oEdit = objForm.Items.Item("12").Specific
        oEdit.Value = sRandomNo

        objForm.Items.Item("4").Enabled = False
        objForm.Items.Item("6").Enabled = False
        objForm.Items.Item("8").Enabled = False
        objForm.Items.Item("10").Enabled = False
    End Sub
#End Region
#Region "Clear Matrix rows"
    Private Sub ClearMatrixRow(ByVal objForm As SAPbouiCOM.Form)
        oMatrix = objForm.Items.Item("11").Specific
        oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific.value = 0.0
        oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.value = ""
        oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.value = ""
        oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = ""
    End Sub
#End Region
#Region "Check Fields before add"
    Private Function CheckAllFields(ByVal objForm As SAPbouiCOM.Form, ByRef sErrDesc As String) As Boolean
        Dim bCheck As Boolean
        bCheck = True
        sErrDesc = ""

        oMatrix = objForm.Items.Item("11").Specific
        If oMatrix.RowCount > 0 Then
            If oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.value = "" Then
                oMatrix.DeleteRow(oMatrix.RowCount)
            End If
        End If

        If oMatrix.RowCount = 0 Then
            objForm.Freeze(True)
            oMatrix.AddRow(1)
            ClearMatrixRow(objForm)
            objForm.Freeze(False)

            bCheck = False
            sErrDesc = "Atleast one row should be added in a matrix"
            Return bCheck
            Exit Function
        End If

        For i As Integer = 1 To oMatrix.RowCount
            Dim sCustCode As String = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value
            If sCustCode <> "" Then
                For k As Integer = 1 To oMatrix.RowCount
                    Dim sCustomer As String = oMatrix.Columns.Item("V_0").Cells.Item(k).Specific.value
                    If sCustomer <> "" Then
                        If i <> k Then
                            If sCustomer = sCustCode Then
                                bCheck = False
                                sErrDesc = "Duplicate Customer Code is not allowed in matrix/Check Line " & i & " and " & k
                                Return bCheck
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next

        Dim dHAmount, dLAmount, dTotLAmt As Double
        oEdit = objForm.Items.Item("8").Specific
        dHAmount = oEdit.Value
        For i As Integer = 1 To oMatrix.RowCount
            If oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value <> "" Then
                Try
                    dLAmount = oMatrix.Columns.Item("V_3").Cells.Item(i).Specific.value
                    If dLAmount <= 0.0 Then
                        bCheck = False
                        sErrDesc = "Amount value should be greater than zero / Check Line : " & i
                        Return bCheck
                        Exit Function
                    End If
                Catch ex As Exception
                    bCheck = False
                    sErrDesc = "Amount value should be greater than zero / Check Line : " & i
                    Return bCheck
                    Exit Function
                End Try
                dTotLAmt = dTotLAmt + dLAmount
            End If
        Next

        If Math.Round(dHAmount, 2) <> Math.Round(dTotLAmt, 2) Then
            bCheck = False
            sErrDesc = "Sum of Matrix line amount should be equal to Header amount"
            Return bCheck
            Exit Function
        End If

        Return bCheck
    End Function
#End Region
#Region "Insert into Temp Table"
    Private Sub InsertIntoTempTable(ByVal objForm As SAPbouiCOM.Form)
        Dim sSql As String = String.Empty
        Dim sDocNo, sId, sLine, sInvRefNo, sRandomNo, sCustCode, sCustName, sInvEntry As String
        Dim dHAmount, dLAmount As Double
        Dim iLine As Integer

        oEdit = objForm.Items.Item("4").Specific
        sId = oEdit.Value
        oEdit = objForm.Items.Item("6").Specific
        sLine = oEdit.Value
        iLine = oEdit.Value
        oEdit = objForm.Items.Item("10").Specific
        sInvRefNo = oEdit.Value
        oEdit = objForm.Items.Item("8").Specific
        dHAmount = oEdit.Value
        oEdit = objForm.Items.Item("12").Specific
        sRandomNo = oEdit.Value
        oEdit = objForm.Items.Item("14").Specific
        sDocNo = oEdit.Value

        sSql = "DELETE FROM AB_SELECTEDCUSTOMER WHERE RANDOMNO = '" & sRandomNo & "' AND ID = '" & sId & "' AND INVREFNO = '" & sInvRefNo & "' AND LINE = '" & sLine & "'"
        If ExecuteSQLNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        oMatrix = objForm.Items.Item("11").Specific
        For i As Integer = 1 To oMatrix.RowCount
            sCustCode = oMatrix.Columns.Item("V_0").Cells.Item(i).Specific.value
            If sCustCode <> "" Then
                sCustName = oMatrix.Columns.Item("V_1").Cells.Item(i).Specific.value
                sInvEntry = oMatrix.Columns.Item("V_2").Cells.Item(i).Specific.value
                dLAmount = oMatrix.Columns.Item("V_3").Cells.Item(i).Specific.value

                sSql = "INSERT INTO AB_SELECTEDCUSTOMER(RANDOMNO,DOCNUM,ID,INVREFNO,LINE,AMOUNT,CUSTCODE,CUSTNAME,CUSTAMT,INVDOCENTRY)" & _
                       " VALUES('" & sRandomNo & "','" & sDocNo & "','" & sId & "','" & sInvRefNo & "','" & sLine & "','" & dHAmount & "', " & _
                       " '" & sCustCode & "','" & sCustName & "','" & dLAmount & "','" & sInvEntry & "') "

                If ExecuteSQLNonQuery(sSql, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If
        Next

        Dim oForm As SAPbouiCOM.Form
        oForm = p_oSBOApplication.Forms.Item("EXPL")
        oMatrix = oForm.Items.Item("10").Specific
        oMatrix.Columns.Item("V_19").Editable = True
        oMatrix.Columns.Item("V_19").Cells.Item(iLine).Specific.value = sDocNo
        oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oMatrix.Columns.Item("V_19").Editable = False
    End Sub
#End Region
    
#Region "Item Event"
    Public Sub CustSelection_SBO_ItemEvent(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal objForm As SAPbouiCOM.Form)
        Dim sFuncName As String = "CustSelection_SBO_ItemEvent"
        Dim sErrDesc As String = String.Empty
        Try
            If pval.Before_Action = True Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If CheckAllFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    InsertIntoTempTable(objForm)
                                End If
                            ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If CheckAllFields(objForm, sErrDesc) = False Then
                                    p_oSBOApplication.StatusBar.SetText(sErrDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    InsertIntoTempTable(objForm)
                                End If
                            End If
                        End If

                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.CharPressed = "9" Then
                            If pval.ItemUID = "11" Then
                                oMatrix = objForm.Items.Item("11").Specific
                                If pval.ColUID = "V_0" Then
                                    If oMatrix.Columns.Item("V_0").Cells.Item(pval.Row).Specific.value <> "" Then
                                        If pval.Row = oMatrix.RowCount Then
                                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                objForm.Freeze(True)
                                                oMatrix.AddRow(1)
                                                ClearMatrixRow(objForm)
                                                oMatrix.Columns.Item("V_2").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                objForm.Freeze(False)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        If pval.ItemUID = "1" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                objForm.Close()
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        objForm = p_oSBOApplication.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount)
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pval
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = objForm.ChooseFromLists.Item(sCFL_ID)
                        Try
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects
                                If pval.ItemUID = "11" Then
                                    oMatrix = objForm.Items.Item("11").Specific
                                    If pval.Row = 1 Then
                                        oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row).Specific.value = 1
                                    Else
                                        Dim intSno As Integer
                                        intSno = oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row - 1).Specific.value
                                        intSno = intSno + 1
                                        oMatrix.Columns.Item("V_-1").Cells.Item(pval.Row).Specific.value = intSno
                                    End If
                                    If pval.ColUID = "V_0" Then
                                        oMatrix.Columns.Item("V_1").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("CardName", 0)
                                        oMatrix.Columns.Item("V_0").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("CardCode", 0)
                                    ElseIf pval.ColUID = "V_2" Then
                                        oMatrix.Columns.Item("V_3").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("DocTotal", 0)
                                        oMatrix.Columns.Item("V_2").Cells.Item(pval.Row).Specific.value = oDataTable.GetValue("DocEntry", 0)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objForm.Update()
                        End Try

                End Select
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region
#Region "Menu Event"
    Public Sub CustSelection_SBO_MenuEvent(ByVal pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Dim objForm As SAPbouiCOM.Form
                If pVal.MenuUID = "1292" Then
                    objForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.UniqueID)
                    oMatrix = objForm.Items.Item("11").Specific
                    If oMatrix.RowCount = 0 Then
                        objForm.Freeze(True)
                        oMatrix.AddRow(1)
                        ClearMatrixRow(objForm)
                        objForm.Freeze(False)
                    Else
                        objForm.Freeze(True)
                        oMatrix.AddRow(1)
                        ClearMatrixRow(objForm)
                        objForm.Freeze(False)
                    End If
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Menu event error", sFuncName)
            Throw New ArgumentException(sErrDesc)
        End Try
    End Sub
#End Region

End Module
