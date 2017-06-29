Imports System.Threading
Imports SAPbobsCOM

Public Class fGSTHeathCheck
    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Sub New(ByVal sbo_application1 As SAPbouiCOM.Application)
        SBO_Application = sbo_application1
        DrawForm()
    End Sub
    Private Sub DrawForm()
        Try
            Dim f As SAPbouiCOM.Form
            Dim oItem As SAPbouiCOM.Item
            Dim oLabel As SAPbouiCOM.StaticText
            Dim cp As SAPbouiCOM.FormCreationParams
            Dim obt As SAPbouiCOM.Button
            Dim oCheck As SAPbouiCOM.CheckBox

            cp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            cp.FormType = "fGSTHeathCheck"
            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 600
            f.ClientHeight = 500

            ' set the form title
            f.Title = "GST Health Check"



            ' ------------------------------0.@GSTSETUP---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[@GSTSETUP]"
            oItem = f.Items.Add("ck0", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 10
            oItem.Width = 90
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------00.Tax Group Description---------------------------------
            oItem = f.Items.Add("lbl00", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 300
            oItem.Top = 10
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Tax Group Description]"
            oItem = f.Items.Add("ck00", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 510
            oItem.Top = 10
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------1.@GSTSETUP.U_Value---------------------------------
            oItem = f.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[@GSTSETUP].[U_Value]"
            oItem = f.Items.Add("ck1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 30
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------11.@GSTSETUP.U_Value---------------------------------
            oItem = f.Items.Add("lbl01", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 300
            oItem.Top = 30
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[SQL Pass]"
            oItem = f.Items.Add("ck01", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 510
            oItem.Top = 30
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------2.@GST_MSIC---------------------------------
            oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 50
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[@GST_MSIC]"
            oItem = f.Items.Add("ck2", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 50
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------3.@GST_MSIC.MSICCode---------------------------------
            oItem = f.Items.Add("lbl3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 70
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[@GST_MSIC].[MSICCode]"
            oItem = f.Items.Add("ck3", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 70
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------4.@GST_MSIC.PERCENTAGE---------------------------------
            oItem = f.Items.Add("lbl4", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 90
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[@GST_MSIC].[PERCENTAGE]"
            oItem = f.Items.Add("ck4", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 90
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------5.[Marketing Documents Title].[21Day]---------------------------------
            oItem = f.Items.Add("lbl5", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 110
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Marketing Documents Title].[U_21Day]"
            oItem = f.Items.Add("ck5", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 110
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------6.[Marketing Documents Title].[21DayJE]---------------------------------
            oItem = f.Items.Add("lbl6", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 130
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Marketing Documents Title].[U_21DayJE]"
            oItem = f.Items.Add("ck6", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 130
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------7.[Marketing Documents Title].[BadDebt]---------------------------------
            oItem = f.Items.Add("lbl7", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 150
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Marketing Documents Title].[U_BadDebt]"
            oItem = f.Items.Add("ck7", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 150
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------8.[Marketing Documents Title].[BadDebtJE]---------------------------------
            oItem = f.Items.Add("lbl8", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 170
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Marketing Documents Title].[U_BadDebtJE]"
            oItem = f.Items.Add("ck8", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 170
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------9.[Journal Entry Header].[BadDebt]---------------------------------
            oItem = f.Items.Add("lbl9", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 190
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Journal Entry Header].[U_BadDebt]"
            oItem = f.Items.Add("ck9", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 190
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------10.[Journal Entry Header].[21Day]---------------------------------
            oItem = f.Items.Add("lbl10", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 210
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Journal Entry Header].[U_21Day]"
            oItem = f.Items.Add("ck10", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 210
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------11.[Journal Entry Header].[ContraPayment]---------------------------------
            oItem = f.Items.Add("lbl11", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 230
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Journal Entry Header].[U_ContraPayment]"
            oItem = f.Items.Add("ck11", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 230
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------12.[Journal Entry Line].[ContraPayment]---------------------------------
            oItem = f.Items.Add("lbl12", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 250
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[Journal Entry Line].[U_InvoiceEntry]"
            oItem = f.Items.Add("ck12", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 250
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------13.[sp_SAPB1Addon_GST03]---------------------------------
            oItem = f.Items.Add("lbl13", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 270
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_SAPB1Addon_GST03]"
            oItem = f.Items.Add("ck13", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 270
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------14.[sp_B1Addon_BadDebtReverse]---------------------------------
            oItem = f.Items.Add("lbl14", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 290
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_B1Addon_BadDebtReverse]"
            oItem = f.Items.Add("ck14", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 290
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------15.[sp_SAPB1Addon_GSTBadDebt]---------------------------------
            oItem = f.Items.Add("lbl15", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 310
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_SAPB1Addon_GSTBadDebt]"
            oItem = f.Items.Add("ck15", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 310
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------16.[sp_SAPB1Addon_PaymentContra]---------------------------------
            oItem = f.Items.Add("lbl16", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 330
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_SAPB1Addon_PaymentContra]"
            oItem = f.Items.Add("ck16", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 330
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------17.[sp_SAPB1Addon_21DO]---------------------------------
            oItem = f.Items.Add("lbl17", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 350
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_SAPB1Addon_21DO]"
            oItem = f.Items.Add("ck17", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 350
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------18.[sp_B1Addon_GSTReturn]---------------------------------
            oItem = f.Items.Add("lbl18", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 370
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_B1Addon_GSTReturn]"
            oItem = f.Items.Add("ck18", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 370
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------19.[sp_SAPB1Addon_GSTBadDebt_AP]---------------------------------
            oItem = f.Items.Add("lbl19", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 390
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_SAPB1Addon_GSTBadDebt_AP]"
            oItem = f.Items.Add("ck19", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 390
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------20.[sp_B1Addon_BadDebtReverse_AP]---------------------------------
            oItem = f.Items.Add("lbl20", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 410
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_B1Addon_BadDebtReverse_AP]"
            oItem = f.Items.Add("ck20", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 410
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' ------------------------------21.[sp_B1Addon_GAF]---------------------------------
            oItem = f.Items.Add("lbl21", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 430
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_B1Addon_GAF]"
            oItem = f.Items.Add("ck21", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 430
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False
            ' ------------------------------22.[sp_B1Addon_GAFtxt]---------------------------------
            oItem = f.Items.Add("lbl22", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 450
            oItem.Width = 205
            oLabel = oItem.Specific
            oLabel.Caption = "[sp_B1Addon_GAFtxt]"
            oItem = f.Items.Add("ck22", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 210
            oItem.Top = 450
            oItem.Width = 300
            oItem.DisplayDesc = True
            oItem.Enabled = False
            oCheck = oItem.Specific
            oCheck.Caption = ""
            oCheck.Checked = False

            ' -------------------Add the OK button----------------------
            oItem = f.Items.Add("btnRun", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 470
            obt = oItem.Specific
            obt.Caption = "Run"
            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 470

            f.Visible = True
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DrawForm Event: " + ex.ToString, , True)
        End Try
    End Sub

    Private Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction = True Then
                Dim oForm As SAPbouiCOM.Form = Nothing
                If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    oForm = SBO_Application.Forms.Item(FormUID)
                End If
                If pVal.FormTypeEx = "fGSTHeathCheck" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                        SBO_Application = Nothing
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "btnRun" Then
                            InitData(oForm)
                            CheckTaxGroupDes(oForm)
                            CheckSQLConnection(oForm)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("ItemEvent Event: " + ex.ToString, , True)
        End Try
    End Sub
    Public Sub InitData(ByVal f As SAPbouiCOM.Form)
        Try
            Dim oCheck As SAPbouiCOM.CheckBox
            Dim cli As New InitData
            Dim Str As String = ""
            Dim CheckStore As String
            Dim ReturnStr As String = ""
            '------------------------------0.GSTSETUP------------------------
            oCheck = f.Items.Item("ck0").Specific
            If Not cli.CheckTableExists("@GSTSETUP") Then
                ReturnStr = cli.CreateUDT("GSTSETUP", "GST SETUP", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------1.@GSTSETUP.U_Value---------------------------------
            oCheck = f.Items.Item("ck1").Specific
            If Not cli.CheckFieldExists("@GSTSETUP", "U_Value") Then
                ReturnStr = cli.CreateUDF("GSTSETUP", "Value", "Value", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------2.@GST_MSIC---------------------------------
            oCheck = f.Items.Item("ck2").Specific
            If Not cli.CheckTableExists("@GST_MSIC") Then
                ReturnStr = cli.CreateUDT("GST_MSIC", "GST MSIC", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------3.@GST_MSIC.MSICCode---------------------------------
            oCheck = f.Items.Item("ck3").Specific
            If Not cli.CheckFieldExists("@GST_MSIC", "U_MSICCode") Then
                ReturnStr = cli.CreateUDF("GST_MSIC", "MSICCode", "MSICCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------4.@GST_MSIC.PERCENTAGE---------------------------------
            oCheck = f.Items.Item("ck4").Specific
            If Not cli.CheckFieldExists("@GST_MSIC", "U_PERCENTAGE") Then
                ReturnStr = cli.CreateUDF("GST_MSIC", "PERCENTAGE", "PERCENTAGE", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------5.[Marketing Documents Title].[21Day]---------------------------------
            oCheck = f.Items.Item("ck5").Specific
            If Not cli.CheckFieldExists("ORDR", "U_21Day") Then
                ReturnStr = cli.CreateUDF("ORDR", "21Day", "21 Days Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------6.[Marketing Documents Title].[21DayJE]---------------------------------
            oCheck = f.Items.Item("ck6").Specific
            If Not cli.CheckFieldExists("ORDR", "U_21DayJE") Then
                ReturnStr = cli.CreateUDF("ORDR", "21DayJE", "21 Days Applied JE", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------7.[Marketing Documents Title].[BadDebt]---------------------------------
            oCheck = f.Items.Item("ck7").Specific
            If Not cli.CheckFieldExists("ORDR", "U_BadDebt") Then
                ReturnStr = cli.CreateUDF("ORDR", "BadDebt", "BadDebt", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------8.[Marketing Documents Title].[BadDebtJE]---------------------------------
            oCheck = f.Items.Item("ck8").Specific
            If Not cli.CheckFieldExists("ORDR", "U_BadDebtJE") Then
                ReturnStr = cli.CreateUDF("ORDR", "BadDebtJE", "BadDebtJE", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------9.[Journal Entry Header].[BadDebt]---------------------------------
            oCheck = f.Items.Item("ck9").Specific
            If Not cli.CheckFieldExists("ORDR", "U_BadDebt") Then
                ReturnStr = cli.CreateUDF("ORDR", "BadDebt", "BadDebt", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------10.[Journal Entry Header].[21Day]---------------------------------
            oCheck = f.Items.Item("ck10").Specific
            If Not cli.CheckFieldExists("OJDT", "U_21Day") Then
                ReturnStr = cli.CreateUDF("OJDT", "21Day", "21 Days Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------11.[Journal Entry Header].[ContraPayment]---------------------------------
            oCheck = f.Items.Item("ck11").Specific
            If Not cli.CheckFieldExists("OJDT", "U_ContraPayment") Then
                ReturnStr = cli.CreateUDF("OJDT", "ContraPayment", "ContraPayment", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True
            ' ------------------------------12.[Journal Entry Line].[InvoiceEntry]---------------------------------
            oCheck = f.Items.Item("ck12").Specific
            If Not cli.CheckFieldExists("JDT1", "U_InvoiceEntry") Then
                ReturnStr = cli.CreateUDF("JDT1", "InvoiceEntry", "InvoiceEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")
                If ReturnStr <> "" Then
                    oCheck.Caption = "Not Found. Error: " + ReturnStr
                Else
                    oCheck.Caption = "Not Found. Created"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True


            '-------------function-----------------

            CheckStore = cli.CheckFunctionExists("uf_GetTaxBalance")
            If CheckStore <> "" Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("drop function uf_GetTaxBalance;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_uf_GetTaxBalance, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("drop function uf_GetTaxBalance")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.uf_GetTaxBalance, PublicVariable.Version, CheckStore))
                End If
            End If
            CheckStore = cli.CheckFunctionExists("uf_GetTaxBalance_AP")
            If CheckStore <> "" Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP FUNCTION uf_GetTaxBalance_AP;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_uf_GetTaxBalance_AP, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("drop function uf_GetTaxBalance_AP")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.uf_GetTaxBalance_AP, PublicVariable.Version, CheckStore))
                End If
            End If
            '-------------function-----------------

            ' ------------------------------13.[sp_SAPB1Addon_GST03]---------------------------------
            oCheck = f.Items.Item("ck13").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_SAPB1Addon_GST03")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_SAPB1ADDON_GST03")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_SAPB1Addon_GST03, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE SP_SAPB1ADDON_GST03")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_SAPB1Addon_GST03, PublicVariable.Version, CheckStore))
                End If
                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            'If Not cli.CheckTableExists("@ADDONSCRIPT") Then
            '    ReturnStr = cli.CreateUDT("ADDONSCRIPT", "ADDON SCRIPT", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            '    If ReturnStr <> "" Then
            '        SBO_Application.SetStatusBarMessage(ReturnStr, , True)
            '    Else
            '        SBO_Application.SetStatusBarMessage("GSTSETUP is created", , False)
            '    End If
            'End If

            'If Not cli.CheckFieldExists("@ADDONSCRIPT", "U_Value") Then
            '    ReturnStr = cli.CreateUDF("ADDONSCRIPT", "Value", "Value", SAPbobsCOM.BoFieldTypes.db_Memo, 5000, "", BoFldSubTypes.st_None)
            '    If ReturnStr <> "" Then
            '        SBO_Application.SetStatusBarMessage(ReturnStr, , True)
            '    Else
            '        SBO_Application.SetStatusBarMessage("GSTSETUP-Value is created", , False)
            '    End If
            'End If
            ' ------------------------------14.[sp_B1Addon_BadDebtReverse]---------------------------------
            oCheck = f.Items.Item("ck14").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_B1Addon_BadDebtReverse")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_B1ADDON_BADDEBTREVERSE;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_B1Addon_BadDebtReverse, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_B1Addon_BadDebtReverse")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_B1Addon_BadDebtReverse, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------15.[sp_SAPB1Addon_GSTBadDebt]---------------------------------
            oCheck = f.Items.Item("ck15").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_SAPB1Addon_GSTBadDebt")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_SAPB1ADDON_GSTBADDEBT;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_SAPB1Addon_GSTBadDebt, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_SAPB1Addon_GSTBadDebt")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_SAPB1Addon_GSTBadDebt, PublicVariable.Version, CheckStore))
                End If
                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------16.[sp_SAPB1Addon_PaymentContra]---------------------------------
            oCheck = f.Items.Item("ck16").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_SAPB1Addon_PaymentContra")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE sp_SAPB1Addon_PaymentContra;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_SAPB1Addon_PaymentContra, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_SAPB1Addon_PaymentContra")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_SAPB1Addon_PaymentContra, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------17.[sp_SAPB1Addon_21DO]---------------------------------
            oCheck = f.Items.Item("ck17").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_SAPB1Addon_21DO")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_SAPB1ADDON_21DO;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_SAPB1Addon_21DO, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_SAPB1Addon_21DO")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_SAPB1Addon_21DO, PublicVariable.Version, CheckStore))
                End If


                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------18.[sp_B1Addon_GSTReturn]---------------------------------
            oCheck = f.Items.Item("ck18").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_B1Addon_GSTReturn")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE sp_B1Addon_GSTReturn;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_B1Addon_GSTReturn, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_B1Addon_GSTReturn")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_B1Addon_GSTReturn, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------19.[sp_SAPB1Addon_GSTBadDebt_AP]---------------------------------
            oCheck = f.Items.Item("ck19").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_SAPB1Addon_GSTBadDebt_AP")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_SAPB1ADDON_GSTBADDEBT_AP;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_SAPB1Addon_GSTBadDebt_AP, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_SAPB1Addon_GSTBadDebt_AP")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_SAPB1Addon_GSTBadDebt_AP, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------20.[sp_B1Addon_BadDebtReverse_AP]---------------------------------
            oCheck = f.Items.Item("ck20").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_B1Addon_BadDebtReverse_AP")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE SP_B1ADDON_BADDEBTREVERSE_AP;")
                    Functions.Hana_RunQuery(String.Format(My.Resources.B1H_sp_B1Addon_BadDebtReverse_AP, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_B1Addon_BadDebtReverse_AP")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_B1Addon_BadDebtReverse_AP, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------21.[sp_B1Addon_GAF]---------------------------------
            oCheck = f.Items.Item("ck21").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_B1Addon_GAF")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    'Functions.Hana_RunQuery("DROP PROCEDURE sp_B1Addon_GAF;")
                    'Functions.DoQueryReturnDT(String.Format(My.Resources.B1H_sp_B1Addon_GAF, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_B1Addon_GAF")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_B1Addon_GAF, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True

            ' ------------------------------22.[sp_B1Addon_GAFtxt]---------------------------------
            oCheck = f.Items.Item("ck22").Specific
            CheckStore = cli.CheckStoreProcedureExists("sp_B1Addon_GAFtxt")
            If CheckStore = "" Or CheckStore <> PublicVariable.Version Then
                If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Functions.Hana_RunQuery("DROP PROCEDURE sp_B1Addon_GAFtxt;")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.B1H_sp_B1Addon_GAFtxt, PublicVariable.Version))
                Else
                    Functions.DoQueryReturnDT("DROP PROCEDURE sp_B1Addon_GAFtxt")
                    Functions.DoQueryReturnDT(String.Format(My.Resources.sp_B1Addon_GAFtxt, PublicVariable.Version, CheckStore))
                End If

                If CheckStore = "" Then
                    oCheck.Caption = "Not Found. Created"
                Else
                    oCheck.Caption = "Old Version:" + CheckStore + ". Updated"
                End If
            Else
                oCheck.Caption = "Found. OK"
            End If
            oCheck.Checked = True


            



        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("Init Error:" + ex.ToString, , True)
        End Try
    End Sub

    Private Sub CheckTaxGroupDes(ByVal f As SAPbouiCOM.Form)
        Dim oCheck As SAPbouiCOM.CheckBox
        oCheck = f.Items.Item("ck00").Specific
        Dim dt As DataTable = Functions.DoQueryReturnDT("select * from OVTG where ISNULL(reportcode,'')<>''")
        If IsNothing(dt) Then
            oCheck.Caption = "Not Found"
        Else
            If dt.Rows.Count = 0 Then
                oCheck.Caption = "Not Found"
            Else
                oCheck.Caption = "Found"
            End If
        End If
        
        oCheck.Checked = True
    End Sub

    Private Sub CheckSQLConnection(ByVal f As SAPbouiCOM.Form)
        Dim oCheck As SAPbouiCOM.CheckBox
        oCheck = f.Items.Item("ck01").Specific
        Dim oUserTable As SAPbobsCOM.UserTable
        oUserTable = PublicVariable.oCompany.UserTables.Item("GSTSETUP")
        If oUserTable.GetByKey("SQLPass") Then
            Dim SQLPass As String = oUserTable.UserFields.Fields.Item("U_Value").Value.ToString
            If IsNothing(Functions.ADO_OpenSQLConnection()) Then
                oCheck.Caption = "Incorrect."
            Else
                oCheck.Caption = "Passed."
            End If
        Else
            oCheck.Caption = "Not found."
        End If

        oCheck.Checked = True
    End Sub
End Class
