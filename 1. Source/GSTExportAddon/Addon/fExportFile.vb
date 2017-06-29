Imports System.Threading

Public Class fExportFile
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
            Dim oEdit As SAPbouiCOM.EditText
            Dim oCombo As SAPbouiCOM.ComboBox
            Dim obt As SAPbouiCOM.Button

            cp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            cp.FormType = "fExportFile"
            f = SBO_Application.Forms.AddEx(cp)
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            ' Defining form dimentions
            f.ClientWidth = 290
            f.ClientHeight = 110

            ' set the form title
            f.Title = "GST Export"

            f.DataSources.UserDataSources.Add("frdateds", SAPbouiCOM.BoDataType.dt_DATE)
            f.DataSources.UserDataSources.Add("todateds", SAPbouiCOM.BoDataType.dt_DATE)

            ' ------------------------------Type---------------------------------
            oItem = f.Items.Add("lbl0", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 10
            oLabel = oItem.Specific
            oLabel.Caption = "File Type"
            oItem = f.Items.Add("cbType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = 95
            oItem.Top = 10
            oItem.Width = 100
            oItem.DisplayDesc = True
            oCombo = oItem.Specific
            oCombo.ValidValues.Add("1", "GST Audit File (xml)")
            oCombo.ValidValues.Add("3", "GST Audit File (txt)")
            oCombo.ValidValues.Add("2", "GST Tap Return File")

            ' ------------------------------From Date---------------------------------
            oItem = f.Items.Add("lbl1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 30
            oLabel = oItem.Specific
            oLabel.Caption = "From Date"
            oItem = f.Items.Add("txtFrDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 30
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "frdateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")
            ' ------------------------------To Date---------------------------------
            oItem = f.Items.Add("lbl2", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = 10
            oItem.Top = 50
            oLabel = oItem.Specific
            oLabel.Caption = "To Date"
            oItem = f.Items.Add("txtToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = 95
            oItem.Top = 50
            oItem.Width = 100
            oItem.DisplayDesc = True
            oEdit = oItem.Specific
            oEdit.DataBind.SetBound(True, "", "todateds")
            oEdit.Value = Format(Date.Today, "yyyyMMdd")

            ' -------------------Add the OK button----------------------
            oItem = f.Items.Add("btnExport", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 10
            oItem.Top = 70
            obt = oItem.Specific
            obt.Caption = "Export"
            ' -------------------Add the cancel button----------------------
            oItem = f.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 70

            f.Visible = True
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("DrawForm Event: " + ex.ToString, , True)
        End Try
    End Sub

    Private Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            If pVal.BeforeAction = False Then
                If pVal.FormTypeEx = "fExportFile" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                        SBO_Application = Nothing
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "btnExport" Then
                            If PublicVariable.SAPPass = "" Then
                                SBO_Application.MessageBox("Please configure SQL Password!")
                                Return
                            End If
                            Dim cb As SAPbouiCOM.ComboBox

                            cb = SBO_Application.Forms.Item(FormUID).Items.Item("cbType").Specific
                            Dim frdate As String
                            Dim todate As String

                            Dim oedit As SAPbouiCOM.EditText
                            Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                            Dim format As String = "yyyyMMdd"
                            '-----FROM DATE---------
                            oedit = SBO_Application.Forms.Item(FormUID).Items.Item("txtFrDate").Specific
                            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                frdate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
                            Else
                                frdate = oedit.Value
                            End If

                            '-----To DATE---------
                            oedit = SBO_Application.Forms.Item(FormUID).Items.Item("txtToDate").Specific
                            If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                todate = DateTime.ParseExact(oedit.Value.ToString, format, provider).ToString("yyyyMMdd")
                            Else
                                todate = oedit.Value
                            End If

                            If cb.Selected.Value = "1" Then
                                Dim f As New GetFileNameClass(1)
                                f.Filter = "Xml files|*.xml"
                                f.InitialDirectory = Environment.SpecialFolder.Desktop
                                f.FileName = "GAF File.xml"
                                Dim othr As ThreadStart, myThread As Thread
                                othr = New ThreadStart(AddressOf f.GetFileName)
                                myThread = New Thread(othr)
                                myThread.SetApartmentState(ApartmentState.STA)
                                Try
                                    myThread.Start()
                                    While (myThread.IsAlive = False)
                                    End While
                                    Thread.Sleep(1)
                                    myThread.Join()
                                    If f.FileName <> "" Then
                                        Dim dt As DataTable
                                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            dt = Functions.Hana_RunQuery("Call sp_B1Addon_GAF ('" + frdate + "','" + todate + "')")
                                        Else
                                            dt = Functions.ADO_RunQuery("exec sp_B1Addon_GAF '" + frdate + "','" + todate + "'")
                                        End If

                                        If dt.Rows.Count > 0 Then
                                            Dim xmlstr As String = dt.Rows(0).Item("xmlGAF").ToString
                                            WriteFile(xmlstr, f.FileName)
                                            If xmlstr = "" Then
                                                SBO_Application.SetStatusBarMessage("There's no data!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Else
                                                SBO_Application.SetStatusBarMessage("Export completed!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage("ItemEvent GAF File XML: " + ex.ToString, , True)
                                End Try
                            ElseIf cb.Selected.Value = "3" Then
                                Dim f As New GetFileNameClass(1)
                                f.Filter = "Text files|*.txt"
                                f.InitialDirectory = Environment.SpecialFolder.Desktop
                                f.FileName = "GAF File.txt"
                                Dim othr As ThreadStart, myThread As Thread
                                othr = New ThreadStart(AddressOf f.GetFileName)
                                myThread = New Thread(othr)
                                myThread.SetApartmentState(ApartmentState.STA)
                                Try
                                    myThread.Start()
                                    While (myThread.IsAlive = False)
                                    End While
                                    Thread.Sleep(1)
                                    myThread.Join()
                                    If f.FileName <> "" Then
                                        Dim dt As DataTable
                                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            dt = Functions.Hana_RunQuery("Call sp_B1Addon_GAFtxt ('" + frdate + "','" + todate + "')")
                                        Else
                                            dt = Functions.ADO_RunQuery("exec sp_B1Addon_GAFtxt '" + frdate + "','" + todate + "'")
                                        End If

                                        If dt.Rows.Count > 0 Then
                                            Dim xmlstr As String = dt.Rows(0).Item("xmlGAF").ToString
                                            WriteFile(xmlstr, f.FileName)
                                            If xmlstr = "" Then
                                                SBO_Application.SetStatusBarMessage("There's no data!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Else
                                                SBO_Application.SetStatusBarMessage("Export completed!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage("ItemEvent GAF File TXT: " + ex.ToString, , True)
                                End Try
                            ElseIf cb.Selected.Value = "2" Then
                                Dim f As New GetFileNameClass(1)
                                f.Filter = "GST Tap Return File|*.txt"
                                f.InitialDirectory = Environment.SpecialFolder.Desktop
                                f.FileName = "GST Tap Return File.txt"
                                Dim othr As ThreadStart, myThread As Thread
                                othr = New ThreadStart(AddressOf f.GetFileName)
                                myThread = New Thread(othr)
                                myThread.SetApartmentState(ApartmentState.STA)
                                Try
                                    myThread.Start()
                                    While (myThread.IsAlive = False)
                                    End While
                                    Thread.Sleep(1)
                                    myThread.Join()
                                    If f.FileName <> "" Then
                                        Dim dt As DataTable
                                        If PublicVariable.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            dt = Functions.Hana_RunQuery("Call sp_B1Addon_GSTReturn ('" + frdate + "','" + todate + "')")
                                        Else
                                            dt = Functions.DoQueryReturnDT("exec sp_B1Addon_GSTReturn '" + frdate + "','" + todate + "'")
                                        End If

                                        If dt.Rows.Count > 0 Then
                                            Dim xmlstr As String = dt.Rows(0).Item("txtFilestr").ToString
                                            WriteFile(xmlstr, f.FileName)
                                            If xmlstr = "" Then
                                                SBO_Application.SetStatusBarMessage("There's no data!!", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                            Else
                                                SBO_Application.SetStatusBarMessage("Export completed!", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    SBO_Application.SetStatusBarMessage("ItemEvent Tap File: " + ex.ToString, , True)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage("ItemEvent Event: " + ex.ToString, , True)
        End Try
    End Sub
    Private Sub WriteFile(ByVal Str As String, ByVal FileName As String)
        Dim oWrite As IO.StreamWriter
        oWrite = IO.File.CreateText(FileName)
        oWrite.Write(Str)
        oWrite.Close()
    End Sub
End Class
