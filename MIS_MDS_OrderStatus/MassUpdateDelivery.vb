Module MassUpdateDelivery

    Sub MassUpdateDelivery_FormEntry()
        Dim oForm As SAPbouiCOM.Form
        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombobox As SAPbouiCOM.ComboBox
        Dim oButton As SAPbouiCOM.Button
        Dim oDOGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord3")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord3"
            fcp.UniqueID = "mds_ord3"

            fcp.XmlData = LoadFromXML("MassUpdateDelivery.srf")
            oForm = oApp.Forms.AddEx(fcp)
            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("DODateFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("DODateTo", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("DelDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("Jam", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("cmbMethod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

            'Default value for Date From - To
            oForm.DataSources.UserDataSources.Item("DODateFrom").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("DODateTo").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("DelDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("cmbMethod").Value = 1
            oForm.DataSources.UserDataSources.Item("Jam").Value = DateTime.Today.Now.ToString("HH:mm")

            oCombobox = oForm.Items.Item("cmbMethod").Specific
            oCombobox.ValidValues.Add("", " All")
            oCombobox.ValidValues.Add("1", "Picked Up by Customer")
            oCombobox.ValidValues.Add("2", "Delivered by Maruni")
            oCombobox.DataBind.SetBound(True, "", "cmbMethod")

            oForm.Items.Item("DODateFrom").Width = 100
            oEditText = oForm.Items.Item("DODateFrom").Specific
            oEditText.DataBind.SetBound(True, "", "DODateFrom")

            oForm.Items.Item("DODateTo").Width = 100
            oEditText = oForm.Items.Item("DODateTo").Specific
            oEditText.DataBind.SetBound(True, "", "DODateTo")

            oForm.Items.Item("DelDate").Width = 100
            oEditText = oForm.Items.Item("DelDate").Specific
            oEditText.DataBind.SetBound(True, "", "DelDate")

            oEditText = oForm.Items.Item("Jam").Specific
            oEditText.DataBind.SetBound(True, "", "Jam")
            'oEditText.Value.Format(SAPbobsCOM.BoFldSubTypes.st_Time, CShort(oForm.Items.Item("Jam").Specific))

            oItem = oForm.Items.Item("DOGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            oDOGrid = oItem.Specific
            oDOGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            Dim DODateFrom As String
            Dim DODateTo As String
            Dim PickedUp As String
            DODateFrom = Format(CDate(oForm.Items.Item("DODateFrom").Specific.string), "yyyyMMdd")
            DODateTo = Format(CDate(oForm.Items.Item("DODateTo").Specific.string), "yyyyMMdd")
            PickedUp = oForm.Items.Item("cmbMethod").Specific.selected.value()

            STMTQuery = "EXEC MIS_MassUpdateDelivery '" + DODateFrom + "', '" + DODateTo + "', '" + PickedUp + "'"

            ' Grid #: 1
            oForm.DataSources.DataTables.Add("DOList")
            oForm.DataSources.DataTables.Item("DOList").ExecuteQuery(STMTQuery)
            oDOGrid.DataTable = oForm.DataSources.DataTables.Item("DOList")

            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oCombobox = Nothing
            oDOGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        RearrangeDOGrid(oForm)
        oForm.Visible = True

    End Sub

    Sub RearrangeDOGrid(ByVal oForm As SAPbouiCOM.Form)
        Dim oColumn As SAPbouiCOM.EditTextColumn
        Dim oDOGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)
        oDOGrid = oForm.Items.Item("DOGrid").Specific

        oDOGrid.Columns.Item("Check").Width = 50
        oDOGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oDOGrid.Columns.Item("Check").TitleObject.Sortable = True

        oDOGrid.Columns.Item("DO").Width = 50
        oColumn = oDOGrid.Columns.Item("DO")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
        oColumn.Editable = False

        oDOGrid.Columns.Item("DO#").Width = 75

        If oForm.DataSources.UserDataSources.Item("Jam").Value = "" Then
            oForm.DataSources.UserDataSources.Item("Jam").Value = DateTime.Today.Now.ToString("HH:mm")
        End If

        oDOGrid.Columns.Item("DO Date").Width = 75
        oDOGrid.Columns.Item("DO Line").Width = 50
        oDOGrid.Columns.Item("Item Code").Width = 90
        oDOGrid.Columns.Item("Item Name").Width = 100
        oDOGrid.Columns.Item("UoM").Width = 50
        oDOGrid.Columns.Item("DO Qty").Width = 80
        oDOGrid.Columns.Item("Pickup/Delivery").Width = 80
        oDOGrid.Columns.Item("Delivery Address").Width = 150
        oDOGrid.Columns.Item("Customer").Width = 100
        oDOGrid.Columns.Item("Sales Rep").Width = 100

        oDOGrid.Columns.Item("DO").Editable = False
        oDOGrid.Columns.Item("DO").TitleObject.Sortable = True
        oDOGrid.Columns.Item("DO#").Editable = False
        oDOGrid.Columns.Item("DO#").TitleObject.Sortable = True
        oDOGrid.Columns.Item("DO Date").Editable = False
        oDOGrid.Columns.Item("DO Date").TitleObject.Sortable = True
        oDOGrid.Columns.Item("DO Line").Editable = False
        oDOGrid.Columns.Item("DO Line").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Item Code").Editable = False
        oDOGrid.Columns.Item("Item Code").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Item Name").Editable = False
        oDOGrid.Columns.Item("Item Name").TitleObject.Sortable = True
        oDOGrid.Columns.Item("UoM").Editable = False
        oDOGrid.Columns.Item("UoM").TitleObject.Sortable = True
        oDOGrid.Columns.Item("DO Qty").Editable = False
        oDOGrid.Columns.Item("DO Qty").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Pickup/Delivery").Editable = False
        oDOGrid.Columns.Item("Pickup/Delivery").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Delivery Address").Editable = False
        oDOGrid.Columns.Item("Delivery Address").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Customer").Editable = False
        oDOGrid.Columns.Item("Customer").TitleObject.Sortable = True
        oDOGrid.Columns.Item("Sales Rep").Editable = False
        oDOGrid.Columns.Item("Sales Rep").TitleObject.Sortable = True

        Dim colcount As Integer = oDOGrid.Columns.Count

        For idx = 0 To colcount - 1 'CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
            If Left(oDOGrid.Columns.Item(idx).TitleObject.Caption, 1) = Left(Now.Year.ToString, 1) Then
                oDOGrid.Columns.Item(idx).TitleObject.Caption = _
                    Mid(oDOGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                    Right(oDOGrid.Columns.Item(idx).TitleObject.Caption, 2)
                oDOGrid.Columns.Item(idx).Width = 40
                oDOGrid.Columns.Item(idx).Editable = False
                oDOGrid.Columns.Item(idx).RightJustified = True
                oDOGrid.Columns.Item(idx).TitleObject.Sortable = True
            End If
        Next

        oForm.Items.Item("DOGrid").Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        oForm.Items.Item("DOGrid").Width = oForm.ClientWidth - 20

        Dim sboDate As String
        Dim dDate As DateTime

        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oDOGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Function ValidateInputDODate_Form_DODateMassUpdation(ByVal oForm As SAPbouiCOM.Form)

        If oForm.Items.Item("DelDate").Specific.string = "" Then
            oApp.SetStatusBarMessage("Delivery Date DO must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If

        If oForm.Items.Item("Jam").Specific.string = "" Then
            oApp.SetStatusBarMessage("Time Date DO must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Return True
    End Function


    Sub LoadDO(ByVal oForm As SAPbouiCOM.Form)
        Dim STMTQuery As String
        Dim DODateFrom As String
        Dim DODateTo As String
        Dim PickedUp As String

        If oForm.Items.Item("DODateTo").Specific.string = "" Then
            oForm.Items.Item("DODateTo").Specific.string = oForm.Items.Item("DODateFrom").Specific.string
        End If

        DODateFrom = Format(CDate(oForm.Items.Item("DODateFrom").Specific.string), "yyyyMMdd")
        DODateTo = Format(CDate(oForm.Items.Item("DODateTo").Specific.string), "yyyyMMdd")
        PickedUp = oForm.Items.Item("cmbMethod").Specific.selected.value()

        STMTQuery = "EXEC MIS_MassUpdateDelivery '" + DODateFrom + "', '" + DODateTo + "', '" + PickedUp + "'"

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        RearrangeDOGrid(oForm)

    End Sub

    Sub GenerateDODateUpdate(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler
        Dim oDOGrid As SAPbouiCOM.Grid
        Dim idx As Long

        oDOGrid = oForm.Items.Item("DOGrid").Specific
        oDOGrid.Columns.Item("Check").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        'Loop only selected/checked in grid rows and exit.
        For idx = oDOGrid.Rows.Count - 1 To 0 Step -1

            oApp.SetStatusBarMessage("Revising DO Date .... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            If UCase(oDOGrid.DataTable.GetValue(0, oDOGrid.GetDataTableRowIndex(idx))) = "Y" Then
                Dim oDelivery As SAPbobsCOM.Documents = Nothing
                Dim errConnect As String = ""

                If Not oCompany.InTransaction Then
                    oCompany.StartTransaction()
                End If

                'Dim PdOno As String = ""

                oDelivery = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                oDelivery.GetByKey(oDOGrid.DataTable.GetValue(1, oDOGrid.GetDataTableRowIndex(idx).ToString))
                oDelivery.UserFields.Fields.Item("U_MIS_ActDelivDt").Value = CDate(oForm.Items.Item("DelDate").Specific.string)

                'Dim time As String
                'Dim formats As String() = New String() {"HHmm"}
                'Dim dt As DateTime
                'Dim strTime As String
                'time = oForm.Items.Item("Jam").Specific.string

                'dt = DateTime.ParseExact(time, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AdjustToUniversal)

                'strTime = dt.ToString("HH:mm")
                'strTime = oForm.Items.Item("Jam").Specific.string.ToString("HHmm")

                'strTime = Left(oForm.Items.Item("Jam").Specific.string, 2) + Right(oForm.Items.Item("Jam").Specific.string, 2)
                'strTime = oForm.Items.Item("Jam").Specific.string

                'oDelivery.UserFields.Fields.Item("U_MIS_ActDelivTm").Value = strTime

                oDelivery.UserFields.Fields.Item("U_MIS_ActDelivTm").Value = oForm.Items.Item("Jam").Specific.string

                lRetCode = oDelivery.Update()

                If lRetCode <> 0 Then
                    'vCompany.GetLastError(lErrCode, sErrMsg)
                    'SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    oApp.MessageBox(lErrCode & ": " & sErrMsg)

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                Else

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDelivery)
                    oDelivery = Nothing

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If

                End If

                GC.Collect()
            Else
                Exit For
            End If
        Next

        'MsgBox("Begin trx: generating... PdO")
        'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        oApp.SetStatusBarMessage("Revising DO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        'MsgBox("generating... PdO; DONE!!!")

        Exit Sub

errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

End Module
