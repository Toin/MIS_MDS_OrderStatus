Module OrderStatus


    Sub OrderStatus_FormEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oSODelvMassUpdateGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord4")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams

            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord4"
            fcp.UniqueID = "mds_ord4"
            'fcp.ObjectType = "MIS_OPTIM"

            'fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = MenuCreation.LoadFromXML("sotomfg.srf")
            fcp.XmlData = LoadFromXML("OrderStatusForMKT.srf")

            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)


            '' Add User DataSource
            '' not binding to SBO data or UDO/UDF
            'oForm.DataSources.UserDataSources.Add("SODueFrom", SAPbouiCOM.BoDataType.dt_DATE)
            'oForm.DataSources.UserDataSources.Add("SODueTo", SAPbouiCOM.BoDataType.dt_DATE)

            'oForm.DataSources.UserDataSources.Add("SOFull", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oForm.DataSources.UserDataSources.Add("SOPartial", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            'oForm.DataSources.UserDataSources.Add("AsOfDate", SAPbouiCOM.BoDataType.dt_DATE)
            'oForm.DataSources.UserDataSources.Add("IntervlDay", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            'oForm.DataSources.UserDataSources.Add("NewDueDate", SAPbouiCOM.BoDataType.dt_DATE)
            'oForm.DataSources.UserDataSources.Add("NewReason", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            ''Default value for PdO Due Date From - To, As of Date, Interval Days
            'oForm.DataSources.UserDataSources.Item("AsOfDate").Value = DateTime.Today.ToString("yyyyMMdd")
            'oForm.DataSources.UserDataSources.Item("IntervlDay").Value = "14"
            'oForm.DataSources.UserDataSources.Item("SOFull").Value = "1"
            ''oForm.DataSources.UserDataSources.Item("SOPartial").Value = "0"

            'oForm.DataSources.UserDataSources.Item("SODueFrom").Value = DateTime.Today.ToString("yyyyMMdd")
            'oForm.DataSources.UserDataSources.Item("SODueTo").Value = DateTime.Today.AddDays(14).ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)


            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            oEditText = oForm.Items.Item("SoNumber").Specific

            'oForm.Items.Item("SODueFrom").Width = 100
            'oEditText = oForm.Items.Item("SODueFrom").Specific
            'oEditText.DataBind.SetBound(True, "", "SODueFrom")

            'oForm.Items.Item("SODueTo").Width = 100
            'oEditText = oForm.Items.Item("SODueTo").Specific
            'oEditText.DataBind.SetBound(True, "", "SODueTo")

            'oForm.Items.Item("AsOfDate").Width = 100
            'oEditText = oForm.Items.Item("AsOfDate").Specific
            'oEditText.DataBind.SetBound(True, "", "AsOfDate")

            'oForm.Items.Item("IntervlDay").Width = 50
            'oEditText = oForm.Items.Item("IntervlDay").Specific
            'oEditText.DataBind.SetBound(True, "", "IntervlDay")


            'Dim optButton As SAPbouiCOM.OptionBtn


            'oForm.Items.Item("optSOFull").Width = 160
            'oItem = oForm.Items.Item("optSOFull")
            'optButton = oItem.Specific
            'optButton.DataBind.SetBound(True, "", "SOFull")

            'oForm.Items.Item("optSOHalf").Width = 160
            'oItem = oForm.Items.Item("optSOHalf")
            'optButton = oItem.Specific
            ''optButton.DataBind.SetBound(True, "", "SOPartial")
            'optButton.DataBind.SetBound(True, "", "SOFull")
            'optButton.GroupWith("optSOFull")

            'oForm.Items.Item("NewDueDate").Width = 100
            'oEditText = oForm.Items.Item("NewDueDate").Specific
            'oEditText.DataBind.SetBound(True, "", "NewDueDate")

            ''oForm.Items.Item("NewReason").Width = 50
            'oEditText = oForm.Items.Item("NewReason").Specific
            'oEditText.DataBind.SetBound(True, "", "NewReason")

            oItem = oForm.Items.Item("SODelvGrid")
            oItem.Left = 5
            oItem.Top = 70
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oSODelvMassUpdateGrid = oItem.Specific

            oSODelvMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            'Dim SODuefrom As String
            'Dim SODueto As String
            'Dim intervalddays As Integer
            'Dim asofdate As String

            'intervalddays = IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string)
            'asofdate = Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd")

            'SODuefrom = Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd")
            'SODueto = Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd")

            'STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
            '    IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
            '    + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            '    + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
            '    + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
            '    + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)

            STMTQuery = "EXEC Get_OrderStatus " _
                + " '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
                + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("SODelvList")
            'oForm.DataSources.DataTables.Item("SODelvList").ExecuteQuery(STMTQuery)
            oSODelvMassUpdateGrid.DataTable = oForm.DataSources.DataTables.Item("SODelvList")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oSODelvMassUpdateGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        ''oForm.Top = 150
        ''oForm.Left = 330
        ''oForm.Width = 900


        ''STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
        'STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
        '    IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
        '    + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
        '    + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
        '    + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
        '    + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)

        STMTQuery = "EXEC Get_OrderStatus " _
            + " '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        oForm.Items.Item("BPCardCode").Click()


        oForm.Top = 250
        oForm.Left = 100
        oForm.Width = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Width - 100
        oForm.Height = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Height - 200


        RearrangeOrderStatusGrid(oForm)


        oForm.Visible = True

    End Sub


    Sub RearrangeOrderStatusGrid(ByVal oForm As SAPbouiCOM.Form)


        Dim oColumn As SAPbouiCOM.EditTextColumn


        Dim oSODelvMassUpdateGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)


        oSODelvMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific


        'oSODelvMassUpdateGrid.RowHeaders.Width = 50


        ' ''Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        ''oColumn = oSODelvMassUpdateGrid.Columns.Item("Cust Code")
        ''oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        ''oColumn.Editable = False

        'oColumn = oSODelvMassUpdateGrid.Columns.Item("DocEntry")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        'oColumn.Editable = False


        'oColumn = oSODelvMassUpdateGrid.Columns.Item("ItemCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        'oColumn.Editable = False
        'oColumn.Width = 100

        'oSODelvMassUpdateGrid.Columns.Item("DocEntry").Visible = False

        'oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").Width = 50
        'oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        'oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").TitleObject.Sortable = True

        'oSODelvMassUpdateGrid.Columns.Item("Cust Code").Width = 70
        oSODelvMassUpdateGrid.Columns.Item("Customer").Width = 130
        oSODelvMassUpdateGrid.Columns.Item("SO#").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("SO Date").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("ETD (SO)").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("Due Date Overdue").Width = 20
        'oSODelvMassUpdateGrid.Columns.Item("Over").ForeColor = ColorTranslator.ToOle(Color.Red)

        'oSODelvMassUpdateGrid.Columns.Item("SOOverDue").Width = 35
        oSODelvMassUpdateGrid.Columns.Item("Line").Width = 35
        oSODelvMassUpdateGrid.Columns.Item("Line").RightJustified = True

        oSODelvMassUpdateGrid.Columns.Item("Item Code").Width = 80
        oSODelvMassUpdateGrid.Columns.Item("Item Name").Width = 80
        oSODelvMassUpdateGrid.Columns.Item("Project").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("P (cm)").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("L (cm)").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("P x L (m2)").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("SO Qty").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("SO Qty").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Terkirim").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("Terkirim").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Open Qty").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("Open Qty").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").Width = 60
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Production Status").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("Delivery Status").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("PdO No.").Width = 75
        'oSODelvMassUpdateGrid.Columns.Item("Processing Finish (Plan, Last)").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("Receipt Date").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("Receipt Created").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("Receipt Time PdO").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("Due Date PdO").Width = 90
        oSODelvMassUpdateGrid.Columns.Item("PPIC Remarks (last)").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("Plan Delivery Date, Last").Width = 100
        oSODelvMassUpdateGrid.Columns.Item("DO Date").Width = 100
        oSODelvMassUpdateGrid.Columns.Item("DO#").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("Delivery Remarks (last)").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("Delivery Method").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("Sales Rep").Width = 120



        oSODelvMassUpdateGrid.Columns.Item("Customer").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Customer").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO#").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO#").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO Date").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO Date").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("ETD (SO)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("ETD (SO)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Due Date Overdue").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Due Date Overdue").TitleObject.Sortable = True

        oSODelvMassUpdateGrid.Columns.Item("Line").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Line").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Item Code").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Item Code").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Item Name").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Item Name").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Project").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Project").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("P (cm)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("P (cm)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("L (cm)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("L (cm)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("P x L (m2)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("P x L (m2)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO Qty").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO Qty").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Terkirim").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Terkirim").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Open Qty").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Open Qty").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Production Status").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Production Status").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Delivery Status").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Delivery Status").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("PdO No.").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("PdO No.").TitleObject.Sortable = True
        'oSODelvMassUpdateGrid.Columns.Item("Processing Finish (Plan, Last)").Editable = False
        'oSODelvMassUpdateGrid.Columns.Item("Processing Finish (Plan, Last)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Receipt Date").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Receipt Date").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Receipt Created").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Receipt Created").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Receipt Time PdO").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Receipt Time PdO").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Due Date PdO").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Due Date PdO").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("PPIC Remarks (last)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("PPIC Remarks (last)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Plan Delivery Date, Last").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Plan Delivery Date, Last").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("DO Date").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("DO Date").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("DO#").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("DO#").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Delivery Remarks (last)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Delivery Remarks (last)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Delivery Method").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Delivery Method").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Sales Rep").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Sales Rep").TitleObject.Sortable = True


        'oColumn = oSODelvMassUpdateGrid.Columns.Item("Customer Name")
        'oColumn.Editable = False

        'oSODelvMassUpdateGrid.Columns.Item("SO Date").Width = 80
        'oSODelvMassUpdateGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oSODelvMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oSODelvMassUpdateGrid.Columns.Item(0).TitleObject.Caption = oSODelvMassUpdateGrid.Rows.Count.ToString


        'If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        'And oSODelvMassUpdateGrid.DataTable.GetValue(0, 0) <> "" Then
        '    oForm.Items.Item("cmdUpdSO").Enabled = True
        'Else
        '    oForm.Items.Item("cmdUpdSO").Enabled = False
        'End If


        'oSODelvMassUpdateGrid.Columns.Item("#").Editable = False

        'In case to enable button 'Filter table...' then all column must be readonly editable = false

        'oSODelvMassUpdateGrid.Columns.Item(10).Width = 70
        'oSODelvMassUpdateGrid.Columns.Item(11).Width = 70
        'oSODelvMassUpdateGrid.Columns.Item(12).Width = 70
        'oSODelvMassUpdateGrid.Columns.Item(13).Width = 70
        'oSODelvMassUpdateGrid.Columns.Item(14).Width = 70

        '' How-To Change Column Title/Header Caption in Grid
        'oSODelvMassUpdateGrid.Columns.Item(10).Width = 60
        'oSODelvMassUpdateGrid.Columns.Item(10).Description = "08"
        'oSODelvMassUpdateGrid.Columns.Item(10).TitleObject.Caption = _
        '    Mid(oSODelvMassUpdateGrid.Columns.Item(10).TitleObject.Caption, 6, 2) + "/" + _
        '    Right(oSODelvMassUpdateGrid.Columns.Item(10).TitleObject.Caption, 2)


        ' Interval days + 2 extra days: one for < first column and one for > last column
        ' e.g: 14days, as of date 2012-02-08, [2012-02-07], [2012-02-08], [2012-02-09], ... [2012-02-22]

        Dim colcount As Integer = oSODelvMassUpdateGrid.Columns.Count

        'Dim oItem As SAPbouiCOM.Item
        'Dim oEditText As SAPbouiCOM.EditText

        For idx = 0 To colcount - 1 'CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
            If Left(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 1) = Left(Now.Year.ToString, 1) Then
                oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = _
                    Mid(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                    Right(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 2)
                oSODelvMassUpdateGrid.Columns.Item(idx).Width = 40
                oSODelvMassUpdateGrid.Columns.Item(idx).Editable = False
                oSODelvMassUpdateGrid.Columns.Item(idx).RightJustified = True
                'oSODelvMassUpdateGrid.Columns.Item(idx).TextStyle = 64
                'oItem = oSODelvMassUpdateGrid.Columns.Item(idx)
                'oColumn = oSODelvMassUpdateGrid.Columns.Item(idx)

                'oEditText = oSODelvMassUpdateGrid.Columns.Item(idx)
                'oEditText = oColumn
                'oEditText.SuppressZeros = True

            End If
        Next

        'For idx = 0 To CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
        '    oSODelvMassUpdateGrid.Columns.Item(10 + idx).Width = 60
        '    oSODelvMassUpdateGrid.Columns.Item(10 + idx).Editable = False
        'Next






        oForm.Items.Item("SODelvGrid").Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        oForm.Items.Item("SODelvGrid").Width = oForm.ClientWidth - 20


        'oForm.Items.Item("NewDue").Left = oForm.ClientWidth - 320
        'oForm.Items.Item("ReasonChg").Left = oForm.ClientWidth - 320
        'oForm.Items.Item("cmdUpdSO").Left = oForm.ClientWidth - 320

        'oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth - 210
        'oForm.Items.Item("NewReason").Left = oForm.ClientWidth - 210

        'oForm.Items.Item("AsOf").Left = oForm.ClientWidth / 3
        'oForm.Items.Item("ViewPeriod").Left = oForm.ClientWidth / 3

        'oForm.Items.Item("optSOFull").Left = oForm.ClientWidth / 3
        'oForm.Items.Item("optSOHalf").Left = oForm.ClientWidth / 3

        'oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth / 3 + 120
        'oForm.Items.Item("IntervlDay").Left = oForm.ClientWidth / 3 + 120
        'oForm.Items.Item("Days").Left = oForm.ClientWidth / 3 + 180


        'Dim sboDate As String
        'Dim dDate As DateTime


        'sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)


        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oSODelvMassUpdateGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Function ValidateInputDueDate_Form_OrderStatus(ByVal oForm As SAPbouiCOM.Form)

        If oForm.Items.Item("NewDueDate").Specific.string = "" Then
            oApp.SetStatusBarMessage("New Due Date PdO must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If

        If oForm.Items.Item("NewReason").Specific.string = "" Then
            oApp.SetStatusBarMessage("New Reason why Change Due Date PdO must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If

        Return True

    End Function

    Sub LoadOrderStatus(ByVal oForm As SAPbouiCOM.Form)
        Dim STMTQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            'oApp.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            'oApp.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        'If oForm.Items.Item("SODueTo").Specific.string = "" Then
        '    oForm.Items.Item("SODueTo").Specific.string = oForm.Items.Item("SODueFrom").Specific.string
        'End If



        ''STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer " + _
        'STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
        '    IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
        '    + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
        '    + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
        '    + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
        '    + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)

        STMTQuery = "EXEC Get_OrderStatus " _
            + " '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string)

        'STMTQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, "

        '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T1.LineStatus = 'O' " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        'RearrangeOrderStatusForMKTGrid(oForm)
        RearrangeOrderStatusGrid(oForm)

    End Sub


End Module
