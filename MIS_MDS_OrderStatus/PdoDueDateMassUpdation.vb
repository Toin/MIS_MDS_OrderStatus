Module PdoDueDateMassUpdation

    Sub PdoDueDateMassUpdation_FormEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord1")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams

            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord1"
            fcp.UniqueID = "mds_ord1"
            'fcp.ObjectType = "MIS_OPTIM"

            'fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = MenuCreation.LoadFromXML("sotomfg.srf")
            fcp.XmlData = LoadFromXML("PdODueDateMassUpdation.srf")

            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)


            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("PdODueFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("PdODueTo", SAPbouiCOM.BoDataType.dt_DATE)

            oForm.DataSources.UserDataSources.Add("AsOfDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("IntervlDay", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            oForm.DataSources.UserDataSources.Add("NewDueDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("NewReason", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            'Default value for PdO Due Date From - To, As of Date, Interval Days
            oForm.DataSources.UserDataSources.Item("AsOfDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("IntervlDay").Value = "14"

            oForm.DataSources.UserDataSources.Item("PdODueFrom").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("PdODueTo").Value = DateTime.Today.AddDays(14).ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)


            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            oEditText = oForm.Items.Item("SoNumber").Specific

            oForm.Items.Item("PdODueFrom").Width = 100
            oEditText = oForm.Items.Item("PdODueFrom").Specific
            oEditText.DataBind.SetBound(True, "", "PdODueFrom")

            oForm.Items.Item("PdODueTo").Width = 100
            oEditText = oForm.Items.Item("PdODueTo").Specific
            oEditText.DataBind.SetBound(True, "", "PdODueTo")

            oForm.Items.Item("AsOfDate").Width = 100
            oEditText = oForm.Items.Item("AsOfDate").Specific
            oEditText.DataBind.SetBound(True, "", "AsOfDate")

            oForm.Items.Item("IntervlDay").Width = 50
            oEditText = oForm.Items.Item("IntervlDay").Specific
            oEditText.DataBind.SetBound(True, "", "IntervlDay")

            oForm.Items.Item("NewDueDate").Width = 100
            oEditText = oForm.Items.Item("NewDueDate").Specific
            oEditText.DataBind.SetBound(True, "", "NewDueDate")

            'oForm.Items.Item("NewReason").Width = 50
            oEditText = oForm.Items.Item("NewReason").Specific
            oEditText.DataBind.SetBound(True, "", "NewReason")

            oItem = oForm.Items.Item("PdODueGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oPdODueMassUpdateGrid = oItem.Specific

            oPdODueMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            Dim pdoduefrom As String
            Dim pdodueto As String
            Dim intervalddays As Integer
            Dim asofdate As String

            intervalddays = IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string)
            asofdate = Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd")

            pdoduefrom = Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd")
            pdodueto = Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd")

            STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer " + _
                IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
                + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
                + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
                + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
                + "' "

            'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
            '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("PdODueList")
            oForm.DataSources.DataTables.Item("PdODueList").ExecuteQuery(STMTQuery)
            oPdODueMassUpdateGrid.DataTable = oForm.DataSources.DataTables.Item("PdODueList")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oPdODueMassUpdateGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        ''oForm.Top = 150
        ''oForm.Left = 330
        ''oForm.Width = 900


        'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
        STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "' "


        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        oForm.Items.Item("BPCardCode").Click()


        oForm.Top = 250
        oForm.Left = 100
        oForm.Width = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Width - 100
        oForm.Height = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Height - 200


        RearrangePdoDueGrid(oForm)


        oForm.Visible = True

    End Sub

    Sub PdoDueDateMassUpdationByMachine_FormEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oPdODueMassUpdateByMachineGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord3")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams

            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord3"
            fcp.UniqueID = "mds_ord3"
            'fcp.ObjectType = "MIS_OPTIM"

            'fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = MenuCreation.LoadFromXML("sotomfg.srf")
            'fcp.XmlData = LoadFromXML("PdODueDateMassUpdationbyMachine.srf")
            'fcp.XmlData = LoadFromXML("PdODueDatebyMachine.srf")
            fcp.XmlData = LoadFromXML("PdoDueDateMassUpdation.srf")

            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)


            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("PdODueFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("PdODueTo", SAPbouiCOM.BoDataType.dt_DATE)

            oForm.DataSources.UserDataSources.Add("AsOfDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("IntervlMch", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            oForm.DataSources.UserDataSources.Add("NewDueDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("NewReason", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            'Default value for PdO Due Date From - To, As of Date, Interval Days
            oForm.DataSources.UserDataSources.Item("AsOfDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("IntervlMch").Value = "14"

            oForm.DataSources.UserDataSources.Item("PdODueFrom").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("PdODueTo").Value = DateTime.Today.AddDays(14).ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)


            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            oEditText = oForm.Items.Item("SoNumber").Specific

            oForm.Items.Item("PdODueFrom").Width = 100
            oEditText = oForm.Items.Item("PdODueFrom").Specific
            oEditText.DataBind.SetBound(True, "", "PdODueFrom")

            oForm.Items.Item("PdODueTo").Width = 100
            oEditText = oForm.Items.Item("PdODueTo").Specific
            oEditText.DataBind.SetBound(True, "", "PdODueTo")

            oForm.Items.Item("AsOfDate").Width = 100
            oEditText = oForm.Items.Item("AsOfDate").Specific
            oEditText.DataBind.SetBound(True, "", "AsOfDate")

            oForm.Items.Item("IntervlMch").Width = 50
            oEditText = oForm.Items.Item("IntervlMch").Specific
            oEditText.DataBind.SetBound(True, "", "IntervlMch")

            oForm.Items.Item("NewDueDate").Width = 100
            oEditText = oForm.Items.Item("NewDueDate").Specific
            oEditText.DataBind.SetBound(True, "", "NewDueDate")

            'oForm.Items.Item("NewReason").Width = 50
            oEditText = oForm.Items.Item("NewReason").Specific
            oEditText.DataBind.SetBound(True, "", "NewReason")

            oItem = oForm.Items.Item("PdOMchGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oPdODueMassUpdateByMachineGrid = oItem.Specific

            oPdODueMassUpdateByMachineGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            Dim pdoduefrom As String
            Dim pdodueto As String
            Dim intervalddays As Integer
            Dim asofdate As String

            intervalddays = IIf(oForm.Items.Item("IntervlMch").Specific.string = "", 0, oForm.Items.Item("IntervlMch").Specific.string)
            asofdate = Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd")

            pdoduefrom = Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd")
            pdodueto = Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd")

            STMTQuery = "EXEC GetPdODueDateListByMachine_ByRangeofPdoDueDateSONumCustomer " + _
                IIf(oForm.Items.Item("IntervlMch").Specific.string = "", 0, oForm.Items.Item("IntervlMch").Specific.string) _
                + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
                + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
                + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
                + "' "

            'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
            '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("PdODueListByMachine")
            oForm.DataSources.DataTables.Item("PdODueListByMachine").ExecuteQuery(STMTQuery)
            oPdODueMassUpdateByMachineGrid.DataTable = oForm.DataSources.DataTables.Item("PdODueListByMachine")


            'oForm = Nothing
            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oPdODueMassUpdateByMachineGrid = Nothing

            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try

        ''oForm.Top = 150
        ''oForm.Left = 330
        ''oForm.Width = 900


        'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
        STMTQuery = "EXEC GetPdODueDateListByMachine_ByRangeofPdoDueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlMch").Specific.string = "", 0, oForm.Items.Item("IntervlMch").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "' "


        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        oForm.Items.Item("BPCardCode").Click()


        oForm.Top = 250
        oForm.Left = 100
        oForm.Width = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Width - 100
        oForm.Height = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Height - 200


        RearrangePdoDueGridByMachine(oForm)


        oForm.Visible = True

    End Sub

    Sub RearrangePdoDueGrid(ByVal oForm As SAPbouiCOM.Form)


        Dim oColumn As SAPbouiCOM.EditTextColumn


        Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)


        oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific


        'oPdODueMassUpdateGrid.RowHeaders.Width = 50

        ''Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        'oColumn = oPdODueMassUpdateGrid.Columns.Item("CustCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        'oColumn.Editable = False

        'oColumn = oPdODueMassUpdateGrid.Columns.Item("SO#")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        'oColumn.Editable = False


        'Dim OCOL As SAPbouiCOM.Column
        'Dim ocols As SAPbouiCOM.Columns
        'ocols = oPdODueMassUpdateGrid.Columns

        'OCOL = ocols.Item("SO#")
        'OCOL.DataBind.SetBound(True, "", "SO#")
        'OCOL.ExtendedObject.linkedobject = SAPbouiCOM.BoLinkedObject.lf_Order

        'OCOL = oPdODueMassUpdateGrid.Columns.Item("SO#")
        'OCOL.DataBind.SetBound(True, "", "SO#")
        'OCOL.ExtendedObject.linkedobject = SAPbouiCOM.BoLinkedObject.lf_Order


        'oColumn = oPdODueMassUpdateGrid.Columns.Item("ItemCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        'oColumn.Editable = False
        'oColumn.Width = 100

        oColumn = oPdODueMassUpdateGrid.Columns.Item("PdOKey")
        oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_ProductionOrder '"2"
        oColumn.Editable = False
        oColumn.Width = 50
        'oPdODueMassUpdateGrid.Columns.Item("PdOKey").Visible = False

        oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").Width = 50
        oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").TitleObject.Sortable = True

        oPdODueMassUpdateGrid.Columns.Item("PdO#").Width = 60
        oPdODueMassUpdateGrid.Columns.Item("SO#").Width = 75
        oPdODueMassUpdateGrid.Columns.Item("SODueDate").Width = 75
        oPdODueMassUpdateGrid.Columns.Item("PdOOverDue").Width = 20
        oPdODueMassUpdateGrid.Columns.Item("PdOOverDue").ForeColor = ColorTranslator.ToOle(Color.Red)
        oPdODueMassUpdateGrid.Columns.Item("PdODue").Width = 75
        oPdODueMassUpdateGrid.Columns.Item("PdODue").Width = 75
        oPdODueMassUpdateGrid.Columns.Item("PdO Qty").Width = 50
        oPdODueMassUpdateGrid.Columns.Item("PdO Qty").RightJustified = True

        oPdODueMassUpdateGrid.Columns.Item("CustCode").Width = 70
        oPdODueMassUpdateGrid.Columns.Item("P (cm)").Width = 50
        oPdODueMassUpdateGrid.Columns.Item("L (cm)").Width = 50
        oPdODueMassUpdateGrid.Columns.Item("PxL (m2)").Width = 50
        oPdODueMassUpdateGrid.Columns.Item("UoM").Width = 30
        oPdODueMassUpdateGrid.Columns.Item("SOLineNum").Width = 30
        oPdODueMassUpdateGrid.Columns.Item("SOOverDue").Width = 20
        oPdODueMassUpdateGrid.Columns.Item("SOOverDue").ForeColor = ColorTranslator.ToOle(Color.Red) ' Convert.ToInt32(System.Drawing.Color.Red.ToArgb)
        '16711680 ' System.Drawing.Color.Red.ToArgb
        'set red
        'System.Drawing.ColorTranslator.FromOle(&HFF0000).ToArgb() 
        'System.Drawing.ColorTranslator.FromHtml("FF0000").ToArgb()

        oPdODueMassUpdateGrid.Columns.Item("ItemName").Width = 80

        oPdODueMassUpdateGrid.Columns.Item("PdO#").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PdO#").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("SO#").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("SO#").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("SODueDate").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("SODueDate").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("PdODocDate").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PdODocDate").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("PdODue").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PdODue").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("PdOOverDue").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PdOOverDue").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("PdO Qty").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PdO Qty").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("CustCode").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("CustCode").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("P (cm)").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("P (cm)").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("L (cm)").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("L (cm)").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("PxL (m2)").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("PxL (m2)").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("UoM").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("UoM").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("SOLineNum").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("SOLineNum").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("SOOverDue").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("SOOverDue").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("ItemCode").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("ItemCode").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("ItemName").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("ItemName").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("Reason Change").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("Reason Change").TitleObject.Sortable = True
        oPdODueMassUpdateGrid.Columns.Item("Deliv Remarks").Editable = False
        oPdODueMassUpdateGrid.Columns.Item("Deliv Remarks").TitleObject.Sortable = True


        'oColumn = oPdODueMassUpdateGrid.Columns.Item("Customer Name")
        'oColumn.Editable = False

        'oPdODueMassUpdateGrid.Columns.Item("SO Date").Width = 80
        'oPdODueMassUpdateGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oPdODueMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = oPdODueMassUpdateGrid.Rows.Count.ToString


        If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        And oPdODueMassUpdateGrid.DataTable.GetValue(0, 0) <> "" Then
            oForm.Items.Item("cmdUpdPdO").Enabled = True
        Else
            oForm.Items.Item("cmdUpdPdO").Enabled = False
        End If


        'oPdODueMassUpdateGrid.Columns.Item("#").Editable = False

        'In case to enable button 'Filter table...' then all column must be readonly editable = false

        'oPdODueMassUpdateGrid.Columns.Item(10).Width = 70
        'oPdODueMassUpdateGrid.Columns.Item(11).Width = 70
        'oPdODueMassUpdateGrid.Columns.Item(12).Width = 70
        'oPdODueMassUpdateGrid.Columns.Item(13).Width = 70
        'oPdODueMassUpdateGrid.Columns.Item(14).Width = 70

        '' How-To Change Column Title/Header Caption in Grid
        'oPdODueMassUpdateGrid.Columns.Item(10).Width = 60
        'oPdODueMassUpdateGrid.Columns.Item(10).Description = "08"
        'oPdODueMassUpdateGrid.Columns.Item(10).TitleObject.Caption = _
        '    Mid(oPdODueMassUpdateGrid.Columns.Item(10).TitleObject.Caption, 6, 2) + "/" + _
        '    Right(oPdODueMassUpdateGrid.Columns.Item(10).TitleObject.Caption, 2)


        ' Interval days + 2 extra days: one for < first column and one for > last column
        ' e.g: 14days, as of date 2012-02-08, [2012-02-07], [2012-02-08], [2012-02-09], ... [2012-02-22]

        Dim colcount As Integer = oPdODueMassUpdateGrid.Columns.Count

        'Dim oItem As SAPbouiCOM.Item
        'Dim oEditText As SAPbouiCOM.EditText

        Dim IntervalIndex As Integer = 0

        'JIKA INTERVAL DAYS=14, DAN RANGE TGL: 2011-08-01 S/D 2011-08-18
        'RESULT ADALAH:
        'BENTUK KOLOM DGN JUDUL MULAI DARI 0, 1,2,...14, 14+1 
        'MISALKAN: 2011-07-31, 2011-08-01, 2011-08-02, ... 2011-08-15
        'HASILNYA: [< 08/01], [08/01], [08/02], [08/03], ... [08/14], [> 08/14]


        For idx = 0 To colcount - 1 'CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
            If Left(oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 1) = Left(Now.Year.ToString, 1) Then
                IntervalIndex += 1
                If IntervalIndex = 1 Then
                    oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = "<" + _
                        Mid(oPdODueMassUpdateGrid.Columns.Item(idx + 1).TitleObject.Caption, 6, 2) + "/" + _
                        Right(oPdODueMassUpdateGrid.Columns.Item(idx + 1).TitleObject.Caption, 2)
                    oPdODueMassUpdateGrid.Columns.Item(idx).Width = 50
                ElseIf IntervalIndex = _
                        CInt(IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, _
                                 oForm.Items.Item("IntervlDay").Specific.string)) + 2 Then
                    oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = ">" + _
                        Mid(oPdODueMassUpdateGrid.Columns.Item(idx - 1).TitleObject.Caption, 1, 2) + "/" + _
                        Right(oPdODueMassUpdateGrid.Columns.Item(idx - 1).TitleObject.Caption, 2)
                    oPdODueMassUpdateGrid.Columns.Item(idx).Width = 60
                Else
                    oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = _
                        Mid(oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                        Right(oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 2)
                    oPdODueMassUpdateGrid.Columns.Item(idx).Width = 40
                End If

                oPdODueMassUpdateGrid.Columns.Item(idx).Editable = False
                oPdODueMassUpdateGrid.Columns.Item(idx).RightJustified = True
                oPdODueMassUpdateGrid.Columns.Item(idx).TitleObject.Sortable = True

                'oPdODueMassUpdateGrid.Columns.Item(idx).TextStyle = 64
                'oItem = oPdODueMassUpdateGrid.Columns.Item(idx)
                'oColumn = oPdODueMassUpdateGrid.Columns.Item(idx)

                'oEditText = oPdODueMassUpdateGrid.Columns.Item(idx)
                'oEditText = oColumn
                'oEditText.SuppressZeros = True

            End If
        Next

        'For idx = 0 To CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
        '    oPdODueMassUpdateGrid.Columns.Item(10 + idx).Width = 60
        '    oPdODueMassUpdateGrid.Columns.Item(10 + idx).Editable = False
        'Next



        'oPdODueMassUpdateGrid.RowHeaders.Width = 20
        'oPdODueMassUpdateGrid.Columns.Item("#").Width = 30
        'oPdODueMassUpdateGrid.Columns.Item(1).Width = 20
        'oPdODueMassUpdateGrid.Columns.Item("SO Date").Width = 60
        'oPdODueMassUpdateGrid.Columns.Item("DocEntry").Width = 60
        'oPdODueMassUpdateGrid.Columns.Item("DocNum").Width = 60
        'oPdODueMassUpdateGrid.Columns.Item("Line").Width = 30
        'oPdODueMassUpdateGrid.Columns.Item("Cust. Code").Width = 80
        'oPdODueMassUpdateGrid.Columns.Item("FG").Width = 100
        'oPdODueMassUpdateGrid.Columns.Item("Exp Delivery Date").Width = 80
        'oPdODueMassUpdateGrid.Columns.Item("WhsCode").Width = 50
        'oPdODueMassUpdateGrid.Columns.Item("PanjangInCm").Width = 50
        'oPdODueMassUpdateGrid.Columns.Item("LebarInCm").Width = 50
        'oPdODueMassUpdateGrid.Columns.Item("SO_Bentuk").Width = 80


        'Dim oItem As SAPbouiCOM.Item

        ' ''oItem = oForm.Items.Item("OptimMtx").Specific
        'oItem = oForm.Items.Item("PdODueGrid")
        ' ''oItem.Height = 200
        ''oItem.Top = 135
        'oItem.Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        'oItem.Width = oForm.ClientWidth - 20

        oForm.Items.Item("PdODueGrid").Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        oForm.Items.Item("PdODueGrid").Width = oForm.ClientWidth - 20



        'oForm.Items.Item("AsOf").Left = oForm.ClientWidth - 250
        'oForm.Items.Item("ViewPeriod").Left = oForm.ClientWidth - 250

        'oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth - 120
        'oForm.Items.Item("IntervlDay").Left = oForm.ClientWidth - 120
        'oForm.Items.Item("Days").Left = oForm.ClientWidth - 50

        'oForm.Items.Item("NewDue").Left = oForm.ClientWidth / 3
        'oForm.Items.Item("ReasonChg").Left = oForm.ClientWidth / 3

        'oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth / 3 + 120
        'oForm.Items.Item("NewReason").Left = oForm.ClientWidth / 3 + 120


        oForm.Items.Item("NewDue").Left = oForm.ClientWidth - 320
        oForm.Items.Item("ReasonChg").Left = oForm.ClientWidth - 320
        oForm.Items.Item("cmdUpdPdO").Left = oForm.ClientWidth - 320

        oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth - 210
        oForm.Items.Item("NewReason").Left = oForm.ClientWidth - 210

        oForm.Items.Item("AsOf").Left = oForm.ClientWidth / 3
        oForm.Items.Item("ViewPeriod").Left = oForm.ClientWidth / 3

        oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("IntervlDay").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("Days").Left = oForm.ClientWidth / 3 + 180


        Dim sboDate As String
        Dim dDate As DateTime


        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oPdODueMassUpdateGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Sub RearrangePdoDueGridByMachine(ByVal oForm As SAPbouiCOM.Form)


        Dim oColumn As SAPbouiCOM.EditTextColumn


        Dim oPdODueMassUpdateByMachineGrid As SAPbouiCOM.Grid

        'oForm.Freeze(True)


        oPdODueMassUpdateByMachineGrid = oForm.Items.Item("PdOMchGrid").Specific


        'oPdODueMassUpdateByMachineGrid.RowHeaders.Width = 50

        ''Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        'oColumn = oPdODueMassUpdateByMachineGrid.Columns.Item("CustCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        'oColumn.Editable = False

        'oColumn = oPdODueMassUpdateByMachineGrid.Columns.Item("SO#")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        'oColumn.Editable = False


        'oColumn = oPdODueMassUpdateByMachineGrid.Columns.Item("ItemCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        'oColumn.Editable = False
        'oColumn.Width = 100

        oPdODueMassUpdateByMachineGrid.Columns.Item("PdOKey").Visible = False

        oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").TitleObject.Sortable = True

        oPdODueMassUpdateByMachineGrid.Columns.Item("PdO#").Width = 60
        oPdODueMassUpdateByMachineGrid.Columns.Item("SO#").Width = 75
        oPdODueMassUpdateByMachineGrid.Columns.Item("SODueDate").Width = 75
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdODocDate").Width = 75
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdODue").Width = 75
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdO Qty").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdO Qty").RightJustified = True

        oPdODueMassUpdateByMachineGrid.Columns.Item("CustCode").Width = 70
        oPdODueMassUpdateByMachineGrid.Columns.Item("P (cm)").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("L (cm)").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOLuasM2").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("UoM").Width = 30
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOLineNum").Width = 30
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOOverDue").Width = 50
        oPdODueMassUpdateByMachineGrid.Columns.Item("ItemName").Width = 80

        oPdODueMassUpdateByMachineGrid.Columns.Item("PdO#").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("SO#").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("SODueDate").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdODocDate").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdODue").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("PdO Qty").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("CustCode").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("P (cm)").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("L (cm)").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOLuasM2").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("UoM").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOLineNum").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("SOOverDue").Editable = False
        oPdODueMassUpdateByMachineGrid.Columns.Item("ItemName").Editable = False


        'oColumn = oPdODueMassUpdateByMachineGrid.Columns.Item("Customer Name")
        'oColumn.Editable = False

        'oPdODueMassUpdateByMachineGrid.Columns.Item("SO Date").Width = 80
        'oPdODueMassUpdateByMachineGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oPdODueMassUpdateByMachineGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oPdODueMassUpdateByMachineGrid.Columns.Item(0).TitleObject.Caption = oPdODueMassUpdateByMachineGrid.Rows.Count.ToString


        If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        And oPdODueMassUpdateByMachineGrid.DataTable.GetValue(0, 0) <> "" Then
            oForm.Items.Item("cmdUpdPdO").Enabled = True
        Else
            oForm.Items.Item("cmdUpdPdO").Enabled = False
        End If


        'oPdODueMassUpdateByMachineGrid.Columns.Item("#").Editable = False

        'In case to enable button 'Filter table...' then all column must be readonly editable = false

        'oPdODueMassUpdateByMachineGrid.Columns.Item(10).Width = 70
        'oPdODueMassUpdateByMachineGrid.Columns.Item(11).Width = 70
        'oPdODueMassUpdateByMachineGrid.Columns.Item(12).Width = 70
        'oPdODueMassUpdateByMachineGrid.Columns.Item(13).Width = 70
        'oPdODueMassUpdateByMachineGrid.Columns.Item(14).Width = 70

        '' How-To Change Column Title/Header Caption in Grid
        'oPdODueMassUpdateByMachineGrid.Columns.Item(10).Width = 60
        'oPdODueMassUpdateByMachineGrid.Columns.Item(10).Description = "08"
        'oPdODueMassUpdateByMachineGrid.Columns.Item(10).TitleObject.Caption = _
        '    Mid(oPdODueMassUpdateByMachineGrid.Columns.Item(10).TitleObject.Caption, 6, 2) + "/" + _
        '    Right(oPdODueMassUpdateByMachineGrid.Columns.Item(10).TitleObject.Caption, 2)


        ' Interval days + 2 extra days: one for < first column and one for > last column
        ' e.g: 14days, as of date 2012-02-08, [2012-02-07], [2012-02-08], [2012-02-09], ... [2012-02-22]

        Dim colcount As Integer = oPdODueMassUpdateByMachineGrid.Columns.Count

        'Dim oItem As SAPbouiCOM.Item
        'Dim oEditText As SAPbouiCOM.EditText

        For idx = 0 To colcount - 1 'CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
            If Left(oPdODueMassUpdateByMachineGrid.Columns.Item(idx).TitleObject.Caption, 1) = Left(Now.Year.ToString, 1) Then
                oPdODueMassUpdateByMachineGrid.Columns.Item(idx).TitleObject.Caption = _
                    Mid(oPdODueMassUpdateByMachineGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                    Right(oPdODueMassUpdateByMachineGrid.Columns.Item(idx).TitleObject.Caption, 2)
                oPdODueMassUpdateByMachineGrid.Columns.Item(idx).Width = 40
                oPdODueMassUpdateByMachineGrid.Columns.Item(idx).Editable = False
                oPdODueMassUpdateByMachineGrid.Columns.Item(idx).RightJustified = True
                'oPdODueMassUpdateByMachineGrid.Columns.Item(idx).TextStyle = 64
                'oItem = oPdODueMassUpdateByMachineGrid.Columns.Item(idx)
                'oColumn = oPdODueMassUpdateByMachineGrid.Columns.Item(idx)

                'oEditText = oPdODueMassUpdateByMachineGrid.Columns.Item(idx)
                'oEditText = oColumn
                'oEditText.SuppressZeros = True

            End If
        Next

        'For idx = 0 To CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
        '    oPdODueMassUpdateByMachineGrid.Columns.Item(10 + idx).Width = 60
        '    oPdODueMassUpdateByMachineGrid.Columns.Item(10 + idx).Editable = False
        'Next



        'oPdODueMassUpdateByMachineGrid.RowHeaders.Width = 20
        'oPdODueMassUpdateByMachineGrid.Columns.Item("#").Width = 30
        'oPdODueMassUpdateByMachineGrid.Columns.Item(1).Width = 20
        'oPdODueMassUpdateByMachineGrid.Columns.Item("SO Date").Width = 60
        'oPdODueMassUpdateByMachineGrid.Columns.Item("DocEntry").Width = 60
        'oPdODueMassUpdateByMachineGrid.Columns.Item("DocNum").Width = 60
        'oPdODueMassUpdateByMachineGrid.Columns.Item("Line").Width = 30
        'oPdODueMassUpdateByMachineGrid.Columns.Item("Cust. Code").Width = 80
        'oPdODueMassUpdateByMachineGrid.Columns.Item("FG").Width = 100
        'oPdODueMassUpdateByMachineGrid.Columns.Item("Exp Delivery Date").Width = 80
        'oPdODueMassUpdateByMachineGrid.Columns.Item("WhsCode").Width = 50
        'oPdODueMassUpdateByMachineGrid.Columns.Item("PanjangInCm").Width = 50
        'oPdODueMassUpdateByMachineGrid.Columns.Item("LebarInCm").Width = 50
        'oPdODueMassUpdateByMachineGrid.Columns.Item("SO_Bentuk").Width = 80


        'Dim oItem As SAPbouiCOM.Item

        ' ''oItem = oForm.Items.Item("OptimMtx").Specific
        'oItem = oForm.Items.Item("PdOMchGrid")
        ' ''oItem.Height = 200
        ''oItem.Top = 135
        'oItem.Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        'oItem.Width = oForm.ClientWidth - 20

        oForm.Items.Item("PdOMchGrid").Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        oForm.Items.Item("PdOMchGrid").Width = oForm.ClientWidth - 20



        'oForm.Items.Item("AsOf").Left = oForm.ClientWidth - 250
        'oForm.Items.Item("ViewPeriod").Left = oForm.ClientWidth - 250

        'oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth - 120
        'oForm.Items.Item("IntervlDay").Left = oForm.ClientWidth - 120
        'oForm.Items.Item("Days").Left = oForm.ClientWidth - 50

        'oForm.Items.Item("NewDue").Left = oForm.ClientWidth / 3
        'oForm.Items.Item("ReasonChg").Left = oForm.ClientWidth / 3

        'oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth / 3 + 120
        'oForm.Items.Item("NewReason").Left = oForm.ClientWidth / 3 + 120


        oForm.Items.Item("NewDue").Left = oForm.ClientWidth - 320
        oForm.Items.Item("ReasonChg").Left = oForm.ClientWidth - 320
        oForm.Items.Item("cmdUpdPdO").Left = oForm.ClientWidth - 320

        oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth - 210
        oForm.Items.Item("NewReason").Left = oForm.ClientWidth - 210

        oForm.Items.Item("AsOf").Left = oForm.ClientWidth / 3
        oForm.Items.Item("ViewColumn").Left = oForm.ClientWidth / 3

        oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("IntervlMch").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("MachCol").Left = oForm.ClientWidth / 3 + 180


        Dim sboDate As String
        Dim dDate As DateTime


        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oPdODueMassUpdateByMachineGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Function ValidateInputDueDate_Form_PdODueDateMassUpdation(ByVal oForm As SAPbouiCOM.Form)

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

    Sub LoadPdO(ByVal oForm As SAPbouiCOM.Form)
        Dim STMTQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            'oApp.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            'oApp.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("PdODueTo").Specific.string = "" Then
            oForm.Items.Item("PdODueTo").Specific.string = oForm.Items.Item("PdODueFrom").Specific.string
        End If


        STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "' "

        'STMTQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, "

        '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T1.LineStatus = 'O' " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        RearrangePdoDueGrid(oForm)

    End Sub


    Sub LoadPdODueDateByMachine(ByVal oForm As SAPbouiCOM.Form)
        Dim STMTQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            'oApp.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            'oApp.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("PdODueTo").Specific.string = "" Then
            oForm.Items.Item("PdODueTo").Specific.string = oForm.Items.Item("PdODueFrom").Specific.string
        End If


        STMTQuery = "EXEC GetPdODueDateListByMachine_ByRangeofPdoDueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlMch").Specific.string = "", 0, oForm.Items.Item("IntervlMch").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "' "

        'STMTQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, "

        '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T1.LineStatus = 'O' " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        RearrangePdoDueGridByMachine(oForm)

    End Sub

    '    Sub GeneratePdOFromSO(ByVal oForm As SAPbouiCOM.Form)
    '        'On Error GoTo errHandler

    '        Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

    '        Dim idx As Long


    '        'Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
    '        'Dim oSalesOrderLines As SAPbobsCOM.Document_Lines = Nothing

    '        Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
    '        Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines = Nothing

    '        Dim errConnect As String = ""

    '        'Dim oPdODocSeriesRec As SAPbobsCOM.Recordset

    '        Dim strQry As String = ""
    '        Dim oPdODocSeriesOrder As String = ""
    '        Dim oPdODocSeriesJasa As String = ""

    '        Dim firstLine As Boolean = True
    '        Dim IsPdOLines_Generated As Boolean = False
    '        Dim atLeastOnePdOLines_notGenerated As Boolean = False

    '        oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific



    '        'GRID - Order by column checkbox
    '        oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


    '        ''Get PdO Doc. Series 
    '        ''oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        'oPdODocSeriesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ' FG-002: PENJUALAN JASA (SERIENAME:2011JS), FG-001: PENJUALAN ORDER (SERIENAME:2011) 


    '        'strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) = 'JS' AND Indicator = YEAR(GETDATE()) "

    '        'oPdODocSeriesRec.DoQuery(strQry)
    '        ''??? 
    '        'If oPdODocSeriesRec.RecordCount <> 0 Then
    '        '    oPdODocSeriesJasa = oPdODocSeriesRec.Fields.Item("Series").Value
    '        'Else
    '        '    MsgBox("Production Order Document Series Jasa Tidak ada, Mohon Setup PdO Document Series!")
    '        '    Exit Sub
    '        'End If

    '        'strQry = "SELECT TOP 1 Series FROM NNM1 WHERE ObjectCode = '202' AND RIGHT(SeriesName, 2) <> 'JS' AND Indicator = YEAR(GETDATE()) "
    '        'oPdODocSeriesRec.DoQuery(strQry)
    '        ''??? 
    '        'If oPdODocSeriesRec.RecordCount <> 0 Then
    '        '    oPdODocSeriesOrder = oPdODocSeriesRec.Fields.Item("Series").Value
    '        'Else
    '        '    MsgBox("Production Order Document Series Kaca Order Tidak ada, Mohon Setup PdO Document Series!")
    '        '    Exit Sub
    '        'End If

    '        'System.Runtime.InteropServices.Marshal.ReleaseComObject(oPdODocSeriesRec)
    '        'oPdODocSeriesRec = Nothing
    '        'GC.Collect()

    '        If oPdODocSeriesOrder <> "" And oPdODocSeriesJasa <> "" Then


    '            For idx = oPdODueMassUpdateGrid.Rows.Count - 1 To 0 Step -1
    '                oForm.Items.Item("SoNumber").Click()
    '                'SBO_Application.SetStatusBarMessage("Generating PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '                oApp.SetStatusBarMessage("Revising Due Date PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)

    '                If oPdODueMassUpdateGrid.DataTable.GetValue(1, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx)) = "Y" Then

    '                    oForm.Items.Item("SoNumber").Click()


    '                    If Not oCompany.InTransaction Then
    '                        oCompany.StartTransaction()
    '                    End If


    '                    Dim oRS As SAPbobsCOM.Recordset

    '                    'vCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                    oRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    strQry = "SELECT DocNum FROM OWOR WHERE Status <> 'C' AND OriginNum =  " & oPdODueMassUpdateGrid.DataTable.GetValue(4, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) _
    '                        & " AND ItemCode = '" & oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) & "' "
    '                    oRS.DoQuery(strQry)
    '                    'oRS.DoQuery("UPDATE RDR1 SET U_bacthNum = 'b321' where docentry = 249 and linenum = 0")

    '                    oForm.Items.Item("BPCardCode").Click()


    '                    'If oRS.RecordCount = 0 Then -- if duplicate don't insert PdO
    '                    If oRS.RecordCount <> 0 Or oRS.RecordCount = 0 Then

    '                        'Check Order Kaca(WhsCode = FG-001) maka - Generate PdO dgn satu PdO Line dan ItemCode = XDUMMY
    '                        If oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" _
    '                        Or (oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "FG-002" _
    '                            And Left(oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString), 2) <> "KT") Then

    '                            oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

    '                            ' Series PdO JS = 202 (NNM1) objectCode = 202 (OWOR PdO) series id = 45

    '                            ' IMPORTANT !!!
    '                            ' PdO SERIES YEAR 2011, 2011JS PdO JASA, SERIES# = 45
    '                            ' PdO SERIES YEAR 2011, 2011   PdO KACA SERIES# = 27 

    '                            If oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" Then
    '                                'oProd1.Series = 27
    '                                oProd1.Series = oPdODocSeriesOrder
    '                            Else
    '                                'oProd1.Series = 45
    '                                oProd1.Series = oPdODocSeriesJasa
    '                            End If

    '                            'oProd1.ItemNo = "KTF12CLXX589"
    '                            oProd1.ItemNo = oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            oProd1.PlannedQuantity = oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            ''oProdOrder.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(13, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            'PdO Posting Date = SO Posting Date
    '                            'oProd1.PostingDate = oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                            oProd1.PostingDate = Format(Now, "yyyy-MM-dd")

    '                            Dim dueDt As DateTime
    '                            Dim sodt As DateTime
    '                            Dim sodelivdt As DateTime
    '                            Dim dtdiff As Integer

    '                            sodt = oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                            sodelivdt = oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                            dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                            dtdiff = DateDiff(DateInterval.Day, CDate(oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)), CDate(oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)))
    '                            'sodelivdt = ""
    '                            dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                            dueDt = DateAdd(DateInterval.Day, IIf(dtdiff < 0, 0, dtdiff), Now)

    '                            'PdO Due Date = SO Deliv. Date
    '                            'oProd1.DueDate = Today + n days (so date - so deliv date)
    '                            ''oProd1.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                            oProd1.DueDate = dueDt

    '                            'oprod1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
    '                            oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
    '                            oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
    '                            'oprod1.Warehouse = "01"
    '                            'oProd1.Warehouse = "FG-001"

    '                            oProd1.Warehouse = oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            oProd1.CustomerCode = oPdODueMassUpdateGrid.DataTable.GetValue(7, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                            oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
    '                            ' so docnum
    '                            oProd1.ProductionOrderOriginEntry = oPdODueMassUpdateGrid.DataTable.GetValue(3, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            oProd1.UserFields.Fields.Item("U_PoD_Pcm").Value = oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                            oProd1.UserFields.Fields.Item("U_PdO_Lcm").Value = oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                            oProd1.UserFields.Fields.Item("U_PdO_Bentuk").Value = oPdODueMassUpdateGrid.DataTable.GetValue(17, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                            oProd1.UserFields.Fields.Item("U_SO_Luas_M2").Value = _
    '                            Left(CStr( _
    '                                Math.Round( _
    '                                  (IIf(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                  IIf(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                  IIf(oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) / 10000) _
    '                                  , 4) _
    '                                ) _
    '                            , 10)

    '                            oProd1.UserFields.Fields.Item("U_ORDRDocEntry").Value = _
    '                            oPdODueMassUpdateGrid.DataTable.GetValue(3, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString()
    '                            oProd1.UserFields.Fields.Item("U_ORDRLineNum").Value = _
    '                            oPdODueMassUpdateGrid.DataTable.GetValue(19, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString()

    '                            oProdLine1 = oProd1.Lines

    '                            ' Generate one line - Dummy item
    '                            oProdLine1.ItemNo = "XDUMMY"
    '                            oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                            oProdLine1.Warehouse = "SRV-DL"


    '                            'MsgBox(GC.GetTotalMemory(True))

    '                            lRetCode = oProd1.Add()

    '                            If lRetCode <> 0 Then
    '                                oCompany.GetLastError(lErrCode, sErrMsg)

    '                                oApp.MessageBox(lErrCode & ": " & sErrMsg)


    '                                If oCompany.InTransaction Then
    '                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                                End If
    '                            Else

    '                                ''vCompany.GetNewObjectCode(tmpKey)
    '                                ''vCompany.GetNewObjectCode(PdOno)
    '                                'oCompany.GetNewObjectCode(PdOno)
    '                                'tmpKey = Convert.ToInt32(PdOno)

    '                                ' !!!! Make sure before create another object type-> clear previous/current object type.
    '                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
    '                                oProdLine1 = Nothing

    '                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
    '                                oProd1 = Nothing

    '                                GC.Collect()

    '                                oForm.Items.Item("SoNumber").Click()



    '                                If oCompany.InTransaction Then
    '                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                                End If

    '                            End If

    '                            'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '                            oApp.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)


    '                        ElseIf oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "FG-002" _
    '                            And Left(oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString), 2) = "KT" Then
    '                            'Jika Order Jasa (WhsCode = FG-002) maka:
    '                            'Check - Jika menemukan mesin di Machine by Ukuran Kaca, maka generate PdO Lines sesuai machine code yg ada
    '                            ' Selain itu JANGAN Generate PdO

    '                            Dim oRS_MachineKaca As SAPbobsCOM.Recordset

    '                            oRS_MachineKaca = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                            strQry = " SELECT U_MIS_LuasM2Min, U_MIS_LuasM2Max, U_MIS_MachineCode " & _
    '                                " FROM [@MIS_MACHINEKACA] " & _
    '                                " WHERE U_MIS_ItemCode = '" & Left(oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString), 5) & "' " & _
    '                                " AND U_MIS_WhsCode = '" & oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) & "' " & _
    '                                " AND " & _
    '                                CStr( _
    '                                    Math.Round( _
    '                                      (IIf(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                      IIf(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) _
    '                                      / 10000) _
    '                                      , 4) _
    '                                    ) & _
    '                                " BETWEEN U_MIS_LuasM2Min AND U_MIS_LuasM2Max "


    '                            oRS_MachineKaca.DoQuery(strQry)


    '                            If oRS_MachineKaca.RecordCount > 0 Then
    '                                oRS_MachineKaca.MoveFirst()

    '                                firstLine = True

    '                                oProd1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)


    '                                ' Series PdO JS = 202 (NNM1) objectCode = 202 (OWOR PdO) series id = 45

    '                                ' IMPORTANT !!!
    '                                ' PdO SERIES YEAR 2011, 2011JS PdO JASA, SERIES# = 45
    '                                ' PdO SERIES YEAR 2011, 2011   PdO KACA SERIES# = 27 

    '                                If oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "FG-001" Then
    '                                    'oProd1.Series = 27
    '                                    oProd1.Series = oPdODocSeriesOrder
    '                                Else
    '                                    'oProd1.Series = 45
    '                                    oProd1.Series = oPdODocSeriesJasa
    '                                End If

    '                                'oProd1.ItemNo = "KTF12CLXX589"
    '                                oProd1.ItemNo = oPdODueMassUpdateGrid.DataTable.GetValue(9, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                oProd1.PlannedQuantity = oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                ''oProdOrder.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(13, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                'PdO Posting Date = SO Posting Date
    '                                'oProd1.PostingDate = oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                                oProd1.PostingDate = Format(Now, "yyyy-MM-dd")

    '                                Dim dueDt As DateTime
    '                                Dim sodt As DateTime
    '                                Dim sodelivdt As DateTime
    '                                Dim dtdiff As Integer

    '                                sodt = oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                                sodelivdt = oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                                dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                                dtdiff = DateDiff(DateInterval.Day, CDate(oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)), CDate(oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)))
    '                                'sodelivdt = ""
    '                                dtdiff = DateDiff(DateInterval.Day, sodt, sodelivdt)
    '                                dueDt = DateAdd(DateInterval.Day, IIf(dtdiff < 0, 0, dtdiff), Now)

    '                                'PdO Due Date = SO Deliv. Date
    '                                'oProd1.DueDate = Today + n days (so date - so deliv date)
    '                                ''oProd1.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(14, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                                oProd1.DueDate = dueDt

    '                                'oprod1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
    '                                oProd1.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
    '                                oProd1.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
    '                                'oprod1.Warehouse = "01"
    '                                'oProd1.Warehouse = "FG-001"

    '                                oProd1.Warehouse = oPdODueMassUpdateGrid.DataTable.GetValue(12, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                oProd1.CustomerCode = oPdODueMassUpdateGrid.DataTable.GetValue(7, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
    '                                oProd1.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
    '                                ' so docnum
    '                                oProd1.ProductionOrderOriginEntry = oPdODueMassUpdateGrid.DataTable.GetValue(3, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                oProd1.UserFields.Fields.Item("U_PoD_Pcm").Value = oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                                oProd1.UserFields.Fields.Item("U_PdO_Lcm").Value = oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString
    '                                oProd1.UserFields.Fields.Item("U_PdO_Bentuk").Value = oPdODueMassUpdateGrid.DataTable.GetValue(17, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)

    '                                oProd1.UserFields.Fields.Item("U_SO_Luas_M2").Value = _
    '                                Left(CStr( _
    '                                    Math.Round( _
    '                                      (IIf(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                      IIf(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                      IIf(oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(11, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) / 10000) _
    '                                      , 4) _
    '                                    ) _
    '                                , 10)

    '                                oProd1.UserFields.Fields.Item("U_ORDRDocEntry").Value = _
    '                                oPdODueMassUpdateGrid.DataTable.GetValue(3, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString()
    '                                oProd1.UserFields.Fields.Item("U_ORDRLineNum").Value = _
    '                                oPdODueMassUpdateGrid.DataTable.GetValue(19, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString()


    '                                While Not oRS_MachineKaca.EoF

    '                                    Dim oRS_OITM As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    strQry = "Select T0.ItemCode ItemCode " & _
    '                                        "FROM OITM T0 " & _
    '                                        "WHERE T0.ItemCode like '%" & _
    '                                            oRS_MachineKaca.Fields.Item("U_MIS_MachineCode").Value & "%' "

    '                                    oRS_OITM.DoQuery(strQry)



    '                                    If oRS_OITM.RecordCount > 0 Then
    '                                        oRS_OITM.MoveFirst()

    '                                        oProdLine1 = oProd1.Lines

    '                                        IsPdOLines_Generated = True

    '                                        'If oProdLine1.Count > 0 Then
    '                                        If oProdLine1.Count > 0 And String.IsNullOrEmpty(oProdLine1.ItemNo) = False Then
    '                                            firstLine = False
    '                                        End If

    '                                        For Row_oitm = 1 To oRS_OITM.RecordCount

    '                                            'oProdLine1 = oProd1.Lines

    '                                            '' Generate one line - Dummy item
    '                                            'oProdLine1.ItemNo = "XDUMMY"
    '                                            'oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
    '                                            'oProdLine1.Warehouse = "SRV-DL"

    '                                            If firstLine = False Then
    '                                                'oProd1.Lines.Add()
    '                                                oProdLine1.Add()
    '                                            Else
    '                                                firstLine = False
    '                                            End If

    '                                            If oRS_OITM.Fields.Item("ItemCode").Value <> "" Then
    '                                                '???
    '                                                'oProdLine1.ItemNo = "xdummy" 'oRS_OITM.Fields.Item("ItemCode").Value
    '                                                'oProdLine1.UserFields.Fields.Item("U_NBS_RunTime").Value = _
    '                                                'IIf(Left(oRS_OITM.Fields.Item("ItemCode").Value, 3) = "XLG", 1, _
    '                                                '    333) 'oRS_OITM.Fields.Item("U_MIS_MachineCode").Value

    '                                                oProdLine1.ItemNo = oRS_OITM.Fields.Item("ItemCode").Value ' "XDUMMY"
    '                                                oProdLine1.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush

    '                                                'case FG, planned qty FG = 5
    '                                                ' in case machine = XLG, eg: XLG2 then runtime = 5 x 1 = 5
    '                                                ' in case machine <> XLG, eg: XTP3, kaca P cm x L cm, then Runtime (fg qty x luas M2) = 5 x (258.9 x 70 / 10000)
    '                                                oProdLine1.UserFields.Fields.Item("U_NBS_RunTime").Value = _
    '                                                IIf(Left(oRS_OITM.Fields.Item("ItemCode").Value, 3) = "XLG", _
    '                                                    1 _
    '                                                    , _
    '                                                    Math.Round( _
    '                                                    (IIf(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                                     IIf(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) _
    '                                                     / 10000) _
    '                                                     , 4) _
    '                                                )


    '                                                oProdLine1.BaseQuantity = _
    '                                                IIf(Left(oRS_OITM.Fields.Item("ItemCode").Value, 3) = "XLG", _
    '                                                    1 _
    '                                                    , _
    '                                                    Math.Round( _
    '                                                    (IIf(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(15, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) * _
    '                                                     IIf(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString = "", 0, CDbl(oPdODueMassUpdateGrid.DataTable.GetValue(16, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString).ToString)) _
    '                                                     / 10000) _
    '                                                     , 4) _
    '                                                )



    '                                            End If
    '                                            oRS_OITM.MoveNext()
    '                                        Next
    '                                        'Else
    '                                        '    oApp.SetStatusBarMessage("Please Check Item Master!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
    '                                        '    Exit Sub
    '                                    End If

    '                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS_OITM)
    '                                    oRS_OITM = Nothing

    '                                    'MsgBox(GC.GetTotalMemory(True))

    '                                    oRS_MachineKaca.MoveNext()
    '                                End While


    '                                If IsPdOLines_Generated = True Then
    '                                    lRetCode = oProd1.Add()


    '                                    If lRetCode <> 0 Then
    '                                        oCompany.GetLastError(lErrCode, sErrMsg)

    '                                        oApp.MessageBox(lErrCode & ": " & sErrMsg)

    '                                        If oCompany.InTransaction Then
    '                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                                        End If
    '                                    Else

    '                                        ' !!!! Make sure before create another object type-> clear previous/current object type.
    '                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
    '                                        oProdLine1 = Nothing



    '                                        If oCompany.InTransaction Then
    '                                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                                        End If

    '                                    End If
    '                                Else
    '                                    ' PDO LINE NOT GENERATED THEN ROLLBACK!!!
    '                                    atLeastOnePdOLines_notGenerated = True

    '                                    If oCompany.InTransaction Then
    '                                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                                    End If

    '                                End If

    '                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS_MachineKaca)
    '                                oRS_MachineKaca = Nothing

    '                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
    '                                oProd1 = Nothing

    '                                GC.Collect()

    '                                oForm.Items.Item("SoNumber").Click()


    '                                'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '                                oApp.SetStatusBarMessage("Generating PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    '                            Else
    '                                atLeastOnePdOLines_notGenerated = True
    '                            End If
    '                            ' End-IF Check - Jika menemukan mesin di Machine by Ukuran Kaca, maka generate PdO Lines sesuai machine code yg ada
    '                            ' Selain itu JANGAN Generate PdO

    '                        End If ' End If - Check Order Kaca atau Order Jasa

    '                    End If ' End-If oRS.RecordCount = 0 Then -- if duplicate don't insert PdO


    '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
    '                    oRS = Nothing

    '                    ' by Toin 2011-02-09 Check Duplicate PdO before Generate PdO


    '                    GC.Collect()
    '                    'MsgBox(GC.GetTotalMemory(True))

    '                    'MsgBox("generating... PdO; DONE!!!")

    '                Else
    '                    Exit For
    '                End If
    '            Next

    '        End If  ' Checking PdO Series

    '        'oApp.MessageBox("Generating PdO.... Finished !!! ", 1, "Ok")
    '        If atLeastOnePdOLines_notGenerated = True Then
    '            oApp.MessageBox("Revising Due Date PdO.... Finished !!! " & _
    '                            "Ada PdO yg tidak berhasil direvisi! Perhatikan dan pastikan ada mapping data di table master Machine by Ukuran Kaca!!! ", 1, "Ok")
    '        Else
    '            oApp.MessageBox("Revising Due Date PdO.... Finished !!! ", 1, "Ok")

    '        End If

    '        Exit Sub


    'errHandler:
    '        MsgBox("Exception: " & Err.Description)
    '        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '    End Sub

    Sub GeneratePdODueDateMassUpdate(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

        Dim idx As Long

        oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

        'GRID - Order by column checkbox
        oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        'Loop only selected/checked in grid rows and exit.
        For idx = oPdODueMassUpdateGrid.Rows.Count - 1 To 0 Step -1

            oApp.SetStatusBarMessage("Revising Due Date PdO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            If UCase(oPdODueMassUpdateGrid.DataTable.GetValue(0, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx))) = "Y" Then
                'If UCase(oPdODueMassUpdateGrid.DataTable.GetValue(0, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx))) = "Y" _
                '    And oPdODueMassUpdateGrid.DataTable.GetValue(1, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "61061" Then

                Dim oProductionOrder As SAPbobsCOM.ProductionOrders = Nothing

                'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
                'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines



                'Dim isconnect As Long
                Dim errConnect As String = ""


                If Not oCompany.InTransaction Then
                    oCompany.StartTransaction()
                End If


                Dim PdOno As String = ""

                oProductionOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                'oPDOStatus.GetByKey(oPdODueMassUpdateGrid.DataTable.GetValue(2, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString))
                oProductionOrder.GetByKey(oPdODueMassUpdateGrid.DataTable.GetValue(1, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString))


                'oProductionOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                'oProductionOrder.UserFields.Fields.Item("U_MIS_Progress").Value = "Released"

                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate03").Value = oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate02").Value
                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReas03").Value = oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReas02").Value
                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate02").Value = oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value
                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReas02").Value = oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReas01").Value

                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value = oProductionOrder.DueDate
                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReas01").Value = oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReason").Value

                'oProductionOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value = oProductionOrder.DueDate
                oProductionOrder.DueDate = CDate(oForm.Items.Item("NewDueDate").Specific.string)
                oProductionOrder.UserFields.Fields.Item("U_MIS_DueDateReason").Value = oForm.Items.Item("NewReason").Specific.string

                lRetCode = oProductionOrder.Update()


                If lRetCode <> 0 Then
                    'vCompany.GetLastError(lErrCode, sErrMsg)
                    'SBO_Application.MessageBox(lErrCode & ": " & sErrMsg)
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    oApp.MessageBox(lErrCode & ": " & sErrMsg)

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                Else

                    '' !!!! Make sure before create another object type-> clear previous/current object type.
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdLine1)
                    'oProdLine1 = Nothing

                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProd1)
                    'oProd1 = Nothing


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionOrder)
                    oProductionOrder = Nothing

                    'vCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                    If oCompany.InTransaction Then
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    End If

                End If


                'vCompany.Disconnect()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(vCompany)
                'vCompany = Nothing

                GC.Collect()
            Else
                Exit For
            End If
        Next

        'MsgBox("Begin trx: generating... PdO")
        'SBO_Application.SetStatusBarMessage("Generating PdO.... Finished !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        oApp.SetStatusBarMessage("Revising PdO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)


        'MsgBox("generating... PdO; DONE!!!")

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub

End Module
