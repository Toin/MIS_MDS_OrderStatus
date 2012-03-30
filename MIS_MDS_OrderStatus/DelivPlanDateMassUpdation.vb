Module DelivPlanDateMassUpdation

    Sub DelivPlanDateMassUpdation_FormEntry()
        Dim oForm As SAPbouiCOM.Form

        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button

        Dim oSODelvMassUpdateGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord2")
            oApp.MessageBox("Form Already Open")
        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams

            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord2"
            fcp.UniqueID = "mds_ord2"
            'fcp.ObjectType = "MIS_OPTIM"

            'fcp.XmlData = LoadFromXML("sotomfg.srf")
            'fcp.XmlData = MenuCreation.LoadFromXML("sotomfg.srf")
            fcp.XmlData = LoadFromXML("DelivPlanDateMassUpdation.srf")

            oForm = oApp.Forms.AddEx(fcp)

            oForm.Freeze(True)


            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("SODueFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("SODueTo", SAPbouiCOM.BoDataType.dt_DATE)

            oForm.DataSources.UserDataSources.Add("SOFull", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("SOPartial", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("AsOfDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("IntervlDay", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)

            oForm.DataSources.UserDataSources.Add("NewDueDate", SAPbouiCOM.BoDataType.dt_DATE)
            oForm.DataSources.UserDataSources.Add("NewReason", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            'Default value for PdO Due Date From - To, As of Date, Interval Days
            oForm.DataSources.UserDataSources.Item("AsOfDate").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("IntervlDay").Value = "14"
            oForm.DataSources.UserDataSources.Item("SOFull").Value = "1"
            'oForm.DataSources.UserDataSources.Item("SOPartial").Value = "0"

            oForm.DataSources.UserDataSources.Item("SODueFrom").Value = DateTime.Today.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Item("SODueTo").Value = DateTime.Today.AddDays(14).ToString("yyyyMMdd")

            'Set value for User DataSource
            oForm.DataSources.UserDataSources.Add("BPDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oForm.DataSources.UserDataSources.Add("SoNumber", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)


            oEditText = oForm.Items.Item("BPCardCode").Specific
            oButton = oForm.Items.Item("BPButton").Specific

            oEditText.DataBind.SetBound(True, "", "BPDS")


            oEditText = oForm.Items.Item("SoNumber").Specific

            oForm.Items.Item("SODueFrom").Width = 100
            oEditText = oForm.Items.Item("SODueFrom").Specific
            oEditText.DataBind.SetBound(True, "", "SODueFrom")

            oForm.Items.Item("SODueTo").Width = 100
            oEditText = oForm.Items.Item("SODueTo").Specific
            oEditText.DataBind.SetBound(True, "", "SODueTo")

            oForm.Items.Item("AsOfDate").Width = 100
            oEditText = oForm.Items.Item("AsOfDate").Specific
            oEditText.DataBind.SetBound(True, "", "AsOfDate")

            oForm.Items.Item("IntervlDay").Width = 50
            oEditText = oForm.Items.Item("IntervlDay").Specific
            oEditText.DataBind.SetBound(True, "", "IntervlDay")


            Dim optButton As SAPbouiCOM.OptionBtn


            oForm.Items.Item("optSOFull").Width = 160
            oItem = oForm.Items.Item("optSOFull")
            optButton = oItem.Specific
            optButton.DataBind.SetBound(True, "", "SOFull")

            oForm.Items.Item("optSOHalf").Width = 160
            oItem = oForm.Items.Item("optSOHalf")
            optButton = oItem.Specific
            'optButton.DataBind.SetBound(True, "", "SOPartial")
            optButton.DataBind.SetBound(True, "", "SOFull")
            optButton.GroupWith("optSOFull")

            oForm.Items.Item("NewDueDate").Width = 100
            oEditText = oForm.Items.Item("NewDueDate").Specific
            oEditText.DataBind.SetBound(True, "", "NewDueDate")

            'oForm.Items.Item("NewReason").Width = 50
            oEditText = oForm.Items.Item("NewReason").Specific
            oEditText.DataBind.SetBound(True, "", "NewReason")

            oItem = oForm.Items.Item("SODelvGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200


            oSODelvMassUpdateGrid = oItem.Specific

            oSODelvMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            Dim SODuefrom As String
            Dim SODueto As String
            Dim intervalddays As Integer
            Dim asofdate As String

            intervalddays = IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string)
            asofdate = Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd")

            SODuefrom = Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd")
            SODueto = Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd")

            STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
                IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
                + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
                + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
                + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
                + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
                + "', " + IIf(oForm.DataSources.UserDataSources.Item("SOFull").Value = "", "0", oForm.DataSources.UserDataSources.Item("SOFull").Value)

            'Dim optButton As SAPbouiCOM.OptionBtn
            Dim isSOFull As Boolean

            isSOFull = oForm.DataSources.UserDataSources.Item("SOFull").Value

            'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
            '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("PdODueFrom").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("PdODueTo").Specific.string), "yyyyMMdd") & "' " _
            '& " AND T0.CardCode = '" & oForm.Items.Item("BPCardCode").Specific.string & "' " _



            ' Grid #: 1
            oForm.DataSources.DataTables.Add("SODelvList")
            oForm.DataSources.DataTables.Item("SODelvList").ExecuteQuery(STMTQuery)
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


        'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer 14, '20110801', '20110501', '20111231', 0"
        STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "', " + IIf(oForm.DataSources.UserDataSources.Item("SOFull").Value = "", "0", oForm.DataSources.UserDataSources.Item("SOFull").Value)


        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        oForm.Items.Item("BPCardCode").Click()


        oForm.Top = 250
        oForm.Left = 100
        oForm.Width = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Width - 100
        oForm.Height = System.Windows.Forms.SystemInformation.MaxWindowTrackSize.Height - 200


        RearrangeSODelivGrid(oForm)


        oForm.Visible = True

    End Sub


    Sub RearrangeSODelivGrid(ByVal oForm As SAPbouiCOM.Form)


        Dim oColumn As SAPbouiCOM.EditTextColumn


        Dim oSODelvMassUpdateGrid As SAPbouiCOM.Grid

        oForm.Freeze(True)


        oSODelvMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific


        'oSODelvMassUpdateGrid.RowHeaders.Width = 50

        'Dim OCOL As SAPbouiCOM.Column
        'Dim ocols As SAPbouiCOM.Columns
        'ocols = oSODelvMassUpdateGrid.Columns

        'OCOL = ocols.Item("SO#")
        'OCOL.DataBind.SetBound(True, "", "SO#")
        'OCOL.ExtendedObject.linkedobject = SAPbouiCOM.BoLinkedObject.lf_Order

        'OCOL = oSODelvMassUpdateGrid.Columns.Item("SO#")
        'OCOL.DataBind.SetBound(True, "", "SO#")
        'OCOL.ExtendedObject.linkedobject = SAPbouiCOM.BoLinkedObject.lf_Order


        ''Adding LinkedButton (Orange) : Set Property-> LinkedObjectType
        'oColumn = oSODelvMassUpdateGrid.Columns.Item("Cust Code")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner '"2" -> BP Master
        'oColumn.Editable = False

        'oColumn = oSODelvMassUpdateGrid.Columns.Item("DocEntry")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Order '"2"
        'oColumn.Editable = False


        'oColumn = oSODelvMassUpdateGrid.Columns.Item("ItemCode")
        'oColumn.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Items '"2"
        'oColumn.Editable = False
        'oColumn.Width = 100

        'oSODelvMassUpdateGrid.Columns.Item("DocEntry").Visible = False

        oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").TitleObject.Sortable = True

        oSODelvMassUpdateGrid.Columns.Item("Cust Code").Width = 70
        oSODelvMassUpdateGrid.Columns.Item("Customer Name").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("DocEntry").Width = 60
        oSODelvMassUpdateGrid.Columns.Item("SO#").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("SO Date").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("SO DueDate").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("Over").Width = 20
        oSODelvMassUpdateGrid.Columns.Item("Over").ForeColor = ColorTranslator.ToOle(Color.Red)

        oSODelvMassUpdateGrid.Columns.Item("SOOverDue").Width = 35
        oSODelvMassUpdateGrid.Columns.Item("Line").Width = 45
        oSODelvMassUpdateGrid.Columns.Item("Line").RightJustified = True

        oSODelvMassUpdateGrid.Columns.Item("ItemCode").Width = 80
        oSODelvMassUpdateGrid.Columns.Item("ItemName").Width = 80
        oSODelvMassUpdateGrid.Columns.Item("SOOpenQty").Width = 60
        oSODelvMassUpdateGrid.Columns.Item("SOOpenQty").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("Deliv Date").Width = 75
        oSODelvMassUpdateGrid.Columns.Item("Deliv Reason").Width = 120
        oSODelvMassUpdateGrid.Columns.Item("SO TotalQty").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("SO TotalQty").RightJustified = True
        oSODelvMassUpdateGrid.Columns.Item("SO LINESTS").Width = 50
        oSODelvMassUpdateGrid.Columns.Item("SOLine").Width = 30
        oSODelvMassUpdateGrid.Columns.Item("SOLine").RightJustified = True

        oSODelvMassUpdateGrid.Columns.Item("Cust Code").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Cust Code").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Customer Name").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Customer Name").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("DocEntry").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("DocEntry").Visible = False
        oSODelvMassUpdateGrid.Columns.Item("SOLine").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SOLine").Visible = False

        oSODelvMassUpdateGrid.Columns.Item("SO#").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO#").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO Date").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO Date").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO DueDate").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO DueDate").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SOOverDue").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SOOverDue").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Over").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Over").TitleObject.Sortable = True

        oSODelvMassUpdateGrid.Columns.Item("Line").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Line").TitleObject.Sortable = True

        oSODelvMassUpdateGrid.Columns.Item("ItemCode").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("ItemCode").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("ItemName").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("ItemName").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SOOpenQty").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SOOpenQty").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Stock Avail").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Deliv Date").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Deliv Date").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("Deliv Reason").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("Deliv Reason").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO TotalQty").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO TotalQty").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO LINESTS").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("SO LINESTS").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("SO LINESTS").Visible = False
        oSODelvMassUpdateGrid.Columns.Item("SO LINESTS").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("P (cm)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("P (cm)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("L (cm)").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("L (cm)").TitleObject.Sortable = True
        oSODelvMassUpdateGrid.Columns.Item("PxL m2").Editable = False
        oSODelvMassUpdateGrid.Columns.Item("PxL m2").TitleObject.Sortable = True


        'oColumn = oSODelvMassUpdateGrid.Columns.Item("Customer Name")
        'oColumn.Editable = False

        'oSODelvMassUpdateGrid.Columns.Item("SO Date").Width = 80
        'oSODelvMassUpdateGrid.Columns.Item("SO Date").TitleObject.Sortable = True


        oSODelvMassUpdateGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

        ' Set Total Row count in colum title/header
        'oSODelvMassUpdateGrid.Columns.Item(0).TitleObject.Caption = oSODelvMassUpdateGrid.Rows.Count.ToString


        If oForm.DataSources.DataTables.Item(0).Rows.Count <> 0 _
        And oSODelvMassUpdateGrid.DataTable.GetValue(0, 0) <> "" Then
            oForm.Items.Item("cmdUpdSO").Enabled = True
        Else
            oForm.Items.Item("cmdUpdSO").Enabled = False
        End If


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


        Dim IntervalIndex As Integer = 0

        'JIKA INTERVAL DAYS=14, DAN RANGE TGL: 2011-08-01 S/D 2011-08-18
        'RESULT ADALAH:
        'BENTUK KOLOM DGN JUDUL MULAI DARI 0, 1,2,...14, 14+1 
        'MISALKAN: 2011-07-31, 2011-08-01, 2011-08-02, ... 2011-08-15
        'HASILNYA: [< 08/01], [08/01], [08/02], [08/03], ... [08/14], [> 08/14]


        For idx = 0 To colcount - 1 'CInt(oForm.Items.Item("IntervlDay").Specific.string) + 1
            If Left(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 1) = Left(Now.Year.ToString, 1) Then
                'oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = _
                '    Mid(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                '    Right(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 2)
                'oSODelvMassUpdateGrid.Columns.Item(idx).Width = 40
                'oSODelvMassUpdateGrid.Columns.Item(idx).Editable = False
                'oSODelvMassUpdateGrid.Columns.Item(idx).RightJustified = True

                IntervalIndex += 1
                If IntervalIndex = 1 Then
                    oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = "<" + _
                        Mid(oSODelvMassUpdateGrid.Columns.Item(idx + 1).TitleObject.Caption, 6, 2) + "/" + _
                        Right(oSODelvMassUpdateGrid.Columns.Item(idx + 1).TitleObject.Caption, 2)
                    oSODelvMassUpdateGrid.Columns.Item(idx).Width = 50
                ElseIf IntervalIndex = _
                        CInt(IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, _
                                 oForm.Items.Item("IntervlDay").Specific.string)) + 2 Then
                    oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = ">" + _
                        Mid(oSODelvMassUpdateGrid.Columns.Item(idx - 1).TitleObject.Caption, 1, 2) + "/" + _
                        Right(oSODelvMassUpdateGrid.Columns.Item(idx - 1).TitleObject.Caption, 2)
                    oSODelvMassUpdateGrid.Columns.Item(idx).Width = 60
                Else
                    oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption = _
                        Mid(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 6, 2) + "/" + _
                        Right(oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Caption, 2)
                    oSODelvMassUpdateGrid.Columns.Item(idx).Width = 40
                End If

                oSODelvMassUpdateGrid.Columns.Item(idx).Editable = False
                oSODelvMassUpdateGrid.Columns.Item(idx).RightJustified = True
                oSODelvMassUpdateGrid.Columns.Item(idx).TitleObject.Sortable = True

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



        'oSODelvMassUpdateGrid.RowHeaders.Width = 20
        'oSODelvMassUpdateGrid.Columns.Item("#").Width = 30
        'oSODelvMassUpdateGrid.Columns.Item(1).Width = 20
        'oSODelvMassUpdateGrid.Columns.Item("SO Date").Width = 60
        'oSODelvMassUpdateGrid.Columns.Item("DocEntry").Width = 60
        'oSODelvMassUpdateGrid.Columns.Item("DocNum").Width = 60
        'oSODelvMassUpdateGrid.Columns.Item("Line").Width = 30
        'oSODelvMassUpdateGrid.Columns.Item("Cust. Code").Width = 80
        'oSODelvMassUpdateGrid.Columns.Item("FG").Width = 100
        'oSODelvMassUpdateGrid.Columns.Item("Exp Delivery Date").Width = 80
        'oSODelvMassUpdateGrid.Columns.Item("WhsCode").Width = 50
        'oSODelvMassUpdateGrid.Columns.Item("PanjangInCm").Width = 50
        'oSODelvMassUpdateGrid.Columns.Item("LebarInCm").Width = 50
        'oSODelvMassUpdateGrid.Columns.Item("SO_Bentuk").Width = 80


        'Dim oItem As SAPbouiCOM.Item

        ' ''oItem = oForm.Items.Item("OptimMtx").Specific
        'oItem = oForm.Items.Item("PdODueGrid")
        ' ''oItem.Height = 200
        ''oItem.Top = 135
        'oItem.Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        'oItem.Width = oForm.ClientWidth - 20

        oForm.Items.Item("SODelvGrid").Height = oForm.ClientHeight - (oForm.ClientHeight / 4) ' 200
        oForm.Items.Item("SODelvGrid").Width = oForm.ClientWidth - 20



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
        oForm.Items.Item("cmdUpdSO").Left = oForm.ClientWidth - 320

        oForm.Items.Item("NewDueDate").Left = oForm.ClientWidth - 210
        oForm.Items.Item("NewReason").Left = oForm.ClientWidth - 210

        oForm.Items.Item("AsOf").Left = oForm.ClientWidth / 3
        oForm.Items.Item("ViewPeriod").Left = oForm.ClientWidth / 3

        oForm.Items.Item("optSOFull").Left = oForm.ClientWidth / 3
        oForm.Items.Item("optSOHalf").Left = oForm.ClientWidth / 3

        oForm.Items.Item("AsOfDate").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("IntervlDay").Left = oForm.ClientWidth / 3 + 120
        oForm.Items.Item("Days").Left = oForm.ClientWidth / 3 + 180


        Dim sboDate As String
        Dim dDate As DateTime


        sboDate = MIS_Toolkit.fctFormatDate(dDate, oCompany)

        oForm.Freeze(False)

        'MsgBox(GC.GetTotalMemory(True))

        oColumn = Nothing
        oSODelvMassUpdateGrid = Nothing
        GC.Collect()
        'MsgBox(GC.GetTotalMemory(True))

    End Sub

    Function ValidateInputDueDate_Form_DelivPlanDateMassUpdation(ByVal oForm As SAPbouiCOM.Form)

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

    Sub LoadSODelivPlanDate(ByVal oForm As SAPbouiCOM.Form)
        Dim STMTQuery As String

        If oForm.Items.Item("BPCardCode").Specific.string = "" Then
            'oApp.SetStatusBarMessage("Customer must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("SoNumber").Specific.value = "" Then
            'oApp.SetStatusBarMessage("So Number Must Be Fill!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            'Exit Sub
        End If

        If oForm.Items.Item("SODueTo").Specific.string = "" Then
            oForm.Items.Item("SODueTo").Specific.string = oForm.Items.Item("SODueFrom").Specific.string
        End If



        'STMTQuery = "EXEC GetPdODueDateList_ByRangeofPdoDueDateSONumCustomer " + _
        STMTQuery = "EXEC GetSODueDateList_MassUpdateDelivPlan_BySODueDateSONumCustomer " + _
            IIf(oForm.Items.Item("IntervlDay").Specific.string = "", 0, oForm.Items.Item("IntervlDay").Specific.string) _
            + ", '" + Format(CDate(oForm.Items.Item("AsOfDate").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("SODueFrom").Specific.string), "yyyyMMdd") _
            + "', '" + Format(CDate(oForm.Items.Item("SODueTo").Specific.string), "yyyyMMdd") _
            + "', " + IIf(oForm.Items.Item("SoNumber").Specific.string = "", "0", oForm.Items.Item("SoNumber").Specific.string) _
            + ", '" + IIf(oForm.Items.Item("BPCardCode").Specific.string = "", "", oForm.Items.Item("BPCardCode").Specific.string) _
            + "', " + IIf(oForm.DataSources.UserDataSources.Item("SOFull").Value = "", "0", oForm.DataSources.UserDataSources.Item("SOFull").Value)
        '+ "', " + IIf(oForm.Items.Item("optSOHalf").Specific.string = "", "0", oForm.Items.Item("optSOHalf").Specific.string)

        'STMTQuery = "SELECT convert(nvarchar(10), ROW_NUMBER() OVER(ORDER BY T0.DOCDATE)) #, "

        '& " Where T0.DocDate >= '" & Format(CDate(oForm.Items.Item("SODateFrom").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T0.DocDate <= '" & Format(CDate(oForm.Items.Item("SODateTo").Specific.string), "yyyyMMdd") & "' " _
        '    & " AND T1.LineStatus = 'O' " _

        oForm.DataSources.DataTables.Item(0).ExecuteQuery(STMTQuery)

        RearrangeSODelivGrid(oForm)

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

    Sub GenerateDelivPlanDateMassUpdate(ByVal oForm As SAPbouiCOM.Form)
        'On Error GoTo errHandler

        Dim oSODelvMassUpdateGrid As SAPbouiCOM.Grid

        Dim idx As Long

        oSODelvMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

        'GRID - Order by column checkbox
        oSODelvMassUpdateGrid.Columns.Item("ReviseSODue").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)

        'Loop only selected/checked in grid rows and exit.
        For idx = oSODelvMassUpdateGrid.Rows.Count - 1 To 0 Step -1

            oApp.SetStatusBarMessage("Revising Due Date SO.... Start !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            If UCase(oSODelvMassUpdateGrid.DataTable.GetValue(0, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx))) = "Y" Then
                'If UCase(oSODelvMassUpdateGrid.DataTable.GetValue(0, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx))) = "Y" _
                '    And oSODelvMassUpdateGrid.DataTable.GetValue(1, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "61061" Then

                Dim oSalesOrder As SAPbobsCOM.Documents = Nothing
                Dim oSalesOrderLines As SAPbobsCOM.Document_Lines

                'Dim oProd1 As SAPbobsCOM.ProductionOrders = Nothing
                'Dim oProdLine1 As SAPbobsCOM.ProductionOrders_Lines



                'Dim isconnect As Long
                Dim errConnect As String = ""


                If Not oCompany.InTransaction Then
                    oCompany.StartTransaction()
                End If


                Dim PdOno As String = ""

                oSalesOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                'oPDOStatus.GetByKey(oSODelvMassUpdateGrid.DataTable.GetValue(2, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString))

                oSalesOrder.GetByKey(oSODelvMassUpdateGrid.DataTable.GetValue(4, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString))
                'oSalesOrder.GetByKey(11627)

                oSalesOrderLines = oSalesOrder.Lines

                'Dim SOLINENUM As Integer
                'SOLINENUM = CInt(oSODelvMassUpdateGrid.DataTable.GetValue(3, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString))

                'oSalesOrderLines.SetCurrentLine(IIf(oSODelvMassUpdateGrid.DataTable.GetValue(3, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString) = "", 0, _
                '                                    oSODelvMassUpdateGrid.DataTable.GetValue(3, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString)))
                oSalesOrderLines.SetCurrentLine(oSODelvMassUpdateGrid.DataTable.GetValue(5, oSODelvMassUpdateGrid.GetDataTableRowIndex(idx).ToString))

                'oSalesOrderLines.SetCurrentLine(2)

                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivPlanDt03").Value = _
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivPlanDt02").Value
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReas03").Value = _
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReas02").Value
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivPlanDt02").Value = _
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivPlanDt01").Value
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReas02").Value = _
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReas01").Value

                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivPlanDt01").Value = _
                oSalesOrderLines.ShipDate
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReas01").Value = _
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReason").Value

                oSalesOrderLines.ShipDate = CDate(oForm.Items.Item("NewDueDate").Specific.string)
                oSalesOrderLines.UserFields.Fields.Item("U_MIS_DelivReason").Value = _
                oForm.Items.Item("NewReason").Specific.string

                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate03").Value = oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate02").Value
                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReas03").Value = oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReas02").Value
                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate02").Value = oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value
                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReas02").Value = oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReas01").Value

                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value = oSalesOrder.DueDate
                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReas01").Value = oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReason").Value

                ''oSalesOrder.UserFields.Fields.Item("U_MIS_DueDate01").Value = oSalesOrder.DueDate
                'oSalesOrder.DueDate = CDate(oForm.Items.Item("NewDueDate").Specific.string)
                'oSalesOrder.UserFields.Fields.Item("U_MIS_DueDateReason").Value = oForm.Items.Item("NewReason").Specific.string

                lRetCode = oSalesOrder.Update()


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


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrderLines)
                    oSalesOrderLines = Nothing

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrder)
                    oSalesOrder = Nothing


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
        oApp.SetStatusBarMessage("Revising SO.... Finished !!! " & idx + 1 & " ", SAPbouiCOM.BoMessageTime.bmt_Short, False)


        'MsgBox("generating... PdO; DONE!!!")

        Exit Sub


errHandler:
        MsgBox("Exception: " & Err.Description)
        Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    End Sub


End Module
