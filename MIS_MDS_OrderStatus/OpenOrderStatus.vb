Module OpenOrderStatus

    Sub OpenOrderStatus_FormEntry()
        Dim oForm As SAPbouiCOM.Form
        Dim STMTQuery As String
        Dim oItem As SAPbouiCOM.Item
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button
        Dim oSOGrid As SAPbouiCOM.Grid

        Try
            oForm = oApp.Forms.Item("mds_ord4")
            oApp.MessageBox("Form Already Open")

        Catch ex As Exception
            Dim fcp As SAPbouiCOM.FormCreationParams
            fcp = oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            fcp.FormType = "mds_ord4"
            fcp.UniqueID = "mds_ord4"

            fcp.XmlData = LoadFromXML("OpenOrderStatus.srf")
            oForm = oApp.Forms.AddEx(fcp)
            oForm.Freeze(True)

            ' Add User DataSource
            ' not binding to SBO data or UDO/UDF
            oForm.DataSources.UserDataSources.Add("CustGroup", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("Customer", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            oForm.DataSources.UserDataSources.Add("SONo", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            'Default value 
            oForm.DataSources.UserDataSources.Item("SONo").Value = "1"

            oEditText = oForm.Items.Item("CustGroup").Specific
            oEditText.DataBind.SetBound(True, "", "CustGroup")

            oEditText = oForm.Items.Item("Customer").Specific
            oEditText.DataBind.SetBound(True, "", "Customer")

            oEditText = oForm.Items.Item("SONo").Specific
            oEditText.DataBind.SetBound(True, "", "SONo")

            oItem = oForm.Items.Item("SOGrid")
            oItem.Left = 5
            oItem.Top = 90
            oItem.Width = oForm.ClientWidth - 10
            oItem.Height = oForm.ClientHeight - 200

            oSOGrid = oItem.Specific
            oSOGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            oForm.Freeze(False)

            STMTQuery = "EXEC MIS_OpenOrderStatus"

            ' Grid #: 1
            oForm.DataSources.DataTables.Add("DOList")
            oForm.DataSources.DataTables.Item("DOList").ExecuteQuery(STMTQuery)
            oSOGrid.DataTable = oForm.DataSources.DataTables.Item("DOList")

            oEditText = Nothing
            oItem = Nothing
            oButton = Nothing
            oSOGrid = Nothing


            GC.Collect()
            'MsgBox(GC.GetTotalMemory(True))

        End Try
    End Sub

End Module
