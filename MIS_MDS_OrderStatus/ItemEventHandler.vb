Module ItemEventHandler

    Public WithEvents oApp4ItemEvent As SAPbouiCOM.Application = Nothing

    'Const ProductionIssue_MenuId As String = "4371"
    Const ProductionIssue_FormId As String = "65213"
    Const ProductionIssueUDF_FormId As String = "-65213"
    Dim objFormProductionIssue As SAPbouiCOM.Form
    Dim objFormProductionIssueUDF As SAPbouiCOM.Form
    'Dim intRowProductionIssueDetail As Integer
    ''karno 
    '' Production Issue
    'Const Production_MenuId As String = "4369"
    Const Production_FormId As String = "65211"
    Const ProductionUDF_FormId As String = "-65211"
    Dim objFormProduction As SAPbouiCOM.Form
    Dim objFormProductionUDF As SAPbouiCOM.Form
    'Dim intRowProductionDetail As Integer


    Sub ItemEventHandler(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                     ByRef BubbleEvent As Boolean) Handles oApp4ItemEvent.ItemEvent
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try

            If pVal.BeforeAction = False Then

                ''karno Copy Optim
                'If pVal.FormTypeEx = ProductionIssueUDF_FormId Then
                '    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                '        'objFormProductionIssueUDF = SBO_Application.Forms.Item(pVal.FormUID)
                '        objFormProductionIssueUDF = oApp.Forms.Item(pVal.FormUID)
                '    End If
                'End If

                Select Case FormUID

                    Case "mds_ord1" '"mds_p1"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            'oItemMat = oForm.Items.Item("matrixName")
                            'oItemMat.Width = oForm.Width - 200
                            If oForm IsNot Nothing Then
                                RearrangePdoDueGrid(oForm)
                            End If
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects

                                Dim xval As String


                                xval = oDataTable.GetValue(0, 0)

                                If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                                    oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                                End If

                                oCFL = Nothing
                                oDataTable = Nothing
                            End If

                            'oForm = Nothing
                            'oCFLEvento = Nothing
                            'GC.Collect()

                        End If


                        ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadPdO") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)


                            LoadPdO(oForm)


                        End If

                        If (pVal.ItemUID = "PdODueFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'Dim vdate As MISToolbox
                            'vdate = New MISToolbox
                            'Dim validDate As Boolean


                            If Len(oForm.Items.Item("PdODueFrom").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("PdO Due Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'validDate = vdate.SBODateisValid("2010918")

                            'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                            'If validDate = False Then
                            '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If

                            If Len(oForm.Items.Item("PdODueFrom").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("PdO Due Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("PdODueFrom").Specific.string) = 8 Then
                                oForm.Items.Item("PdODueFrom").Specific.string = _
                                    CDate(Left(oForm.Items.Item("PdODueFrom").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("PdODueFrom").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("PdODueFrom").Specific.string, 2))
                            End If

                            If oForm.Items.Item("PdODueFrom").Specific.string = "" Then
                                oForm.Items.Item("PdODueFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                            End If

                            If oForm.Items.Item("PdODueTo").Specific.string = "" Then
                                oForm.Items.Item("PdODueTo").Specific.string = oForm.Items.Item("PdODueFrom").Specific.string
                            End If

                            'vdate = Nothing

                            'oForm.Items.Item("SODateFrom").Click()
                            '                        BubbleEvent = False
                        End If

                        If pVal.ItemUID = "PdODueTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("PdODueTo").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("PdO Due Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                'Exit Sub
                            End If

                            If Len(oForm.Items.Item("PdODueTo").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("PdO Due Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                'BubbleEvent = False
                            End If

                            If Len(oForm.Items.Item("PdODueTo").Specific.string) = 8 Then
                                oForm.Items.Item("PdODueTo").Specific.string = _
                                    CDate(Left(oForm.Items.Item("PdODueTo").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("PdODueTo").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("PdODueTo").Specific.string, 2))
                            End If
                            'BubbleEvent = True
                            'oForm.Items.Item("SODateTo").Click('')

                            'oForm = Nothing
                            'GC.Collect()

                        End If

                        If pVal.ItemUID = "TogglChkBx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            'oForm.Freeze(True)

                            oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

                            If oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").Editable = True Then
                                'oForm.Items.Item("TogglChkBx").Enabled = False
                                oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").Editable = False
                            Else
                                'oForm.Items.Item("TogglChkBx").Enabled = True
                                oPdODueMassUpdateGrid.Columns.Item("RevisePdODue").Editable = True
                            End If






                            'GeneratePdODueDateMassUpdate(oForm)

                            'LoadPdO(oForm)


                        End If

                        If pVal.ItemUID = "cmdUpdPdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            'Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("PdODueList")

                            'oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

                            'get total row count selected
                            'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                            'selection rows -> e.g: user select row# by order respectively: 1, 3, 2, 5

                            'get row index of selected grid, has two method:
                            'method# 1: ot_RowOrder (value=1)
                            'result row selected: 1, 2, 3, 5

                            'method# 2: ot_SelectionOrder (value=0)
                            'result row selected: 1, 3, 2, 5

                            'For idx = 0 To oPdODueMassUpdateGrid.Rows.SelectedRows.Count - 1
                            '    MsgBox("selected row#:" & idx.ToString & _
                            '           "; selectedrow->row#: " & oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
                            '           & "docnum: " & oPdODueMassUpdateGrid.DataTable.GetValue(0, oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                            'Next

                            'Dim oPdO As SAPbobsCOM.ProductionOrders
                            'oPdO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                            ''Fill PdO properties...
                            'oPdO.ItemNo = "LM4029"
                            ''oPdO.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(13, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
                            'oPdO.DueDate = DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                            'oPdO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                            'oPdO.PlannedQuantity = 188
                            'oPdO.PostingDate = DateTime.Today 'DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.Add()

                            '???
                            Dim isValid As Boolean

                            isValid = ValidateInputDueDate_Form_PdODueDateMassUpdation(oForm)
                            If isValid = True Then
                                GeneratePdODueDateMassUpdate(oForm)
                            End If


                            LoadPdO(oForm)


                        End If

                        'toggle select/unselect all
                        If pVal.ColUID = "RevisePdODue" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("PdODueList")

                            oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

                            'get total row count selected
                            'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                            oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

                            'If oPdODueMassUpdateGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                            If oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("RevisePdODue", idx, "Y")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("RevisePdODue", idx, "N")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If

                    Case "mds_ord2" '"mds_p2"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            'oItemMat = oForm.Items.Item("matrixName")
                            'oItemMat.Width = oForm.Width - 200
                            If oForm IsNot Nothing Then
                                RearrangeSODelivGrid(oForm)
                            End If
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects

                                Dim xval As String


                                xval = oDataTable.GetValue(0, 0)

                                If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                                    oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                                End If

                                oCFL = Nothing
                                oDataTable = Nothing
                            End If

                            'oForm = Nothing
                            'oCFLEvento = Nothing
                            'GC.Collect()

                        End If


                        ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadSO") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)


                            LoadSODelivPlanDate(oForm)


                        End If

                        If (pVal.ItemUID = "SODueFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'Dim vdate As MISToolbox
                            'vdate = New MISToolbox
                            'Dim validDate As Boolean


                            If Len(oForm.Items.Item("SODueFrom").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'validDate = vdate.SBODateisValid("2010918")

                            'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                            'If validDate = False Then
                            '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If

                            If Len(oForm.Items.Item("SODueFrom").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODueFrom").Specific.string) = 8 Then
                                oForm.Items.Item("SODueFrom").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODueFrom").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODueFrom").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODueFrom").Specific.string, 2))
                            End If

                            If oForm.Items.Item("SODueFrom").Specific.string = "" Then
                                oForm.Items.Item("SODueFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                            End If

                            If oForm.Items.Item("SODueTo").Specific.string = "" Then
                                oForm.Items.Item("SODueTo").Specific.string = oForm.Items.Item("SODueFrom").Specific.string
                            End If

                            'vdate = Nothing

                            'oForm.Items.Item("SODateFrom").Click()
                            '                        BubbleEvent = False
                        End If

                        If pVal.ItemUID = "SODueTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("SODueTo").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                'Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODueTo").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                'BubbleEvent = False
                            End If

                            If Len(oForm.Items.Item("SODueTo").Specific.string) = 8 Then
                                oForm.Items.Item("SODueTo").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODueTo").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODueTo").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODueTo").Specific.string, 2))
                            End If
                            'BubbleEvent = True
                            'oForm.Items.Item("SODateTo").Click('')

                            'oForm = Nothing
                            'GC.Collect()

                        End If

                        If pVal.ItemUID = "TogglChkBx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            'oForm.Freeze(True)

                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            If oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = True Then
                                'oForm.Items.Item("TogglChkBx").Enabled = False
                                oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = False
                            Else
                                'oForm.Items.Item("TogglChkBx").Enabled = True
                                oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = True
                            End If






                            'GeneratePdODueDateMassUpdate(oForm)

                            'LoadPdO(oForm)


                        End If

                        If pVal.ItemUID = "cmdUpdSO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            'Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SODelvList")

                            'oPdODueMassUpdateGrid = oForm.Items.Item("PdODueGrid").Specific

                            'get total row count selected
                            'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                            'selection rows -> e.g: user select row# by order respectively: 1, 3, 2, 5

                            'get row index of selected grid, has two method:
                            'method# 1: ot_RowOrder (value=1)
                            'result row selected: 1, 2, 3, 5

                            'method# 2: ot_SelectionOrder (value=0)
                            'result row selected: 1, 3, 2, 5

                            'For idx = 0 To oPdODueMassUpdateGrid.Rows.SelectedRows.Count - 1
                            '    MsgBox("selected row#:" & idx.ToString & _
                            '           "; selectedrow->row#: " & oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
                            '           & "docnum: " & oPdODueMassUpdateGrid.DataTable.GetValue(0, oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                            'Next

                            'Dim oPdO As SAPbobsCOM.ProductionOrders
                            'oPdO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                            ''Fill PdO properties...
                            'oPdO.ItemNo = "LM4029"
                            ''oPdO.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(13, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
                            'oPdO.DueDate = DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                            'oPdO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                            'oPdO.PlannedQuantity = 188
                            'oPdO.PostingDate = DateTime.Today 'DateTime.Today.ToString("yyyyMMdd")
                            'oPdO.Add()

                            '???
                            Dim isValid As Boolean

                            isValid = ValidateInputDueDate_Form_DelivPlanDateMassUpdation(oForm)
                            If isValid = True Then
                                'GeneratePdODueDateMassUpdate(oForm)
                                GenerateDelivPlanDateMassUpdate(oForm)
                            End If


                            LoadSODelivPlanDate(oForm)


                        End If

                        'toggle select/unselect all
                        If pVal.ColUID = "ReviseSODue" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SODelvList")

                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            'get total row count selected
                            'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            'If oPdODueMassUpdateGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                            If oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("ReviseSODue", idx, "Y")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("ReviseSODue", idx, "N")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If



                        'Case "mds_ord3" '"mds_p11"  versi ada mesin

                        '    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                        '        oForm = oApp.Forms.Item(FormUID)

                        '        'oItemMat = oForm.Items.Item("matrixName")
                        '        'oItemMat.Width = oForm.Width - 200
                        '        If oForm IsNot Nothing Then
                        '            RearrangePdoDueGridByMachine(oForm)
                        '        End If
                        '        DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))

                        '    End If

                        '    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        '        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        '        oCFLEvento = pVal

                        '        Dim sCFL_ID As String
                        '        sCFL_ID = oCFLEvento.ChooseFromListUID

                        '        oForm = oApp.Forms.Item(FormUID)

                        '        Dim oCFL As SAPbouiCOM.ChooseFromList
                        '        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                        '        If oCFLEvento.BeforeAction = False Then
                        '            Dim oDataTable As SAPbouiCOM.DataTable
                        '            oDataTable = oCFLEvento.SelectedObjects

                        '            Dim xval As String


                        '            xval = oDataTable.GetValue(0, 0)

                        '            If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                        '                oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                        '            End If

                        '            oCFL = Nothing
                        '            oDataTable = Nothing
                        '        End If

                        '        'oForm = Nothing
                        '        'oCFLEvento = Nothing
                        '        'GC.Collect()

                        '    End If


                        '    ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        '    ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        '    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadPdO") Then
                        '        'Dim oForm As SAPbouiCOM.Form
                        '        'oForm = SBO_Application.Forms.Item(FormUID)
                        '        oForm = oApp.Forms.Item(FormUID)


                        '        LoadPdODueDateByMachine(oForm)


                        '    End If

                        '    If (pVal.ItemUID = "PdODueFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                        '        'Dim oForm As SAPbouiCOM.Form
                        '        'oForm = SBO_Application.Forms.Item(FormUID)
                        '        oForm = oApp.Forms.Item(FormUID)

                        '        'Dim vdate As MISToolbox
                        '        'vdate = New MISToolbox
                        '        'Dim validDate As Boolean


                        '        If Len(oForm.Items.Item("PdODueFrom").Specific.string) = 0 Then
                        '            'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            oApp.SetStatusBarMessage("PdO Due Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            Exit Sub
                        '        End If

                        '        'validDate = vdate.SBODateisValid("2010918")

                        '        'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                        '        'If validDate = False Then
                        '        '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '        '    BubbleEvent = False
                        '        '    Exit Sub
                        '        'End If

                        '        If Len(oForm.Items.Item("PdODueFrom").Specific.string) < 8 Then
                        '            'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            oApp.SetStatusBarMessage("PdO Due Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            Exit Sub
                        '        End If

                        '        If Len(oForm.Items.Item("PdODueFrom").Specific.string) = 8 Then
                        '            oForm.Items.Item("PdODueFrom").Specific.string = _
                        '                CDate(Left(oForm.Items.Item("PdODueFrom").Specific.string, 4) & "/" & _
                        '                    Mid(oForm.Items.Item("PdODueFrom").Specific.string, 5, 2) & "/" & _
                        '                    Right(oForm.Items.Item("PdODueFrom").Specific.string, 2))
                        '        End If

                        '        If oForm.Items.Item("PdODueFrom").Specific.string = "" Then
                        '            oForm.Items.Item("PdODueFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                        '        End If

                        '        If oForm.Items.Item("PdODueTo").Specific.string = "" Then
                        '            oForm.Items.Item("PdODueTo").Specific.string = oForm.Items.Item("PdODueFrom").Specific.string
                        '        End If

                        '        'vdate = Nothing

                        '        'oForm.Items.Item("SODateFrom").Click()
                        '        '                        BubbleEvent = False
                        '    End If

                        '    If pVal.ItemUID = "PdODueTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                        '        'Dim oForm As SAPbouiCOM.Form
                        '        'oForm = SBO_Application.Forms.Item(FormUID)
                        '        oForm = oApp.Forms.Item(FormUID)

                        '        If Len(oForm.Items.Item("PdODueTo").Specific.string) = 0 Then
                        '            'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            oApp.SetStatusBarMessage("PdO Due Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            BubbleEvent = False
                        '            'Exit Sub
                        '        End If

                        '        If Len(oForm.Items.Item("PdODueTo").Specific.string) < 8 Then
                        '            'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            oApp.SetStatusBarMessage("PdO Due Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            'BubbleEvent = False
                        '        End If

                        '        If Len(oForm.Items.Item("PdODueTo").Specific.string) = 8 Then
                        '            oForm.Items.Item("PdODueTo").Specific.string = _
                        '                CDate(Left(oForm.Items.Item("PdODueTo").Specific.string, 4) & "/" & _
                        '                    Mid(oForm.Items.Item("PdODueTo").Specific.string, 5, 2) & "/" & _
                        '                    Right(oForm.Items.Item("PdODueTo").Specific.string, 2))
                        '        End If
                        '        'BubbleEvent = True
                        '        'oForm.Items.Item("SODateTo").Click('')

                        '        'oForm = Nothing
                        '        'GC.Collect()

                        '    End If

                        '    If pVal.ItemUID = "TogglChkBx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        '        oForm = oApp.Forms.Item(FormUID)

                        '        Dim oPdODueMassUpdateByMachineGrid As SAPbouiCOM.Grid

                        '        'oForm.Freeze(True)

                        '        oPdODueMassUpdateByMachineGrid = oForm.Items.Item("PdOMchGrid").Specific

                        '        If oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").Editable = True Then
                        '            'oForm.Items.Item("TogglChkBx").Enabled = False
                        '            oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").Editable = False
                        '        Else
                        '            'oForm.Items.Item("TogglChkBx").Enabled = True
                        '            oPdODueMassUpdateByMachineGrid.Columns.Item("RevisePdODue").Editable = True
                        '        End If






                        '        'GeneratePdODueDateMassUpdate(oForm)

                        '        'LoadPdO(oForm)


                        '    End If

                        '    If pVal.ItemUID = "cmdUpdPdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        '        'Dim oForm As SAPbouiCOM.Form
                        '        'oForm = SBO_Application.Forms.Item(FormUID)
                        '        oForm = oApp.Forms.Item(FormUID)
                        '        'Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                        '        Dim dt As SAPbouiCOM.DataTable

                        '        dt = oForm.DataSources.DataTables.Item("PdODueListByMachine")

                        '        'oPdODueMassUpdateGrid = oForm.Items.Item("PdOMchGrid").Specific

                        '        'get total row count selected
                        '        'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                        '        'selection rows -> e.g: user select row# by order respectively: 1, 3, 2, 5

                        '        'get row index of selected grid, has two method:
                        '        'method# 1: ot_RowOrder (value=1)
                        '        'result row selected: 1, 2, 3, 5

                        '        'method# 2: ot_SelectionOrder (value=0)
                        '        'result row selected: 1, 3, 2, 5

                        '        'For idx = 0 To oPdODueMassUpdateGrid.Rows.SelectedRows.Count - 1
                        '        '    MsgBox("selected row#:" & idx.ToString & _
                        '        '           "; selectedrow->row#: " & oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder) _
                        '        '           & "docnum: " & oPdODueMassUpdateGrid.DataTable.GetValue(0, oPdODueMassUpdateGrid.Rows.SelectedRows.Item(idx, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))

                        '        'Next

                        '        'Dim oPdO As SAPbobsCOM.ProductionOrders
                        '        'oPdO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                        '        ''Fill PdO properties...
                        '        'oPdO.ItemNo = "LM4029"
                        '        ''oPdO.DueDate = oPdODueMassUpdateGrid.DataTable.GetValue(13, oPdODueMassUpdateGrid.GetDataTableRowIndex(idx).ToString)
                        '        'oPdO.DueDate = DateTime.Today.ToString("yyyyMMdd")
                        '        'oPdO.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                        '        'oPdO.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                        '        'oPdO.PlannedQuantity = 188
                        '        'oPdO.PostingDate = DateTime.Today 'DateTime.Today.ToString("yyyyMMdd")
                        '        'oPdO.Add()

                        '        '???
                        '        Dim isValid As Boolean

                        '        isValid = ValidateInputDueDate_Form_PdODueDateMassUpdation(oForm)
                        '        If isValid = True Then
                        '            GeneratePdODueDateMassUpdate(oForm)
                        '        End If


                        '        LoadPdO(oForm)


                        '    End If

                        '    'toggle select/unselect all
                        '    If pVal.ColUID = "RevisePdODue" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                        '        'Dim oForm As SAPbouiCOM.Form
                        '        'oForm = SBO_Application.Forms.Item(FormUID)
                        '        oForm = oApp.Forms.Item(FormUID)
                        '        Dim oPdODueMassUpdateByMachineGrid As SAPbouiCOM.Grid

                        '        Dim idx As Long
                        '        Dim dt As SAPbouiCOM.DataTable

                        '        dt = oForm.DataSources.DataTables.Item("PdODueList")

                        '        oPdODueMassUpdateByMachineGrid = oForm.Items.Item("PdODueGrid").Specific

                        '        'get total row count selected
                        '        'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                        '        oPdODueMassUpdateByMachineGrid = oForm.Items.Item("PdODueGrid").Specific

                        '        'If oPdODueMassUpdateGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                        '        If oPdODueMassUpdateByMachineGrid.Columns.Item(0).TitleObject.Caption = "Select All" Then
                        '            'select/check all
                        '            oForm.Freeze(True)

                        '            For idx = 0 To oPdODueMassUpdateByMachineGrid.Rows.Count - 1
                        '                dt.SetValue("RevisePdODue", idx, "Y")
                        '            Next
                        '            oPdODueMassUpdateByMachineGrid.Columns.Item(0).TitleObject.Caption = "Reset All"
                        '            oForm.Freeze(False)
                        '        Else
                        '            'unselect/uncheck all
                        '            oForm.Freeze(True)
                        '            For idx = 0 To oPdODueMassUpdateByMachineGrid.Rows.Count - 1
                        '                dt.SetValue("RevisePdODue", idx, "N")
                        '            Next
                        '            oPdODueMassUpdateByMachineGrid.Columns.Item(0).TitleObject.Caption = "Select All"
                        '            oForm.Freeze(False)
                        '        End If

                        '        'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        '    End If

                    Case "mds_ord3" '"mds_p1"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            If oForm IsNot Nothing Then
                                RearrangeDOGrid(oForm)
                            End If
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))

                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects

                                Dim xval As String

                                xval = oDataTable.GetValue(0, 0)

                                oCFL = Nothing
                                oDataTable = Nothing
                            End If

                        End If

                        ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadDO") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            LoadDO(oForm)

                        End If

                        If (pVal.ItemUID = "DODateFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("DODateFrom").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("Doc Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("DODateFrom").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("Doc Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("DODateFrom").Specific.string) = 8 Then
                                oForm.Items.Item("DODateFrom").Specific.string = _
                                    CDate(Left(oForm.Items.Item("DODateFrom").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("DODateFrom").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("DODateFrom").Specific.string, 2))
                            End If

                            If oForm.Items.Item("DODateFrom").Specific.string = "" Then
                                oForm.Items.Item("DODateFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                            End If

                            If oForm.Items.Item("DODateTo").Specific.string = "" Then
                                oForm.Items.Item("DDOateTo").Specific.string = oForm.Items.Item("DateFrom").Specific.string
                            End If

                        End If

                        If pVal.ItemUID = "DODateTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("DODateTo").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("Doc Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                'Exit Sub
                            End If

                            If Len(oForm.Items.Item("DODateTo").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("Doc Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                'BubbleEvent = False
                            End If

                            If Len(oForm.Items.Item("DODateTo").Specific.string) = 8 Then
                                oForm.Items.Item("DODateTo").Specific.string = _
                                    CDate(Left(oForm.Items.Item("DODateTo").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("DODateTo").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("DODateTo").Specific.string, 2))
                            End If
                        End If

                        If pVal.ItemUID = "Jam" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("Jam").Specific.string) < 4 Or Len(oForm.Items.Item("Jam").Specific.string) > 5 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("Jam Must be entered! Format: HH:mm ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                            Else
                                'oForm.DataSources.UserDataSources.Item("Jam").Value = DateTime.Today.Now.ToString("HH:mm")
                                oForm.DataSources.UserDataSources.Item("Jam").Value = _
                                    Left(oForm.Items.Item("Jam").Specific.string, 2) _
                                    + ":" + Right(oForm.Items.Item("Jam").Specific.string, 2)
                                'Exit Sub
                            End If
                        End If

                        If pVal.ItemUID = "cmdUpdDO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            oForm = oApp.Forms.Item(FormUID)
                            Dim dt As SAPbouiCOM.DataTable
                            dt = oForm.DataSources.DataTables.Item("DOList")

                            Dim isValid As Boolean
                            isValid = ValidateInputDODate_Form_DODateMassUpdation(oForm)
                            If isValid = True Then
                                GenerateDODateUpdate(oForm)
                            End If

                            LoadDO(oForm)

                        End If

                        If pVal.ColUID = "Check" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oDOGrid As SAPbouiCOM.Grid
                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable
                            dt = oForm.DataSources.DataTables.Item("DOList")

                            oDOGrid = oForm.Items.Item("DOGrid").Specific

                            'oDOGrid = oForm.Items.Item("DOGrid").Specific

                            'If oPdODueMassUpdateGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                            If oDOGrid.Columns.Item(0).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oDOGrid.Rows.Count - 1
                                    dt.SetValue("Check", idx, "Y")
                                Next
                                oDOGrid.Columns.Item(0).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oDOGrid.Rows.Count - 1
                                    dt.SetValue("Check", idx, "N")
                                Next
                                oDOGrid.Columns.Item(0).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                        End If

                    Case "mds_ord4" '"mds_p4"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            If oForm IsNot Nothing Then
                                'RearrangeSODelivGrid(oForm)
                                'RearrangeOrderStatusForMKTGrid(oForm)
                                RearrangeOrderStatusGrid(oForm)
                            End If
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                            oCFLEvento = pVal

                            Dim sCFL_ID As String
                            sCFL_ID = oCFLEvento.ChooseFromListUID

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oCFL As SAPbouiCOM.ChooseFromList
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)



                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvento.SelectedObjects

                                Dim xval As String


                                xval = oDataTable.GetValue(0, 0)

                                If pVal.ItemUID = "BPCardCode" Or pVal.ItemUID = "BPButton" Then

                                    oForm.DataSources.UserDataSources.Item("BPDS").ValueEx = xval
                                End If

                                oCFL = Nothing
                                oDataTable = Nothing
                            End If

                            'oForm = Nothing
                            'oCFLEvento = Nothing
                            'GC.Collect()

                        End If


                        ' Button is clicked/pressed, event = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED 
                        ' clicked, event = SAPbouiCOM.BoEventTypes.et_CLICK
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "cmdLoadSO") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)


                            'LoadOrderStatusForMarketing(oForm)
                            LoadOrderStatus(oForm)

                        End If

                        If (pVal.ItemUID = "SODueFrom") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then

                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'Dim vdate As MISToolbox
                            'vdate = New MISToolbox
                            'Dim validDate As Boolean


                            If Len(oForm.Items.Item("SODueFrom").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            'validDate = vdate.SBODateisValid("2010918")

                            'validDate = vdate.SBODateisValid(oForm.Items.Item("SODateFrom").Specific.string)
                            'If validDate = False Then
                            '    SBO_Application.SetStatusBarMessage("SO Date From is invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If

                            If Len(oForm.Items.Item("SODueFrom").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date From invalid!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODueFrom").Specific.string) = 8 Then
                                oForm.Items.Item("SODueFrom").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODueFrom").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODueFrom").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODueFrom").Specific.string, 2))
                            End If

                            If oForm.Items.Item("SODueFrom").Specific.string = "" Then
                                oForm.Items.Item("SODueFrom").Specific.string = Format(Today, "yyyyMMdd") ' "20100929"
                            End If

                            If oForm.Items.Item("SODueTo").Specific.string = "" Then
                                oForm.Items.Item("SODueTo").Specific.string = oForm.Items.Item("SODueFrom").Specific.string
                            End If

                        End If

                        If pVal.ItemUID = "SODueTo" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            If Len(oForm.Items.Item("SODueTo").Specific.string) = 0 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date To Must be entered!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                BubbleEvent = False
                                'Exit Sub
                            End If

                            If Len(oForm.Items.Item("SODueTo").Specific.string) < 8 Then
                                'SBO_Application.SetStatusBarMessage("SO Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                oApp.SetStatusBarMessage("SO Due Date To is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                'BubbleEvent = False
                            End If

                            If Len(oForm.Items.Item("SODueTo").Specific.string) = 8 Then
                                oForm.Items.Item("SODueTo").Specific.string = _
                                    CDate(Left(oForm.Items.Item("SODueTo").Specific.string, 4) & "/" & _
                                        Mid(oForm.Items.Item("SODueTo").Specific.string, 5, 2) & "/" & _
                                        Right(oForm.Items.Item("SODueTo").Specific.string, 2))
                            End If
                            'BubbleEvent = True
                            'oForm.Items.Item("SODateTo").Click('')

                            'oForm = Nothing
                            'GC.Collect()

                        End If

                        If pVal.ItemUID = "TogglChkBx" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            oForm = oApp.Forms.Item(FormUID)

                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            'oForm.Freeze(True)

                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            If oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = True Then
                                'oForm.Items.Item("TogglChkBx").Enabled = False
                                oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = False
                            Else
                                'oForm.Items.Item("TogglChkBx").Enabled = True
                                oPdODueMassUpdateGrid.Columns.Item("ReviseSODue").Editable = True
                            End If






                            'GeneratePdODueDateMassUpdate(oForm)

                            'LoadPdO(oForm)


                        End If

                        If pVal.ItemUID = "cmdUpdSO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            oForm = oApp.Forms.Item(FormUID)

                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SODelvList")


                            '???
                            Dim isValid As Boolean

                            isValid = ValidateInputDueDate_Form_DelivPlanDateMassUpdation(oForm)
                            If isValid = True Then
                                'GeneratePdODueDateMassUpdate(oForm)
                                GenerateDelivPlanDateMassUpdate(oForm)
                            End If


                            LoadSODelivPlanDate(oForm)


                        End If

                        'toggle select/unselect all
                        If pVal.ColUID = "ReviseSODue" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oPdODueMassUpdateGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            dt = oForm.DataSources.DataTables.Item("SODelvList")

                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            'get total row count selected
                            'oPdODueMassUpdateGrid.Rows.SelectedRows.Count.ToString()


                            oPdODueMassUpdateGrid = oForm.Items.Item("SODelvGrid").Specific

                            'If oPdODueMassUpdateGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                            If oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("ReviseSODue", idx, "Y")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oPdODueMassUpdateGrid.Rows.Count - 1
                                    dt.SetValue("ReviseSODue", idx, "N")
                                Next
                                oPdODueMassUpdateGrid.Columns.Item(0).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If



                    Case "mds_p3"
                        'If pVal.Before_Action = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE Then 'And oForm IsNot (Nothing) Then
                            oForm = oApp.Forms.Item(FormUID)

                            'oItemMat = oForm.Items.Item("matrixName")
                            'oItemMat.Width = oForm.Width - 200
                            'RearrangeFormOptimEntry(oForm)
                            DIErrHandler("Form Height: " + CStr(oForm.ClientHeight) + ", Width: " + CStr(oForm.ClientWidth))
                        End If
                        'End If


                        '-------------- Yadi FC ----------------------------
                    Case "MDS_P6"
                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnShow") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            'LoadProductionClosed(oForm)
                        End If


                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnCancel") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            oForm.Close()
                        End If

                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "btnUpdate") Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)

                            'UpdateProductionClosed(oForm)
                            'LoadProductionClosed(oForm)
                        End If

                        '-------------- Yadi FC ----------------------------
                        'toggle select/unselect all
                        'If pVal.ColUID = "Release PdO" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                        If pVal.ColUID = "Check" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row = -1 Then
                            'Dim oForm As SAPbouiCOM.Form
                            'oForm = SBO_Application.Forms.Item(FormUID)
                            oForm = oApp.Forms.Item(FormUID)
                            Dim oProdClosedGrid As SAPbouiCOM.Grid

                            Dim idx As Long
                            Dim dt As SAPbouiCOM.DataTable

                            'Dim str As String

                            dt = oForm.DataSources.DataTables.Item("PCLst")

                            oProdClosedGrid = oForm.Items.Item("grdPC").Specific

                            'get total row count selected
                            'oProdClosedGrid.Rows.SelectedRows.Count.ToString()


                            oProdClosedGrid = oForm.Items.Item("grdPC").Specific

                            If oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Select All" Then
                                'select/check all
                                oForm.Freeze(True)

                                For idx = 0 To oProdClosedGrid.Rows.Count - 1
                                    If oProdClosedGrid.DataTable.GetValue(6, oProdClosedGrid.GetDataTableRowIndex(idx)) = _
                                        oProdClosedGrid.DataTable.GetValue(7, oProdClosedGrid.GetDataTableRowIndex(idx)) Then
                                        dt.SetValue("Check", idx, "Y")
                                    End If
                                    'str = oProdClosedGrid.Columns.Item(idx).Description
                                    'str = oProdClosedGrid.DataTable.GetValue(1, oProdClosedGrid.GetDataTableRowIndex(idx))
                                Next
                                oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Reset All"
                                oForm.Freeze(False)
                            Else
                                'unselect/uncheck all
                                oForm.Freeze(True)
                                For idx = 0 To oProdClosedGrid.Rows.Count - 1
                                    dt.SetValue("Check", idx, "N")
                                Next
                                oProdClosedGrid.Columns.Item(1).TitleObject.Caption = "Select All"
                                oForm.Freeze(False)
                            End If

                            'MsgBox("dblclick grid column header: " & pVal.ColUID.ToString)

                        End If

                End Select
            End If

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try
    End Sub

End Module
