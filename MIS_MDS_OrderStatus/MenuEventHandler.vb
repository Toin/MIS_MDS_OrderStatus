Module MenuEventHandler
    Public WithEvents oApp4MenuEvent As SAPbouiCOM.Application = Nothing

    Sub MenuEventHandler(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) _
    Handles oApp4MenuEvent.MenuEvent
        Try

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "ORD01_01"
                        'SOToMFGEntry()
                        PdoDueDateMassUpdation_FormEntry()
                    Case "ORD01_03"
                        'PdoDueDateMassUpdationByMachine_FormEntry()
                        MassUpdateDelivery_FormEntry()
                    Case "ORD01_02"
                        'OutDelEntry()
                        DelivPlanDateMassUpdation_FormEntry()
                    Case "ORD01_04"
                        'OrderStatusForMarketing_FormEntry()
                        OrderStatus_FormEntry()

                        '-------------- Yadi FC ----------------------------
                    Case "ORD01_05"
                        'ProductionClosed()
                        '-------------- Yadi FC ----------------------------
                End Select
            End If

            If pVal.BeforeAction = True Then
                Dim oForm As SAPbouiCOM.Form

                'oForm = SBO_Application.Forms.ActiveForm
                oForm = oApp.Forms.ActiveForm
                'MsgBox(oForm.Type)
                'MsgBox(oForm.TypeEx)
                'MsgBox(oForm.UniqueID)
                Select Case pVal.MenuUID
                    Case "1290" ' 1st Record
                    Case "1289" ' Prev Record
                    Case "1288" ' Next Record
                    Case "1291" ' Last Record
                    Case "1292" ' Add a row
                        'MsgBox("Add a row")

                    Case "1293" ' Delete a row
                        'MsgBox("Delete a row")
                    Case "1282"
                        'MsgBox("add new doc!")

                End Select

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                GC.Collect()

            End If

        Catch ex As Exception
            MsgBoxWrapper(ex.Message)
        End Try

    End Sub


End Module
