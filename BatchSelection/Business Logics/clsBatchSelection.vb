Public Class clsBatchSelection
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Function AddControls(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            oApplication.Utilities.AddControls(aForm, "BtnAuto", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, , "Auto Selection(FIFO)", 150)
            aForm.Items.Item("16").Enabled = False
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Function
#Region "Assign Serianumbers"
    Private Sub AssignBatchNumber(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        Try
            '   MsgBox(frmSourceForm.TypeEx)
            Dim oRowsMatrix, oSerialMatrix As SAPbouiCOM.Matrix
            Dim dblSelectedqty, MatQuantity, Quantity, diffQuantity As Double
            Dim strItemCode, strwhs, strSerial, strqry, MatSerial As String
            Dim oSerialRec, oTemp1 As SAPbobsCOM.Recordset
            oSerialRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRowsMatrix = aForm.Items.Item("3").Specific
            oSerialMatrix = aForm.Items.Item("4").Specific
            Dim strCardCode As String = ""
            If intSelectedMatrixrow <> 99999 Then
                If frmSourceForm.TypeEx = "85" Then
                    Dim OMatrix As SAPbouiCOM.Matrix
                    OMatrix = frmSourceForm.Items.Item("11").Specific
                    strCardCode = oApplication.Utilities.getMatrixValues(OMatrix, "17", intSelectedMatrixrow)
                ElseIf frmSourceForm.TypeEx = "81" Then
                    Dim OMatrix As SAPbouiCOM.Matrix
                    OMatrix = frmSourceForm.Items.Item("10").Specific
                    strCardCode = oApplication.Utilities.getMatrixValues(OMatrix, "10", intSelectedMatrixrow)
                Else
                    strCardCode = oApplication.Utilities.getEdittextvalue(frmSourceForm, "4")
                End If
            End If
            Dim blnAttributeRequired As Boolean = False
            If strCardCode <> "" Then
                oSerialRec.DoQuery("Select *,isnull(U_BAttriReq,'N') 'AttReq' from OCRD where CardCode='" & strCardCode & "'")
                If oSerialRec.RecordCount > 0 Then
                    If oSerialRec.Fields.Item("AttReq").Value = "Y" Then
                        blnAttributeRequired = True
                    Else
                        blnAttributeRequired = False
                    End If
                    If oSerialRec.Fields.Item("U_Attribute1").Value <> "" Then
                        strCardCode = oSerialRec.Fields.Item("U_Attribute1").Value
                    Else
                        strCardCode = ""
                    End If
                End If
            End If
            For intRow As Integer = 1 To oRowsMatrix.VisualRowCount
                oRowsMatrix = aForm.Items.Item("3").Specific
                oRowsMatrix.Columns.Item("0").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                strItemCode = oRowsMatrix.Columns.Item("1").Cells.Item(intRow).Specific.value
                dblSelectedqty = oRowsMatrix.Columns.Item("55").Cells.Item(intRow).Specific.value
                strwhs = oRowsMatrix.Columns.Item("3").Cells.Item(intRow).Specific.value
                If dblSelectedqty > 0 Then
                    strqry = "select DistNumber FROM OBTQ T0  INNER JOIN OBTN T1 ON T0.MdAbsEntry = T1.AbsEntry INNER JOIN OITM T2 ON "
                    If blnAttributeRequired = True Then
                        If strCardCode <> "" Then
                            strqry = strqry & " T0.ItemCode = T2.ItemCode where DateDiff(day,getDate(),T1.ExpDate)>90 and   T1.MnfSerial='" & strCardCode & "' and T2.ItemCode='" & strItemCode & "' and  T0.[Quantity] > 0 order by T1.ExpDate,T1.SysNumber asc "
                        Else
                            strqry = strqry & " T0.ItemCode = T2.ItemCode where DateDiff(day,getDate(),T1.ExpDate)>90 and  T2.ItemCode='" & strItemCode & "' and  T0.[Quantity] > 0 order by T1.ExpDate,T1.SysNumber asc "
                        End If
                    Else
                        If strCardCode <> "" Then
                            strqry = strqry & " T0.ItemCode = T2.ItemCode where DateDiff(day,getDate(),T1.ExpDate)>90 and   T1.MnfSerial<>'" & strCardCode & "' and T2.ItemCode='" & strItemCode & "' and  T0.[Quantity] > 0 order by T1.ExpDate,T1.SysNumber asc "
                        Else
                            strqry = strqry & " T0.ItemCode = T2.ItemCode where DateDiff(day,getDate(),T1.ExpDate)>90 and  T2.ItemCode='" & strItemCode & "' and  T0.[Quantity] > 0 order by T1.ExpDate,T1.SysNumber asc "
                        End If
                    End If
                    
                    oSerialRec.DoQuery(strqry)
                    Quantity = dblSelectedqty
                    diffQuantity = Quantity
                    Dim dblAlloc As Double
                    For intLoop As Integer = 0 To oSerialRec.RecordCount - 1
                        strSerial = oSerialRec.Fields.Item("DistNumber").Value
                        If Quantity >= 0 Then
                            For intloop1 As Integer = 1 To oSerialMatrix.VisualRowCount
                                MatSerial = oApplication.Utilities.getMatrixValues(oSerialMatrix, "0", intloop1)
                                MatQuantity = oApplication.Utilities.getMatrixValues(oSerialMatrix, "3", intloop1)
                                dblAlloc = oApplication.Utilities.getMatrixValues(oSerialMatrix, "24", intloop1)
                                MatQuantity = MatQuantity - dblAlloc
                                If strSerial = MatSerial Then
                                    If diffQuantity > 0 Then
                                        If diffQuantity >= MatQuantity Then
                                            oApplication.Utilities.SetMatrixValues(oSerialMatrix, "4", intloop1, MatQuantity)
                                            If aForm.TypeEx = frm_PickBatchSelect Then
                                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "1320000037", intloop1, MatQuantity)
                                            End If
                                            oSerialMatrix.Columns.Item(1).Cells.Item(intloop1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)
                                            diffQuantity = diffQuantity - MatQuantity
                                        ElseIf diffQuantity < MatQuantity Then
                                            oApplication.Utilities.SetMatrixValues(oSerialMatrix, "4", intloop1, diffQuantity)
                                            If aForm.TypeEx = frm_PickBatchSelect Then
                                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "1320000037", intloop1, diffQuantity)
                                            End If
                                            oSerialMatrix.Columns.Item(1).Cells.Item(intloop1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)
                                            diffQuantity = diffQuantity - MatQuantity
                                        End If
                                    Else
                                        If MatQuantity >= Quantity Then
                                            oApplication.Utilities.SetMatrixValues(oSerialMatrix, "4", intloop1, Quantity)
                                            If aForm.TypeEx = frm_PickBatchSelect Then
                                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "1320000037", intloop1, Quantity)
                                            End If
                                            oSerialMatrix.Columns.Item(1).Cells.Item(intloop1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)
                                            diffQuantity = Quantity - MatQuantity

                                        ElseIf Quantity > MatQuantity Then
                                            oApplication.Utilities.SetMatrixValues(oSerialMatrix, "4", intloop1, MatQuantity)
                                            If aForm.TypeEx = frm_PickBatchSelect Then
                                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "1320000037", intloop1, MatQuantity)
                                            End If
                                            oSerialMatrix.Columns.Item(1).Cells.Item(intloop1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0)
                                            diffQuantity = Quantity - MatQuantity
                                        End If
                                    End If
                                    aForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Exit For
                                End If
                            Next
                            oSerialRec.MoveNext()
                        End If
                        If diffQuantity <= 0 Then
                            Exit For
                        End If
                    Next
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
            Next
            aForm.Freeze(False)
            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            ElseIf aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BatchSelect Or pVal.FormTypeEx = frm_PickBatchSelect Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                frmSourceForm = oApplication.SBO_Application.Forms.ActiveForm()
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControls(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "BtnAuto" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to select the batches on FIFO Basic?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    AssignBatchNumber(oForm)
                                    oApplication.Utilities.Message("Operation Completed Successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "5896"
                    If pVal.BeforeAction = False Then
                        'oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        'AddControls(oForm)
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
