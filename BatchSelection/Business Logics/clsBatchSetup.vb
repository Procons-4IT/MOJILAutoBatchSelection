Public Class clsBatchSetup
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
            oApplication.Utilities.AddControls(aForm, "BtnAuto", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, , "Create Batches", 150)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Function
#Region "Assign Serianumbers"
    Private Sub AssignBatchNumber(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)
        Try
            Dim oRowsMatrix, oSerialMatrix As SAPbouiCOM.Matrix
            Dim dblSelectedqty, MatQuantity, Quantity, diffQuantity As Double
            Dim strItemCode, strwhs, strQry, strBatchNumber, strMaxBatch As String
            Dim BatchNumber As Integer
            Dim strBatchPrefix As String = ""
            Dim strmaxCode As String = ""
            Dim oSerialRec, oTemp1, oTemp, oTemp2 As SAPbobsCOM.Recordset
            oSerialRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRowsMatrix = aForm.Items.Item("35").Specific
            oSerialMatrix = aForm.Items.Item("3").Specific
            Dim strDate As String
            Dim dtDate As Date
            Select Case frmSourceForm.TypeEx
                Case "65214"
                    strDate = oApplication.Utilities.getEdittextvalue(frmSourceForm, "9")
                Case "143"
                    strDate = oApplication.Utilities.getEdittextvalue(frmSourceForm, "10")
                Case "721"
                    strDate = oApplication.Utilities.getEdittextvalue(frmSourceForm, "9")

            End Select
            Dim dtDate1 As Date = CDate(strDate)
            strDate = dtDate1.ToString("ddMMyyyy")
            Dim ote As SAPbobsCOM.Recordset
            ote = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRouteRS As SAPbobsCOM.Recordset
            oRouteRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRouteRS.DoQuery("Select * from [@Z_ItemType] ")
            Dim strType As String
            Dim intCount As Integer = 0
            For intLoop As Integer = 0 To oRouteRS.RecordCount - 1
                strType = oRouteRS.Fields.Item("U_ItemType").Value
                'ote.DoQuery("Update OITM set U_Z_BatchNumber=''")
                For intRow As Integer = 1 To oRowsMatrix.RowCount
                    oRowsMatrix = aForm.Items.Item("35").Specific
                    oRowsMatrix.Columns.Item("0").Cells.Item(intRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    strItemCode = oRowsMatrix.Columns.Item("5").Cells.Item(intRow).Specific.value
                    dblSelectedqty = oRowsMatrix.Columns.Item("39").Cells.Item(intRow).Specific.value
                    strwhs = oRowsMatrix.Columns.Item("40").Cells.Item(intRow).Specific.value
                    Dim intExpDays As Integer = 1
                    strQry = "SELECT * FROM OITM where U_Z_ItemType='" & strType & "' and ItemCode='" & strItemCode & "'"
                    oTemp.DoQuery(strQry)
                    If oTemp.RecordCount > 0 Then
                        If dblSelectedqty > 0 Then

                            strQry = "SELECT U_Z_ItemType FROM OITM where ItemCode='" & strItemCode & "'"
                            oTemp.DoQuery(strQry)
                            If oTemp.RecordCount > 0 Then
                                oTemp.DoQuery("Select * from [@Z_ItemType] where U_ItemType='" & oTemp.Fields.Item(0).Value & "'")
                                strBatchPrefix = oTemp.Fields.Item("U_BatchPrefix").Value
                                intExpDays = oTemp.Fields.Item("U_BatExpDays").Value
                            Else
                                strBatchPrefix = "B"
                            End If
                            If strBatchPrefix = "" Then
                                strBatchPrefix = "B"
                            End If
                            If 1 = 1 Then 'oTemp1.RecordCount > 0 Then
                                strQry = "SELECT Top 1 DistNumber FROM [OBTN] where distNumber  like '" & strBatchPrefix & strDate & "%' order by SysNumber Desc"
                                oTemp2.DoQuery(strQry)
                                If oTemp2.RecordCount > 0 Then
                                    strBatchNumber = oTemp2.Fields.Item(0).Value
                                    If strBatchNumber = "" Then
                                        strBatchNumber = oTemp2.Fields.Item("DistNumber").Value
                                    End If
                                    strBatchNumber = strBatchNumber.Replace(strBatchPrefix, "")
                                    strBatchNumber = strBatchNumber.Replace(strDate, "")
                                    If intCount > 0 Then
                                        BatchNumber = CInt(strBatchNumber + intCount)
                                    Else
                                        BatchNumber = CInt(strBatchNumber)
                                    End If
                                    strmaxCode = Format(BatchNumber + 1, "000")
                                Else
                                    strmaxCode = Format(1, "000")
                                End If
                            End If
                            strMaxBatch = strBatchPrefix & strDate & strmaxCode
                            Dim dtExpDate As Date = dtDate1.AddDays(intExpDays)
                            Dim strBatch As String = strMaxBatch
                            Dim intRow1 As Integer = 1
                            If oSerialMatrix.RowCount > 1 Then
                                intRow1 = oSerialMatrix.RowCount - 1
                            End If
                            Dim str As String = dtExpDate.ToString("dd.MM.yy")
                            For intloop1 As Integer = intRow1 To intRow1
                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "2", intloop1, strMaxBatch)
                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "5", intloop1, dblSelectedqty)
                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "10", intloop1, dtExpDate.ToString("yyyyMMdd"))
                                oApplication.Utilities.SetMatrixValues(oSerialMatrix, "11", intloop1, dtDate1.ToString("yyyyMMdd"))
                            Next
                            intCount = intCount + 1
                            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                    End If
                Next
                oRouteRS.MoveNext()
            Next
            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
            aForm.Freeze(False)
            'If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            '    aForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'End If

        Catch ex As Exception
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

    'Private Function AddtoUDT1(ByVal aItemCode As String, ByVal aBatchNumber As String) As Boolean
    '    Dim oUserTable As SAPbobsCOM.UserTable
    '    Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
    '    Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
    '    oGrid = aform.Items.Item("5").Specific
    '    For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
    '        '
    '        If oGrid.DataTable.GetValue(0, intRow) <> "" Or oGrid.DataTable.GetValue(1, intRow) <> "" Then
    '            strCode = oGrid.DataTable.GetValue(0, intRow)
    '            strECode = oGrid.DataTable.GetValue(1, intRow)
    '            ' strGLAcc = oGrid.DataTable.GetValue(2, intRow)
    '            oUserTable = oApplication.Company.UserTables.Item("Z_OBTN")
    '            If oUserTable.GetByKey(strCode) = False Then
    '                'strCode = oApplication.Utilities.getMaxCode("@Z_PAY_LOAN", "Code")
    '                oUserTable.Code = strCode
    '                oUserTable.Name = strECode
    '                ' oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = (oGrid.DataTable.GetValue(2, intRow))
    '                If oUserTable.Add() <> 0 Then
    '                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                End If

    '            Else
    '                strCode = oGrid.DataTable.GetValue(0, intRow)
    '                If oUserTable.GetByKey(strCode) Then
    '                    oUserTable.Code = strCode
    '                    oUserTable.Name = strECode
    '                    ' oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = (oGrid.DataTable.GetValue(2, intRow))
    '                    If oUserTable.Update() <> 0 Then
    '                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        Committrans("Cancel")
    '                        Return False
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Next
    '    oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Committrans("Add")
    '    Databind(aform)
    'End Function


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BatchSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
                                    If oApplication.SBO_Application.MessageBox("Do you want to Create the batches Automatically?", , "Yes", "No") = 2 Then
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
                Case "5946", "5896"
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        AssignBatchNumber(oForm)
                    End If
                    If pVal.BeforeAction = True Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        frmSourceForm = oForm
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
