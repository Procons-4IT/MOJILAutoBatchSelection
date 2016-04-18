Public Class clsCustomerRoute
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oColumn As SAPbouiCOM.Column
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private RowtoDelete As Integer
    Private oMenuobject As Object
    Private InvForConsumedItems, count As Integer
    Private blnFlag As Boolean = False
    Dim MatrixId As Integer
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()

        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_CustomerRoute) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_CustomerRoute, frm_CustomerRoute)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        AddChooseFromList(oForm)
        oEditText = oForm.Items.Item("4").Specific
        oEditText.ChooseFromListUID = "CFL_3"
        oEditText.ChooseFromListAlias = "U_RouteCode"

        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu("1283", False)
        oForm.DataBrowser.BrowseBy = "4"
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oMatrix = oForm.Items.Item("8").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String
        Try
            aForm.Freeze(True)

            If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                Try
                    oForm.Items.Item("4").Enabled = True
                    oForm.Items.Item("6").Enabled = True
                    oForm.Items.Item("8").Enabled = True
                    oForm.Items.Item("7").Enabled = True
                Catch ex As Exception

                End Try
                oMatrix = aForm.Items.Item("8").Specific
                oMatrix.FlushToDataSource()
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
                For count = 1 To oDataSrc_Line.Size - 1
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strsubfee, strMAfee As Integer
        aForm.Freeze(True)
        If oApplication.Utilities.getEdittextvalue(oForm, "4") = "" Then
            oApplication.Utilities.Message("Route Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If
        Dim strCode As String = oApplication.Utilities.getEdittextvalue(aForm, "4")
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            Dim strterms, strLeavecode As String
            AddMode(aForm)
            strterms = oApplication.Utilities.getEdittextvalue(oForm, "4")
            otemp.DoQuery("Select * from ""@Z_OCURT"" where ""U_RouteCode""='" & strterms & "'")
            If otemp.RecordCount > 0 Then
                oApplication.Utilities.Message("This Entry already exists... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
        End If
        oMatrix = aForm.Items.Item("8").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Customer details are missing... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If
        oMatrix.FlushToDataSource()
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        aForm.Freeze(False)
        Return True
    End Function
    Private Function Matrix_Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strType, strValue, strCode As String
        oMatrix = aForm.Items.Item("7").Specific

        For intRow As Integer = 1 To oMatrix.RowCount
            'strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_-1", intRow)
            'strValue = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
            ''If strCode <> "" Then
            'oCombobox = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific
            'strType = oCombobox.Selected.Value
            'If strType = "" And strValue <> "" Then
            '    oApplication.Utilities.Message("Type is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'ElseIf strType <> "" And strValue = "" Then
            '    oApplication.Utilities.Message("Value is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'oMatrix.DeleteRow(intRow)
            'End If
        Next
        RefereshRowLineValues(aForm)
        Return True
    End Function

    Private Sub RefereshRowLineValues(ByVal aForm As SAPbouiCOM.Form)
        Try

            oMatrix = aForm.Items.Item("8").Specific
            oMatrix.FlushToDataSource()
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next

            oMatrix.LoadFromDataSource()

        Catch ex As Exception

        End Try


    End Sub
    Private Function CheckDuplicate(ByVal aCode As String, ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from ""@Z_OCURT"" where ""U_RouteCode""='" & aCode & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.Message("This entry already exists .....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return True
        End If
        Return False
    End Function
#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        oMatrix = aForm.Items.Item("8").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        Try
            aForm.Freeze(True)
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            Dim intRowCount As Integer = 1
            If intSelectedMatrixrow > 0 Then
                oEditText = oMatrix.Columns.Item("V_1").Cells.Item(intSelectedMatrixrow).Specific
                If oEditText.String <> "" Then
                    oMatrix.AddRow(1, intSelectedMatrixrow)
                    oMatrix.ClearRowData(intSelectedMatrixrow + 1)
                    ' oMatrix.Columns.Item("V_1").Cells.Item(intSelectedMatrixrow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    intRowCount = intSelectedMatrixrow + 1
                End If
            Else
                oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                If oEditText.String <> "" Then
                    oMatrix.AddRow()
                    oMatrix.ClearRowData(oMatrix.RowCount)
                    ' oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    intRowCount = oMatrix.RowCount
                End If
            End If

            Try
              
            Catch ex As Exception
                aForm.Freeze(False)
                oMatrix.AddRow()
            End Try
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_1").Cells.Item(intRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#End Region

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)

        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        frmSourceMatrix = aForm.Items.Item("8").Specific
        If intSelectedMatrixrow <= 0 Then
            Exit Sub
        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        aForm.Freeze(False)

    End Sub


    Private Sub DeleteRow(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)

        oMatrix = aform.Items.Item("8").Specific
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CURT1")
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
            End If
        Next
        aform.Freeze(False)
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CustomerRoute Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If (pVal.ItemUID = "8") And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "8"
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    'Case "10"
                                    '    oForm.PaneLevel = 1
                                    'Case "11"
                                    '    oForm.PaneLevel = 2
                                    'Case "12"
                                    '    oForm.PaneLevel = 3
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oItm As SAPbobsCOM.Items
                                Dim sCHFL_ID, val As String
                                Dim intChoice, introw As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "4" Then
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", oDataTable.GetValue("U_RouteName", 0))
                                            Try
                                                oCombobox = oForm.Items.Item("10").Specific
                                                oCombobox.Select(oDataTable.GetValue("U_RouteType", 0), SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            Catch ex As Exception
                                            End Try
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, "4", oDataTable.GetValue("U_RouteCode", 0))
                                            Catch ex As Exception
                                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                End If
                                            End Try
                                        End If
                                        If pVal.ItemUID = "8" And pVal.ColUID = "V_1" Then
                                            oMatrix = oForm.Items.Item("8").Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", pVal.Row, oDataTable.GetValue("CardName", 0))
                                            Try
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, oDataTable.GetValue("CardCode", 0))

                                            Catch ex As Exception
                                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                                End If
                                            End Try
                                           

                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    'MsgBox(ex.Message)
                                End Try

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_CustomerRoute
                    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        Exit Sub
                    End If
                    AddRow(oForm)
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("8").Enabled = True
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Try
                    oForm.Items.Item("4").Enabled = False
                    oForm.Items.Item("6").Enabled = True
                    oForm.Items.Item("8").Enabled = True
                Catch ex As Exception

                End Try

            End If

            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_CustomerRoute Then
                    Dim strRoute As String = oApplication.Utilities.getEdittextvalue(oForm, "4")
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("Select * from [@Z_OCURT] where U_RouteCode='" & strRoute & "'")
                    If oTest.RecordCount > 0 Then
                        oTest.DoQuery("Update [@Z_ORUT] set U_Active='" & oTest.Fields.Item("U_Active").Value & "' where U_RouteCode='" & strRoute & "'")

                    End If

                End If
            End If






        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
   

End Class
