Public Class clsTax
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oTemp As SAPbobsCOM.Recordset
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems, count As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Dim oDataSrc_Line, oDataSrc_Line1 As SAPbouiCOM.DBDataSource
    Private RowtoDelete As Integer
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_TaxMaster, frm_LEB_TaxMaster)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu("1287", True)
        oForm.EnableMenu("1283", False)

        oForm.DataBrowser.BrowseBy = "4"
        oCombobox = oForm.Items.Item("4").Specific

        oCombobox.ValidValues.Add("", "")
        For intRow As Integer = 2010 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oForm.Items.Item("4").DisplayDesc = True
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

        oCombobox = oForm.Items.Item("30").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 1 To 31
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oForm.Items.Item("30").DisplayDesc = True
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oCombobox = oForm.Items.Item("32").Specific
        oCombobox.ValidValues.Add("33", "")
        For intRow As Integer = 1 To 31
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oForm.Items.Item("32").DisplayDesc = True
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oMatrix = oForm.Items.Item("8").Specific
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oMatrix.AutoResizeColumns()
        oMatrix = oForm.Items.Item("9").Specific
        Dim ocolumn As SAPbouiCOM.Column
        ocolumn = oMatrix.Columns.Item("V_1")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select U_Z_Code,U_Z_Name from [@Z_PAY_OFAM] order by convert(numeric,code)")
        ocolumn.ValidValues.Add("", "")
        For intRow As Integer = 0 To oTemp.RecordCount - 1
            ocolumn.ValidValues.Add(oTemp.Fields.Item(0).Value, oTemp.Fields.Item(1).Value)
            oTemp.MoveNext()
        Next
        ocolumn.DisplayDesc = True
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        oMatrix.AutoResizeColumns()
        AddChooseFromList(oForm)
        Databind(oForm)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub

#Region "AddChooseFromList"
    Private Sub AddChooseFromList(ByVal aform As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = aform.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL1"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"


            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL2"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"

            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "1"
            oCFLCreationParams.UniqueID = "CFL3"

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub


#End Region
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oDataSrc_Line1 = oForm.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line1.SetValue("LineId", count - 1, count)
            Next
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'oEditText = aform.Items.Item("26").Specific
            'oEditText.ChooseFromListUID = "CFL1"
            'oEditText.ChooseFromListAlias = "FormatCode"
            'oEditText = aform.Items.Item("28").Specific
            'oEditText.ChooseFromListUID = "CFL2"
            'oEditText.ChooseFromListAlias = "FormatCode"
            'oEditText = aform.Items.Item("30").Specific
            'oEditText.ChooseFromListUID = "CFL3"
            'oEditText.ChooseFromListAlias = "FormatCode"
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub calculateamount(ByVal aform As SAPbouiCOM.Form, Optional ByVal arow As Integer = 9999)
        Dim dblEndAmount, dblPercentage, dblamount, dblStartAmount As Double
        oMatrix = aform.Items.Item("8").Specific
        aform.Freeze(True)
        If arow <> 9999 Then
            dblStartAmount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", arow))
            dblEndAmount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", arow))
            dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_3", arow))
            dblamount = (dblEndAmount - dblStartAmount) * dblPercentage / 100
            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", arow, Math.Round(dblamount, 0))
        Else
            For arow = 1 To oMatrix.RowCount
                dblStartAmount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", arow))
                dblEndAmount = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", arow))
                dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_3", arow))
                dblamount = (dblEndAmount - dblStartAmount) * dblPercentage / 100
                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", arow, Math.Round(dblamount, 0))
            Next
        End If
        aform.Freeze(False)
    End Sub

#Region "AddMode"
    Private Sub AddMode(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strCode = oApplication.Utilities.getMaxCode("@Z_PAY_TAX", "DocEntry")
            oApplication.Utilities.setEdittextvalue(aForm, "11", strCode)
            'oApplication.Utilities.setEdittextvalue(aForm, "6", "")
            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("11").Enabled = False
            'oForm.Items.Item("4").Enabled = True
        End If
    End Sub
#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            If aForm.PaneLevel = 1 Then
                oMatrix = aForm.Items.Item("8").Specific
                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
            ElseIf aForm.PaneLevel = 2 Then
                oMatrix = aForm.Items.Item("9").Specific
                oDataSrc_Line1 = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
            Else
                Exit Sub
            End If
            count = 0
            If aForm.PaneLevel = 1 Then
                If oMatrix.RowCount <= 0 Then
                    oMatrix.AddRow()
                    oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Else
                    Dim dblAmount As Double
                    oEditText = oMatrix.Columns.Item("V_3").Cells.Item(oMatrix.RowCount).Specific
                    dblAmount = oApplication.Utilities.getDocumentQuantity(oEditText.String)
                    Try
                        If 1 = 1 Then
                            If dblAmount > 0 Then
                                oMatrix.AddRow()
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "0")
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, "0")
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "0")
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "0")
                                oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                    Catch ex As Exception
                        ' oMatrix.AddRow()
                    End Try
                End If
                oMatrix.FlushToDataSource()
                oMatrix = aForm.Items.Item("8").Specific
                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
                For count = 1 To oDataSrc_Line.Size
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()

            Else
                If oMatrix.RowCount <= 0 Then
                    oMatrix.AddRow()
                    oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Else
                    Dim dblAmount As Double
                    oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                    dblAmount = oApplication.Utilities.getDocumentQuantity(oEditText.String)
                    Try
                        If 1 = 1 Then
                            If dblAmount > 0 Then
                                oMatrix.AddRow()
                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "0")
                                oCombobox = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                                oCombobox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                End If
                oMatrix.FlushToDataSource()
                oMatrix = aForm.Items.Item("9").Specific
                oDataSrc_Line1 = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
                For count = 1 To oDataSrc_Line1.Size
                    oDataSrc_Line1.SetValue("LineId", count - 1, count)
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

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        aForm.Freeze(True)

        If aForm.PaneLevel = 1 Then
            oMatrix = aForm.Items.Item("8").Specific

            oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
            oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            'If oMatrix.RowCount > 0 Then
            '    oMatrix.DeleteRow(oMatrix.RowCount)
            'End If
        ElseIf aForm.PaneLevel = 2 Then
            oMatrix = aForm.Items.Item("9").Specific
            oDataSrc_Line1 = aForm.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
            oDataSrc_Line1.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line1.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            'If oMatrix.RowCount > 0 Then
            '    oMatrix.DeleteRow(oMatrix.RowCount)
            'End If
        Else
            Exit Sub
        End If







        aForm.Freeze(False)

    End Sub

    Private Sub DeleteRow(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        If aform.PaneLevel = 1 Then
            oMatrix = aform.Items.Item("8").Specific
            oDataSrc_Line = aform.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
        ElseIf aform.PaneLevel = 2 Then
            oMatrix = aform.Items.Item("9").Specific
            oDataSrc_Line = aform.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
        Else
            Exit Sub
        End If

        Dim intRow As Integer
        For intRow = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                oMatrix.DeleteRow(intRow)
                ' AddRow(aform)
                If aform.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And aform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    aform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                aform.Freeze(False)
                Exit Sub
            End If
        Next
        aform.Freeze(False)
    End Sub


    Private Function CheckDuplicate(ByVal aCode As String) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from [@Z_PAY_TAX] where U_Z_Type='" & aCode & "'")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.Message("Income Tax Details already defined for selected Year .....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return True
        End If
        Return False
    End Function

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strsubfee, strMAfee As Integer

        oCombobox = aForm.Items.Item("4").Specific
        If oCombobox.Selected.Value = "" Then
            oApplication.Utilities.Message("Year is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select * from [@Z_PAY_TAX] where U_Z_Year=" & oCombobox.Selected.Value)
            If otemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Income Tax Details already defined for selected Year... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        'If oApplication.Utilities.getEdittextvalue(aForm, "26") = "" Then
        '    oApplication.Utilities.Message("Family Allowance G/L Account missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If
        'If oApplication.Utilities.getEdittextvalue(aForm, "28") = "" Then
        '    oApplication.Utilities.Message("Hospitalization Employee G/L Account missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If
        'If oApplication.Utilities.getEdittextvalue(aForm, "30") = "" Then
        '    oApplication.Utilities.Message("Hospitalization Employeer G/L Account missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If
        oMatrix = aForm.Items.Item("8").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Tax  details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        Dim dblStart, dblEnd, dblPrevEnd, dblPercentage As Double
        For introw As Integer = 1 To oMatrix.RowCount
            dblStart = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_1", introw))
            dblEnd = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", introw))
            dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_3", introw))
            If dblEnd <= 0 Then
                oApplication.Utilities.Message("To Slab should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oMatrix.Columns.Item("V_2").Cells.Item(introw).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                Return False
            End If

            If dblStart > dblEnd Then
                oApplication.Utilities.Message("To Slab should be greater than from Slab amount", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oMatrix.Columns.Item("V_2").Cells.Item(introw).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                Return False
            End If

            If dblPercentage <= 0 Then
                oApplication.Utilities.Message("Percentage should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oMatrix.Columns.Item("V_3").Cells.Item(introw).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                Return False
            End If

            If introw > 1 Then
                dblPrevEnd = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_2", introw - 1))
                If dblStart <= dblPrevEnd Then
                    oApplication.Utilities.Message("From Slab should be greater than the previous End Slab", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oMatrix.Columns.Item("V_1").Cells.Item(introw).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 1)
                    Return False
                End If
            Else
                dblPrevEnd = dblEnd + 1
            End If
        Next
        oMatrix.FlushToDataSource()
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAY_TAX1")
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        oMatrix = aForm.Items.Item("9").Specific
        If oMatrix.RowCount <= 0 Then
            oApplication.Utilities.Message("Exculsion details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        oMatrix.FlushToDataSource()
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PAY_TAX2")
        For count = 1 To oDataSrc_Line.Size
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        calculateamount(aForm)
        Return True
    End Function


    Private Sub RefereshRowLineValues(ByVal aForm As SAPbouiCOM.Form)
        Try

            oMatrix = aForm.Items.Item("8").Specific
            For introw As Integer = oMatrix.RowCount - 1 To 0 Step -1
                If oMatrix.Columns.Item("V_0").Cells.Item(introw).Specific.value = "" Then
                    oMatrix.DeleteRow(introw)
                End If

            Next
            oMatrix.FlushToDataSource()

            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_OVAG1")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next

            oMatrix.LoadFromDataSource()

        Catch ex As Exception

        End Try


    End Sub
    Private Function Matrix_Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strType, strValue, strCode As String
        oMatrix = aForm.Items.Item("8").Specific

        For intRow As Integer = 1 To oMatrix.RowCount
            strCode = oApplication.Utilities.getMatrixValues(oMatrix, "V_-1", intRow)
            strValue = oApplication.Utilities.getMatrixValues(oMatrix, "V_1", intRow)
            'If strCode <> "" Then
            oCombobox = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific
            strType = oCombobox.Selected.Value
            If strType = "" And strValue <> "" Then
                oApplication.Utilities.Message("Type is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strType <> "" And strValue = "" Then
                oApplication.Utilities.Message("Value is missing . Line Number : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            'oMatrix.DeleteRow(intRow)
            'End If
        Next
        ' RefereshRowLineValues(aForm)
        Return True
    End Function


#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strChoice = "Cancel" Then
            oTemprec.DoQuery("Update [@Z_PAY_OTAX] set Name=Code where Name Like '%D'")
        Else
            oTemprec.DoQuery("Select * from [@Z_PAY_OTAX] where Name like '%D'")
            For intRow As Integer = 0 To oTemprec.RecordCount - 1
                oItemRec.DoQuery("delete from [@Z_PAY_OTAX] where Name='" & oTemprec.Fields.Item("Name").Value & "' and Code='" & oTemprec.Fields.Item("Code").Value & "'")
                oTemprec.MoveNext()
            Next
            oTemprec.DoQuery("Delete from  [@Z_PAY_OTAX]  where Name Like '%D'")
        End If

    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strTCode, strFrom, strTo, StrPercentage As String

        oGrid = aform.Items.Item("5").Specific
        If validation(oGrid) = False Then
            Return False
        End If
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strCode = oApplication.Utilities.getMaxCode("@Z_PAY_OTAX", "Code")
            ' If oGrid.DataTable.GetValue(2, intRow) <> "" Or oGrid.DataTable.GetValue(3, intRow) <> "" Then
            If 1 = 1 Then
                strTCode = oGrid.DataTable.GetValue(0, intRow)
                strFrom = oGrid.DataTable.GetValue(2, intRow)
                strTo = oGrid.DataTable.GetValue(3, intRow)
                StrPercentage = oGrid.DataTable.GetValue(4, intRow)
                If StrPercentage <> "" Then
                    If CDbl(StrPercentage > 0) Then


                        oUserTable = oApplication.Company.UserTables.Item("Z_PAY_OTAX")
                        If oUserTable.GetByKey(strTCode) Then
                            oUserTable.Code = strCode
                            oUserTable.Name = strCode
                            oUserTable.UserFields.Fields.Item("U_Z_SLAP_FROM").Value = (oGrid.DataTable.GetValue(2, intRow))
                            oUserTable.UserFields.Fields.Item("U_Z_SLAP_TO").Value = (oGrid.DataTable.GetValue(3, intRow))
                            oUserTable.UserFields.Fields.Item("U_Z_TAX_PERC_TAGE").Value = (oGrid.DataTable.GetValue(4, intRow))
                            If oUserTable.Update <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oUserTable.Code = strCode
                            oUserTable.Name = strCode
                            oUserTable.UserFields.Fields.Item("U_Z_SLAP_FROM").Value = (oGrid.DataTable.GetValue(2, intRow))
                            oUserTable.UserFields.Fields.Item("U_Z_SLAP_TO").Value = (oGrid.DataTable.GetValue(3, intRow))
                            oUserTable.UserFields.Fields.Item("U_Z_TAX_PERC_TAGE").Value = (oGrid.DataTable.GetValue(4, intRow))
                            If oUserTable.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                End If

            End If
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Committrans("Add")
        Databind(aform)
    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_LEB_TaxMaster Then
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
                                If (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.CharPressed <> 9 Then
                                    If pVal.ItemUID = "4" Then
                                        '  BubbleEvent = False
                                        ' Exit Sub
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" Or pVal.ItemUID = "9" Then
                                    Me.RowtoDelete = pVal.Row
                                End If

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "8" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_2") And pVal.CharPressed = 9 Then
                                    calculateamount(oForm, pVal.Row)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "6"
                                        oForm.PaneLevel = 1
                                    Case "7"
                                        oForm.PaneLevel = 2
                                    Case "12"
                                        oForm.PaneLevel = 3
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim objEdit As SAPbouiCOM.EditTextColumn
                                Dim oGr As SAPbouiCOM.Grid
                                Dim oItm As SAPbobsCOM.BusinessPartners
                                Dim sCHFL_ID, val, strBPCode As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    Dim oDataTable As SAPbouiCOM.DataTable
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "26" Or pVal.ItemUID = "28" Or pVal.ItemUID = "30" Then
                                        val = oDataTable.GetValue("FormatCode", 0)
                                        oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                    End If
                                Catch ex As Exception

                                End Try
                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
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
                Case mnu_InvSO
                Case mnu_Tax
                    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        AddRow(oForm)
                        BubbleEvent = False
                        Exit Sub
                    End If
                Case "1287"
                    If pVal.BeforeAction = True Then
                        If oApplication.SBO_Application.MessageBox("Do you want to Duplicate the TAX and NSSF Setup?", , "Continue", "Cancel") = 2 Then
                            BubbleEvent = False
                            Exit Sub

                        End If
                    End If

                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        DeleteRow(oForm)
                        RefereshDeleteRow(oForm)

                        BubbleEvent = False
                        Exit Sub
                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddMode(oForm)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("11").Enabled = True
                        ' oForm.Items.Item("6").Enabled = True
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_Tax
                        oMenuobject = New clsTax
                        oMenuobject.MenuEvent(pVal, BubbleEvent)
                End Select
            End If
        Catch ex As Exception
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
