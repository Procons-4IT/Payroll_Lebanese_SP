Public Class clsVariableEarning
    Inherits clsBase
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oCheckbox As SAPbouiCOM.CheckBox
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private oTemp As SAPbobsCOM.Recordset
    Private InvBaseDocNo, strname As String
    Private InvForConsumedItems As Integer
    Private oMenuobject As Object
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_VariableEarning) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_VariableEarning, frm_VariableEarning)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        AddChooseFromList(oForm)
        Databind(oForm)
    End Sub
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            oGrid = aform.Items.Item("5").Specific
            dtTemp = oGrid.DataTable
            '  dtTemp.ExecuteQuery("Select * from [@Z_PAY_OEAR] order by CODE")
            dtTemp.ExecuteQuery("SELECT T0.[Code], T0.[Name],T0.[U_Z_FrgnName], T0.[U_Z_DefAmt], T0.[U_Z_SOCI_BENE], T0.[U_Z_INCOM_TAX], T0.[U_Z_EOS],T0.[U_Z_AvgYear], T0.[U_Z_AffDedu],T0.[U_Z_Max],T0.[U_Z_BaiscPer], T0.[U_Z_OffCycle], T0.[U_Z_EAR_GLACC],T0.[U_Z_DED_GLACC], T0.[U_Z_PostType] FROM [dbo].[@Z_PAY_OEAR1]  T0 order by Code")
            oGrid.DataTable = dtTemp
            Formatgrid(oGrid)
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL = oCFLs.Item("CFL1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL11")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item("Code").TitleObject.Caption = "Earning Code"
        agrid.Columns.Item("Name").TitleObject.Caption = "Earning Name"
        agrid.Columns.Item("U_Z_FrgnName").TitleObject.Caption = "Second Language Name"
        agrid.Columns.Item("U_Z_EAR_GLACC").TitleObject.Caption = "G/L Account"
        oEditTextColumn = agrid.Columns.Item("U_Z_EAR_GLACC")
        oEditTextColumn.LinkedObjectType = "1"
        oEditTextColumn.ChooseFromListUID = "CFL1"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"

        agrid.Columns.Item("U_Z_DED_GLACC").TitleObject.Caption = "Deduction G/L Account"
        oEditTextColumn = agrid.Columns.Item("U_Z_DED_GLACC")
        oEditTextColumn.LinkedObjectType = "1"
        oEditTextColumn.ChooseFromListUID = "CFL11"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"
        agrid.Columns.Item("U_Z_SOCI_BENE").TitleObject.Caption = "Affect NSSF"
        agrid.Columns.Item("U_Z_SOCI_BENE").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_SOCI_BENE").Visible = True
        agrid.Columns.Item("U_Z_INCOM_TAX").TitleObject.Caption = "Taxable"
        agrid.Columns.Item("U_Z_INCOM_TAX").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_INCOM_TAX").Visible = True

        agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "Affect EOS"
        agrid.Columns.Item("U_Z_EOS").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_EOS").Visible = True

        agrid.Columns.Item("U_Z_AffDedu").TitleObject.Caption = "Part of Payroll Deduction"
        agrid.Columns.Item("U_Z_AffDedu").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_AffDedu").Visible = True

        agrid.Columns.Item("U_Z_OffCycle").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item("U_Z_OffCycle").TitleObject.Caption = "Off Cycle "
        agrid.Columns.Item("U_Z_OffCycle").Visible = False
      
        agrid.Columns.Item("U_Z_PostType").TitleObject.Caption = "Posting Type"
        agrid.Columns.Item("U_Z_PostType").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oComboColumn = agrid.Columns.Item("U_Z_PostType")
        oComboColumn.ValidValues.Add("B", "Business Partner")
        oComboColumn.ValidValues.Add("A", "G/L Account")
        oComboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both


        agrid.Columns.Item("U_Z_DefAmt").TitleObject.Caption = "Default Amount "
        agrid.Columns.Item("U_Z_DefAmt").Editable = True


     

        agrid.Columns.Item("U_Z_Max").TitleObject.Caption = "Max.Exemption Amount "
        agrid.Columns.Item("U_Z_Max").Editable = True
        agrid.Columns.Item("U_Z_AvgYear").TitleObject.Caption = "Average EOS Affect Year"
        agrid.Columns.Item("U_Z_AvgYear").Editable = True

        agrid.Columns.Item("U_Z_BaiscPer").TitleObject.Caption = "Percentage in Basic Salary"


        'agrid.Columns.Item("U_Z_EOS").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        'agrid.Columns.Item("U_Z_EOS").TitleObject.Caption = "Include for EOS"
        'oCheckbox.Checked = True
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "AddRow"
    Private Sub AddEmptyRow(ByVal aGrid As SAPbouiCOM.Grid)
        If aGrid.DataTable.GetValue("Code", aGrid.DataTable.Rows.Count - 1) <> "" Then
            aGrid.DataTable.Rows.Add()
            aGrid.Columns.Item("Code").Click(aGrid.DataTable.Rows.Count - 1, False)
        End If
    End Sub
#End Region

#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strChoice = "Cancel" Then
            oTemprec.DoQuery("Update [@Z_PAY_OEAR1] set NAME=CODE where Name Like '%_XD'")
        Else
            oTemprec.DoQuery("Delete from  [@Z_PAY_OEAR1]  where NAME Like '%_XD'")
        End If

    End Sub
#End Region

#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc, strType As String
        Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
        oGrid = aform.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oGrid.DataTable.GetValue("Code", intRow) <> "" Or oGrid.DataTable.GetValue("Name", intRow) <> "" Then
                strCode = oGrid.DataTable.GetValue(0, intRow)
                strECode = oGrid.DataTable.GetValue(1, intRow)
                OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_SOCI_BENE")
                If OCHECKBOXCOLUMN.IsChecked(intRow) = True Then
                    strESocial = "Y"
                Else
                    strESocial = "N"
                End If
                OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_INCOM_TAX")
                If OCHECKBOXCOLUMN.IsChecked(intRow) = True Then
                    strETax = "Y"
                Else
                    strETax = "N"
                End If
                Dim stPosttype As String
                oComboColumn = oGrid.Columns.Item("U_Z_PostType")
                Try
                    stPosttype = oComboColumn.GetSelectedValue(intRow).Value
                Catch ex As Exception
                    stPosttype = "A"
                End Try
                'strbindesc = oGrid.DataTable.GetValue(5, intRow)
                oUserTable = oApplication.Company.UserTables.Item("Z_PAY_OEAR1")
                If oUserTable.GetByKey(strCode) = False Then
                    oUserTable.Code = strCode
                    oUserTable.Name = strECode
                    oUserTable.UserFields.Fields.Item("U_Z_AvgYear").Value = oGrid.DataTable.GetValue("U_Z_AvgYear", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_FrgnName").Value = oGrid.DataTable.GetValue("U_Z_FrgnName", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_DefAmt").Value = oGrid.DataTable.GetValue("U_Z_DefAmt", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_SOCI_BENE").Value = strESocial
                    oUserTable.UserFields.Fields.Item("U_Z_INCOM_TAX").Value = strETax
                    oUserTable.UserFields.Fields.Item("U_Z_Max").Value = oGrid.DataTable.GetValue("U_Z_Max", intRow)
                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_OffCycle")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        oUserTable.UserFields.Fields.Item("U_Z_OffCycle").Value = "Y"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_OffCycle").Value = "N"
                    End If
                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_EOS")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        oUserTable.UserFields.Fields.Item("U_Z_EOS").Value = "Y"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_EOS").Value = "N"
                    End If


                    OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_AffDedu")
                    If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                        oUserTable.UserFields.Fields.Item("U_Z_AffDedu").Value = "Y"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_AffDedu").Value = "N"
                    End If
                    oUserTable.UserFields.Fields.Item("U_Z_EAR_GLACC").Value = (oGrid.DataTable.GetValue("U_Z_EAR_GLACC", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_DED_GLACC").Value = (oGrid.DataTable.GetValue("U_Z_DED_GLACC", intRow))

                    oUserTable.UserFields.Fields.Item("U_Z_BaiscPer").Value = (oGrid.DataTable.GetValue("U_Z_BaiscPer", intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_PostType").Value = stPosttype
                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Committrans("Cancel")
                        Return False
                    Else
                        'If AddToUDT_Employee(oGrid.DataTable.GetValue(2, intRow).ToString.ToUpper(), oGrid.DataTable.GetValue("U_Z_Percentage", intRow), oGrid.DataTable.GetValue(4, intRow)) = False Then
                        '    Return False
                        'End If
                    End If
                Else
                    strCode = oGrid.DataTable.GetValue(0, intRow)
                    If oUserTable.GetByKey(strCode) Then
                        oUserTable.Code = strCode
                        oUserTable.Name = strECode
                        oUserTable.UserFields.Fields.Item("U_Z_AvgYear").Value = oGrid.DataTable.GetValue("U_Z_AvgYear", intRow)
                        oUserTable.UserFields.Fields.Item("U_Z_FrgnName").Value = oGrid.DataTable.GetValue("U_Z_FrgnName", intRow)
                        oUserTable.UserFields.Fields.Item("U_Z_DefAmt").Value = oGrid.DataTable.GetValue("U_Z_DefAmt", intRow)
                        oUserTable.UserFields.Fields.Item("U_Z_SOCI_BENE").Value = strESocial
                        oUserTable.UserFields.Fields.Item("U_Z_INCOM_TAX").Value = strETax
                        oUserTable.UserFields.Fields.Item("U_Z_Max").Value = oGrid.DataTable.GetValue("U_Z_Max", intRow)
                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_OffCycle")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            oUserTable.UserFields.Fields.Item("U_Z_OffCycle").Value = "Y"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_OffCycle").Value = "N"
                        End If
                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_EOS")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            oUserTable.UserFields.Fields.Item("U_Z_EOS").Value = "Y"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_EOS").Value = "N"
                        End If

                        OCHECKBOXCOLUMN = oGrid.Columns.Item("U_Z_AffDedu")
                        If OCHECKBOXCOLUMN.IsChecked(intRow) Then
                            oUserTable.UserFields.Fields.Item("U_Z_AffDedu").Value = "Y"
                        Else
                            oUserTable.UserFields.Fields.Item("U_Z_AffDedu").Value = "N"
                        End If
                        oUserTable.UserFields.Fields.Item("U_Z_EAR_GLACC").Value = (oGrid.DataTable.GetValue("U_Z_EAR_GLACC", intRow))
                        oUserTable.UserFields.Fields.Item("U_Z_DED_GLACC").Value = (oGrid.DataTable.GetValue("U_Z_DED_GLACC", intRow))
                        oUserTable.UserFields.Fields.Item("U_Z_BaiscPer").Value = (oGrid.DataTable.GetValue("U_Z_BaiscPer", intRow))
                        oUserTable.UserFields.Fields.Item("U_Z_PostType").Value = stPosttype
                        If oUserTable.Update() <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Committrans("Cancel")
                            Return False
                        Else
                            'If AddToUDT_Employee(oGrid.DataTable.GetValue(2, intRow).ToString.ToUpper(), oGrid.DataTable.GetValue("U_Z_Percentage", intRow), oGrid.DataTable.GetValue(4, intRow)) = False Then
                            '    Return False
                            'End If
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


    Private Function AddToUDT_Employee(ByVal aType As String, ByVal dblvalue1 As Double, ByVal GLAccount As String) As Boolean
        Dim strTable, strEmpId, strCode, strType As String
        Dim dblValue As Double
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oValidateRS, oTemp As SAPbobsCOM.Recordset
        oUserTable = oApplication.Company.UserTables.Item("Z_PAY1")
        oValidateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select * from [OHEM] order by EmpID ")
        strTable = "@Z_PAY1"
        strType = aType
        dblValue = dblvalue1
        Dim strQuery As String
        strQuery = "Update [@Z_PAY1] set U_Z_GLACC='" & GLAccount & "' where U_Z_EARN_TYPE='" & strType & "'"
        oValidateRS.DoQuery(strQuery)

        For intRow As Integer = 0 To oTemp.RecordCount - 1
            If strType <> "" Then
                strEmpId = oTemp.Fields.Item("empID").Value
                oValidateRS.DoQuery("Select * from [@Z_PAY1] where U_Z_EARN_TYPE='" & strType & "' and U_Z_EMPID='" & strEmpId & "'")
                If oValidateRS.RecordCount > 0 Then
                    strCode = oValidateRS.Fields.Item("Code").Value
                Else
                    strCode = ""
                End If
                dblValue = dblvalue1
                If strCode <> "" Then ' oUserTable.GetByKey(strCode) Then
                    'oUserTable.Code = strCode
                    'oUserTable.Name = strCode
                    'oUserTable.UserFields.Fields.Item("U_Z_EMPID").Value = strEmpId
                    'oUserTable.UserFields.Fields.Item("U_Z_EARN_TYPE").Value = strType
                    'Dim dblBasic As Double
                    'dblBasic = oTemp.Fields.Item("Salary").Value
                    'dblBasic = (oApplication.Utilities.getDocumentQuantity(oTemp.Fields.Item("Salary").Value))

                    'dblValue = (dblBasic * dblValue) / 100
                    ''       oUserTable.UserFields.Fields.Item("U_Z_EARN_VALUE").Value = dblValue
                    'oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = GLAccount
                    'If oUserTable.Update <> 0 Then
                    '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Return False
                    'End If
                Else
                    strCode = oApplication.Utilities.getMaxCode(strTable, "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode + "N"
                    oUserTable.UserFields.Fields.Item("U_Z_EMPID").Value = strEmpId
                    oUserTable.UserFields.Fields.Item("U_Z_EARN_TYPE").Value = strType
                    Dim dblBasic As Double
                    dblBasic = oTemp.Fields.Item("Salary").Value
                    dblBasic = (oApplication.Utilities.getDocumentQuantity(oTemp.Fields.Item("Salary").Value))
                    dblValue = (dblBasic * dblValue) / 100
                    oUserTable.UserFields.Fields.Item("U_Z_EARN_VALUE").Value = dblValue
                    oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = GLAccount
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End If
            oTemp.MoveNext()
        Next
        oUserTable = Nothing
        Return True
    End Function

#Region "Remove Row"
    Private Sub RemoveRow(ByVal intRow As Integer, ByVal agrid As SAPbouiCOM.Grid)
        Dim strCode, strname As String
        Dim otemprec As SAPbobsCOM.Recordset
        For intRow = 0 To agrid.DataTable.Rows.Count - 1
            If agrid.Rows.IsSelected(intRow) Then
                strCode = agrid.DataTable.GetValue("Code", intRow)
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If oApplication.Utilities.ValidateDeletionMaster(oGrid.DataTable.GetValue("Code", intRow), "Variable Earning") = False Then
                    Exit Sub
                End If

                oApplication.Utilities.ExecuteSQL(oTemp, "update [@Z_PAY_OEAR1] set  NAME =NAME +'_XD'  where Code='" & strCode & "'")
                agrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No row selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
#End Region


#Region "Validate Grid details"
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strECode, strECode1, strEname, strEname1 As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strECode = aGrid.DataTable.GetValue("Code", intRow)
            strEname = aGrid.DataTable.GetValue("Name", intRow)
            If strECode <> "" And strEname = "" Then
                oApplication.Utilities.Message("Allowance Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oGrid.Columns.Item("Name").Click(intRow)
                Return False
            End If
            If strECode = "" And strEname <> "" Then
                oApplication.Utilities.Message("Allowance Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oGrid.Columns.Item("Code").Click(intRow)
                Return False
            End If
            If oGrid.DataTable.GetValue("U_Z_EAR_GLACC", intRow) = "" Then
                oApplication.Utilities.Message("G/L Account Missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oGrid.Columns.Item("U_Z_EAR_GLACC").Click(intRow)
                Return False
            End If
            For intInnerLoop As Integer = intRow To aGrid.DataTable.Rows.Count - 1
                strECode1 = aGrid.DataTable.GetValue("Code", intInnerLoop)
                strEname1 = aGrid.DataTable.GetValue("Name", intInnerLoop)
                If strECode1 <> "" And strEname <> "" Then
                    If strECode1 <> "" And strEname1 = "" Then
                        oApplication.Utilities.Message("Name can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode1 = "" And strEname1 <> "" Then
                        oApplication.Utilities.Message("Code can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    If strECode = strECode1 And intRow <> intInnerLoop Then
                        oApplication.Utilities.Message("This Earning Code already exists. Code no : " & strECode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aGrid.Columns.Item(2).Click(intInnerLoop, , 1)
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function

#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_VariableEarning Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "5" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    'If oApplication.Utilities.ValidateDeletionMaster(oGrid.DataTable.GetValue("Code", pVal.Row), "Variable Earning") = False Then
                                    '    BubbleEvent = False
                                    '    Exit Sub
                                    'End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "5" And pVal.ColUID = "Code" And pVal.CharPressed <> 9 Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    If oApplication.Utilities.ValidateDeletionMaster(oGrid.DataTable.GetValue("Code", pVal.Row), "Variable Earning") = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "2" Then
                                    Committrans("Cancel")
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                ' oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "13" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    If validation(oGrid) = True Then
                                        AddtoUDT1(oForm)
                                    End If
                                End If
                                If pVal.ItemUID = "3" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    AddEmptyRow(oGrid)
                                End If
                                If pVal.ItemUID = "4" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    RemoveRow(pVal.Row, oGrid)
                                End If
                                If pVal.ItemUID = "6" Then
                                    oGrid = oForm.Items.Item("5").Specific
                                    For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                        If oGrid.Rows.IsSelected(intRow) Then
                                            Dim oObj As New clsAllowanceLeaveMapping
                                            oObj.LoadForm(oGrid.DataTable.GetValue("U_Z_CODE", intRow), oGrid.DataTable.GetValue("U_Z_NAME", intRow))
                                            Exit Sub
                                        End If
                                    Next
                                End If
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
                                        oForm.Update()
                                        If pVal.ItemUID = "5" Then
                                            oGrid = oForm.Items.Item("5").Specific
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            Try
                                                oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                            Catch ex As Exception
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
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_VariableEarning
                    LoadForm()
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = False Then
                        AddEmptyRow(oGrid)
                        oApplication.Utilities.assignMatrixLineno(oGrid, oForm)
                    End If

                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oGrid = oForm.Items.Item("5").Specific
                    If pVal.BeforeAction = True Then
                        RemoveRow(1, oGrid)
                        oApplication.Utilities.assignMatrixLineno(oGrid, oForm)
                        BubbleEvent = False
                        Exit Sub
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
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_Earning
                        oMenuobject = New clsEarning
                        oMenuobject.MenuEvent(pVal, BubbleEvent)
                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
