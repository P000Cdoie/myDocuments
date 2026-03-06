Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Friend Class frmMKTTRN0051
    Inherits System.Windows.Forms.Form
    '************************************************************************
    '(C) 2001 MIND, All rights reserved
    '
    'File Name          :   frmMKTTRN0051.frm
    'Function           :   Sales Order Entry
    'Created By         :   Manoj Kr. vaish
    'Created on         :   24 April 2007
    '----------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   26/05/2011
    'Modified to support MultiUnit functionality
    '----------------------------------------------------
    'Modified by    :   Prashant Rajpal
    'Modified ON    :   14/05/2012
    'Modified to resolve error
    'issue id : 10224351
    '-----------------------------------------------------------------------
    'MODIFIED BY       -  Prashant Rajpal
    'MODIFIED ON       -  23/02/2015
    'ISSUE ID          -  10685163 
    'ISSUE DESCRIPTION -  Declaration No column add in item master (only possible in ADD mode )
    '********************************************************************************************
    'CREATED BY       : Parveen Kumar
    'CREATED ON       : 09 JUN 2015
    'DESCRIPTION      : eMPro- Declaration No. in Export Invoice
    'AGAINST ISSUE ID : 10826755
    '-----------------------------------------------------------------------
    'MODIFIED BY       -  Prashant Rajpal
    'MODIFIED ON       -  11/12/2017
    'ISSUE ID          -  101391276 
    'ISSUE DESCRIPTION -  CHANGES FOR MTL SHARJAH ,VAT INCLUDED 
    '------------------------------------------------------------------------------

    Dim m_blnHelpFlag, m_blnCloseFlag As Boolean
    Dim rsdb As New ClsResultSetDB
    Dim mintFormIndex, intRow As Short
    Dim mstrCode As String
    Dim m_Item_Code As String
    Dim m_blnChangeFormFlg As Boolean
    Dim mvalid As Boolean
    Dim m_strSql, strsql As String
    Dim m_blnGetAmendmentDetails As Boolean
    Dim rsRefNo As New ClsResultSetDB
    Dim m_ItemDesc, m_custItemDesc As String
    Dim strdt As String
    Dim strpotype As String
    Dim strSOType As String
    Dim blnInvalidData As Boolean
    Dim blnValidCurrency As Boolean
    Dim mstrPrevAccountCode As String
    Dim mstrPrevRefNo As String
    Dim blnValidAmendDate As Boolean
    Dim ArrDispatchQty() As Double
    Dim dtSODate As Date
    Public mstrFormDetails As String
    Dim m_blnDateFlag As Boolean
    Private blnmsgbox As Boolean = False
    Private blnmsgbox1 As Boolean = False
    Dim blnCheckforSave As Boolean = False
    Dim blnCheck As Boolean = False
    Dim blnLeavetxt As Boolean = False
    Private Sub chkOpenSo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkOpenSo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                If ctlPerValue.Enabled = True Then
                    ctlPerValue.Focus()
                End If
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkOpenSo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOpenSo.Leave
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        intMaxLoop = SSPOentry.MaxRows
        With SSPOentry
            If chkOpenSo.CheckState = 1 Then
                For intLoopCounter = 1 To intMaxLoop
                    .Col = 1
                    .Col2 = 1
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .Value = System.Windows.Forms.CheckState.Checked
                    .Col = 5
                    .Col2 = 5
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .Text = CStr(0)
                Next
                .Col = 5
                .Col2 = 5
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
                'tO LOCK Open Item Flag
                .Col = 1
                .Col2 = 1
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            Else
                'tO UnLOCK Open Item Flag
                .Col = 1
                .Col2 = 1
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = False
                .BlockMode = False
            End If
        End With
    End Sub
    Private Sub cmbPOType_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPOType.Enter
        Call Dropdown(Me.cmbPOType.Handle.ToInt32)
    End Sub
    Private Sub cmbPOType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbPOType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            txtCreditTerms.Focus()
        End If
    End Sub
    Private Sub cmbPOType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbPOType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmbPOType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbPOType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        'Procedure to enbale cell when PO type is MRP-SPARES
        On Error GoTo ErrHandler
        If cmbPOType.Text = "MRP-SPARES" Then
            SSPOentry.BlockMode = True
            SSPOentry.Col = 19
            SSPOentry.Col2 = 19
            SSPOentry.ColHidden = False
            SSPOentry.TypeFloatMin = "0.0000"
            SSPOentry.Col = 20
            SSPOentry.Col2 = 20
            SSPOentry.ColHidden = False
            SSPOentry.BlockMode = False
            SSPOentry.Col = 21
            SSPOentry.Col2 = 21
            SSPOentry.ColHidden = False
            SSPOentry.TypeFloatMin = "0.0000"
        Else
            SSPOentry.BlockMode = True
            SSPOentry.Col = 19
            SSPOentry.Col2 = 19
            SSPOentry.ColHidden = True
            SSPOentry.Col = 20
            SSPOentry.Col2 = 20
            SSPOentry.ColHidden = True
            SSPOentry.Col = 21
            SSPOentry.Col2 = 21
            SSPOentry.ColHidden = True
            SSPOentry.BlockMode = False
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmdButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdButtons.ButtonClick
        Dim varExtraExciseDuty, varSalestax, varToolCost, varordqty, varItemCode, varRate, varCustSuppMaterial, varPkg, varExiseDuty, varSurchargeSalesTax As Object
        Dim varDespatchQty, varCustItemCode, varCustItemDesc, varOthers As Object
        Dim varDeleteFlag As Object
        Dim intLoop As Short
        Dim intMaxLoop As Short
        Dim arrMain() As String
        Dim arrDet() As String
        Dim strFormSQL As String
        Dim intOuterCount As Short
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strErrMsg As String
        Dim blnInvalidData As Boolean
        Dim intRow As Short
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim varItem As Object
        Dim vardrawing As Object
        blnInvalidData = False
        Dim counter As Short
        '10685163
        Dim VARDECLARATIONMSG As String = String.Empty
        Dim clsitemdeclarationno As New ClsResultSetDB
        '10685163
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD  'Add Record
                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.Color.White
                Me.lblIntSONoDes.Text = ""
                cmdHelp(0).Enabled = True
                txtConsCode.Enabled = True
                txtConsCode.BackColor = System.Drawing.Color.White
                cmdHelp(6).Enabled = True

                SSPOentry.MaxRows = 0
                With SSPOentry
                    .Col = 19
                    .Col2 = 19
                    .ColHidden = True
                    .Col = 20
                    .Col2 = 20
                    .ColHidden = True
                End With
                m_strSalesTaxType = ""
                txtCustomerCode.Focus()
                cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 1)
                cmdButtons.Enabled(3) = True
                frmSearch.Enabled = True
                optPartNo.Enabled = True
                optPartNo.Checked = True
                optItem.Enabled = True
                txtsearch.Enabled = True
                txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                mstrFormDetails = ""
                cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, False)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT  'Edit Record
                strsql = "Select Future_so From Cust_ord_hdr Where Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                strsql = strsql & " and cust_ref = '" & Trim(txtReferenceNo.Text) & "' and amendment_no = '"
                strsql = strsql & Trim(txtAmendmentNo.Text) & "' and Unit_code = '" & gstrUNITID & "'"
                rsSalesParameter.GetResult(strsql)
                If rsSalesParameter.GetValue("Future_so") = True Then
                    If MsgBox("This is Future SO [Authorised], Changes in this SO with Update the Deatls of this SO and Make it UnAuthorised. Would you like to Proceed...", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                        cmdButtons.Revert()
                        Exit Sub
                    End If
                End If
                SSPOentry.Enabled = True
                'change account Plug in
                With SSPOentry
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 5
                    .Col2 = 5
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End With
                cmdForms.Enabled = True
                txtCreditTerms.Enabled = True : txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(5).Enabled = True
                txtSTax.Enabled = True : txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(7).Enabled = True

                chkOpenSo.Enabled = True
                txtCustomerCode.Enabled = False
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(0).Enabled = False
                txtReferenceNo.Enabled = False
                txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(1).Enabled = False
                txtAmendmentNo.Enabled = False
                txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(3).Enabled = False
                txtCurrencyType.Enabled = False
                cmdHelp(2).Enabled = False
                cmbPOType.Enabled = False
                cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtAmendReason.Enabled = True
                cmdchangetype.Enabled = True
                If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                    DTAmendmentDate.Enabled = False
                    txtAmendReason.Enabled = False
                    txtAmendReason.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmbPOType.Enabled = True
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtCurrencyType.Enabled = True
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(2).Enabled = True
                    ctlPerValue.Enabled = True
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    With SSPOentry
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                        intMaxLoop = .MaxRows
                        For intLoop = 1 To intMaxLoop
                            .Col = 2
                            .Col2 = 4
                            .Row = intLoop
                            .Row2 = intLoop
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                            rsSalesParameter.GetResult("select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                            If rsSalesParameter.GetValue("ItemRateLink") = True Then
                                If checkforitemRate(intLoop) = False Then
                                    .Col = 6
                                    .Col2 = .MaxCols
                                    .Row = intLoop
                                    .Row2 = intLoop
                                    .BlockMode = True
                                    .Lock = True
                                    .BlockMode = False
                                Else
                                    .Col = 6
                                    .Col2 = .MaxCols
                                    .Row = intLoop
                                    .Row2 = intLoop
                                    .BlockMode = True
                                    .Lock = False
                                    .BlockMode = False
                                End If
                                .Col = 6
                                .Col2 = .MaxCols
                                .Row = intLoop
                                .Row2 = intLoop
                                .BlockMode = True
                                .Lock = False
                                .BlockMode = False
                            End If
                        Next
                    End With
                Else
                    DTAmendmentDate.Enabled = True
                    txtAmendReason.Enabled = True
                    txtAmendReason.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmbPOType.Enabled = False
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCurrencyType.Enabled = False
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(2).Enabled = False
                    With SSPOentry
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                        .Col = 1
                        .Col2 = 1
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                End If
                DTDate.Enabled = True
                DTEffectiveDate.Enabled = True
                DTValidDate.Enabled = True
                With SSPOentry
                    .Row = 1
                    .Col = 2
                    .Action = 0
                End With
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE  'Delete Record
                enmValue = ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    ' deleting the record from cust_ord_hdr table
                    strsql = "delete cust_ord_hdr where Account_Code='"
                    strsql = strsql & txtCustomerCode.Text & "' and Cust_Ref='"
                    strsql = strsql & txtReferenceNo.Text & "' and Amendment_No='"
                    strsql = strsql & txtAmendmentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                    ' deleting the record from cust_ord_dtl table
                    Call DeleteRow()
                    mP_Connection.Execute("DELETE FROM Forms_dtl WHERE DOC_TYPE=9998 AND PO_NO='" & Trim(txtReferenceNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendmentNo.Text) & "' and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    cmdButtons.Revert()
                    cmdButtons.Enabled(0) = True
                    cmdButtons.Enabled(1) = False
                    cmdButtons.Enabled(2) = False
                    cmdButtons.Enabled(5) = False
                    Call EnableControls(False, Me, True)
                    txtCustomerCode.Enabled = True
                    txtCustomerCode.BackColor = System.Drawing.Color.White
                    txtConsCode.Enabled = True
                    txtConsCode.BackColor = System.Drawing.Color.White
                    cmdHelp(0).Enabled = True
                    cmdHelp(6).Enabled = True
                    SSPOentry.MaxRows = 0
                    txtCustomerCode.Focus()
                Else
                    txtCustomerCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE  'Save Record
                If ValidRecord() = False Then Exit Sub
                counter = SSPOentry.MaxRows

                If counter = 0 Then
                    MessageBox.Show("Item Details Not Entered.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If

                For intRow = 1 To counter 'Checking if all details have been entered correctly
                    If SSPOentry.MaxRows = 1 Or intRow <> SSPOentry.MaxRows Then
                        If Not ValidRowData(intRow, 0) Then
                            gblnCancelUnload = True : gblnFormAddEdit = True
                            Exit Sub
                        End If
                    Else
                        varItem = Nothing
                        Call SSPOentry.GetText(2, SSPOentry.MaxRows, varItem)
                        vardrawing = Nothing
                        Call SSPOentry.GetText(2, SSPOentry.MaxRows, vardrawing)
                        If ((Len(Trim(varItem)) <= 0) And (Len(Trim(vardrawing)) <= 0)) And SSPOentry.MaxRows > 1 Then
                            SSPOentry.MaxRows = SSPOentry.MaxRows - 1
                        Else
                            If Not ValidRowData(intRow, 0) Then
                                gblnCancelUnload = True : gblnFormAddEdit = True
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD  'Check for mode when save button was clicked
                        ReDim ArrDispatchQty(SSPOentry.MaxRows - 1)
                        'Call DeleteRow
                        arrMain = Split(mstrFormDetails, "^")
                        strFormSQL = ""
                        For intOuterCount = 0 To UBound(arrMain) - 1
                            arrDet = Split(arrMain(intOuterCount), "|")
                            strFormSQL = strFormSQL & "INSERT INTO Forms_dtl(DOC_TYPE, PO_NO, AMENDMENT_NO, SERIAL_NO, FORM_TYPE, FORM_NO, Account_code, Unit_Code)"
                            strFormSQL = strFormSQL & " VALUES(9998,'" & txtReferenceNo.Text & "','" & Trim(txtAmendmentNo.Text) & "','" & intOuterCount & "','" & arrDet(0) & "','" & arrDet(1) & "', '" & Trim(txtCustomerCode.Text) & "', '" & gstrUNITID & "')" & vbCrLf
                        Next

                        '10685163
                        'VARDECLARATIONMSG = ""
                        'If UCase(cmbPOType.Text) = "EXPORT" Then
                        '    rsSalesParameter.GetResult("Select DeclarationNo_Exportinvoice from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                        '    If rsSalesParameter.GetValue("DeclarationNo_Exportinvoice") = True Then
                        '        For intRow = 1 To SSPOentry.MaxRows
                        '            varItemCode = Nothing
                        '            Call SSPOentry.GetText(4, intRow, varItemCode)
                        '            strsql = "SELECT DECLARATIONNO FROM  ITEM_MST WHERE ITEM_CODE = '" & Trim(varItemCode) & "' AND UNIT_CODE = '" & gstrUNITID & "' AND ISNULL(DECLARATIONNO,'')<=0 "
                        '            clsitemdeclarationno = New ClsResultSetDB
                        '            If clsitemdeclarationno.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsitemdeclarationno.GetNoRows > 0 Then
                        '                clsitemdeclarationno.MoveFirst()
                        '                VARDECLARATIONMSG = VARDECLARATIONMSG + varItemCode + " ," + vbCrLf
                        '                clsitemdeclarationno.ResultSetClose()
                        '            End If
                        '        Next
                        '    End If
                        '    If VARDECLARATIONMSG.ToString <> "" Then
                        '        MsgBox("Declaration No. Not defined for item code: " + VARDECLARATIONMSG, MsgBoxStyle.Information, ResolveResString(100))
                        '        Exit Sub
                        '    End If
                        'End If
                        '10685163
                        '10826755--Starts Here
                        VARDECLARATIONMSG = ""
                        If UCase(cmbPOType.Text) = "EXPORT" Then
                            rsSalesParameter.GetResult("Select DeclarationNo_Exportinvoice from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                            If rsSalesParameter.GetValue("DeclarationNo_Exportinvoice") = True Then
                                For intRow = 1 To SSPOentry.MaxRows
                                    varItemCode = Nothing
                                    varCustItemCode = Nothing
                                    Call SSPOentry.GetText(2, intRow, varCustItemCode)
                                    Call SSPOentry.GetText(4, intRow, varItemCode)
                                    strsql = "SELECT Decl_No FROM CUSTITEM_MST WHERE ITEM_CODE = '" & Trim(varItemCode) & "' AND UNIT_CODE = '" & gstrUNITID & "' AND Account_code = '" & Trim(txtCustomerCode.Text) & "' AND Cust_Drgno = '" & Trim(varCustItemCode) & "' AND Decl_No='' "
                                    clsitemdeclarationno = New ClsResultSetDB
                                    If clsitemdeclarationno.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsitemdeclarationno.GetNoRows > 0 Then
                                        clsitemdeclarationno.MoveFirst()
                                        VARDECLARATIONMSG = VARDECLARATIONMSG + varItemCode + " ," + varCustItemCode + " ," + vbCrLf
                                        clsitemdeclarationno.ResultSetClose()
                                    End If
                                Next
                            End If
                            If VARDECLARATIONMSG.ToString <> "" Then
                                MsgBox("Declaration No. has Not defined for item code,Drg No: " + VARDECLARATIONMSG, MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                        '10826755--Ends Here
                        With mP_Connection
                            .BeginTrans()
                            Call InsertRowCustOrdHdr()
                            Call InsertRow()
                            .CommitTrans()
                            rsSalesParameter.GetResult("Select AppendSOItem from sales_parameter where Unit_Code = '" & gstrUNITID & "'")
                            If rsSalesParameter.GetValue("AppendSOItem") = True Then
                                Call InsertPreviousSODetails(Trim(txtCustomerCode.Text), Trim(txtReferenceNo.Text), Trim(txtAmendmentNo.Text), Trim(lblIntSONoDes.Text), CShort(Trim(lblRevisionNo.Text)))
                            End If
                        End With
                        cmdButtons.Revert()
                        cmdButtons.Enabled(0) = True
                        cmdButtons.Enabled(1) = False
                        cmdButtons.Enabled(2) = False
                        cmdButtons.Enabled(5) = False
                        Call EnableControls(False, Me, False)
                        txtCustomerCode.Enabled = True
                        txtCustomerCode.BackColor = System.Drawing.Color.White
                        txtConsCode.Enabled = True
                        txtConsCode.BackColor = System.Drawing.Color.White
                        cmdHelp(0).Enabled = True
                        cmdHelp(6).Enabled = True

                        SSPOentry.MaxRows = 0
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT  ' in case of edit, update the record in the cust_ord_hdr table
                        strFormSQL = "DELETE FROM Forms_dtl WHERE DOC_TYPE=9998 AND Account_code='" & Trim(txtCustomerCode.Text) & "' and PO_NO='" & Trim(txtReferenceNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendmentNo.Text) & "' and Unit_Code = '" & gstrUNITID & "'" & vbCrLf
                        If Len(mstrFormDetails) > 0 Then
                            arrMain = Split(mstrFormDetails, "^")
                            For intOuterCount = 0 To UBound(arrMain) - 1
                                arrDet = Split(arrMain(intOuterCount), "|")
                                strFormSQL = strFormSQL & "INSERT INTO Forms_dtl(DOC_TYPE, PO_NO, AMENDMENT_NO, SERIAL_NO, FORM_TYPE, FORM_NO, Account_code,Unit_Code)"
                                strFormSQL = strFormSQL & " VALUES(9998,'" & txtReferenceNo.Text & "','" & txtAmendmentNo.Text & "','" & intOuterCount & "','" & arrDet(0) & "','" & arrDet(1) & "', '" & Trim(txtCustomerCode.Text) & "','" & gstrUNITID & "')" & vbCrLf
                            Next
                        End If
                        VARDECLARATIONMSG = ""
                        If UCase(cmbPOType.Text) = "EXPORT" Then
                            rsSalesParameter.GetResult("Select DeclarationNo_Exportinvoice from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                            If rsSalesParameter.GetValue("DeclarationNo_Exportinvoice") = True Then
                                For intRow = 1 To SSPOentry.MaxRows
                                    varItemCode = Nothing
                                    varCustItemCode = Nothing
                                    Call SSPOentry.GetText(2, intRow, varCustItemCode)
                                    Call SSPOentry.GetText(4, intRow, varItemCode)
                                    strsql = "SELECT Decl_No FROM CUSTITEM_MST WHERE ITEM_CODE = '" & Trim(varItemCode) & "' AND UNIT_CODE = '" & gstrUNITID & "' AND Account_code = '" & Trim(txtCustomerCode.Text) & "' AND Cust_Drgno = '" & Trim(varCustItemCode) & "' AND Decl_No='' "
                                    clsitemdeclarationno = New ClsResultSetDB
                                    If clsitemdeclarationno.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsitemdeclarationno.GetNoRows > 0 Then
                                        clsitemdeclarationno.MoveFirst()
                                        VARDECLARATIONMSG = VARDECLARATIONMSG + varItemCode + " ," + varCustItemCode + " ," + vbCrLf
                                        clsitemdeclarationno.ResultSetClose()
                                    End If
                                Next
                            End If
                            If VARDECLARATIONMSG.ToString <> "" Then
                                MsgBox("Declaration No. has Not defined for item code,Drg No: " + VARDECLARATIONMSG, MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                        With mP_Connection
                            'To Confirm the deletion of marked rows
                            If ConfirmDeletion() = False Then Exit Sub
                            .BeginTrans()
                            Call UpdateRow()
                            'delete the record in the cust_ord_dtl table
                            Call DeleteRow()
                            ' Add all the records in the grid to the table cust_ord_dtl which are not marked for deletion
                            Call InsertRow()
                            If Len(strFormSQL) > 0 Then .Execute(strFormSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .CommitTrans()
                        End With
                        cmdButtons.Revert()
                        cmdButtons.Enabled(0) = True
                        cmdButtons.Enabled(1) = False
                        cmdButtons.Enabled(2) = False
                        cmdButtons.Enabled(5) = False
                        Call EnableControls(False, Me, False)
                        txtCustomerCode.Enabled = True
                        txtConsCode.Enabled = True
                        txtCustomerCode.BackColor = System.Drawing.Color.White
                        txtConsCode.BackColor = System.Drawing.Color.White
                        cmdHelp(0).Enabled = True
                        cmdHelp(6).Enabled = True
                        SSPOentry.MaxRows = 0
                End Select
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                        MsgBox("SO Successfully updated with Internal SO No " & lblIntSONoDes.Text, MsgBoxStyle.Information, ResolveResString(100))
                    Else
                        Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    End If
                End If
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtConsCode.Enabled = True
                txtConsCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelp(0).Enabled = True
                cmdHelp(6).Enabled = True
                frmSearch.Enabled = True
                optPartNo.Enabled = True
                optPartNo.Checked = True
                optItem.Enabled = True
                txtsearch.Enabled = True
                txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                mstrFormDetails = ""
                txtCustomerCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Call EnableControls(False, Me, True)
                    frmSearch.Enabled = True
                    optPartNo.Enabled = True
                    optPartNo.Checked = True
                    optItem.Enabled = True
                    txtsearch.Enabled = True
                    txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtCustomerCode.Enabled = True
                    txtCustomerCode.BackColor = System.Drawing.Color.White
                    cmdHelp(0).Enabled = True
                    txtConsCode.Enabled = True
                    txtConsCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(6).Enabled = True

                    SSPOentry.MaxRows = 0
                    txtCustomerCode.Focus()
                    gblnCancelUnload = False : gblnFormAddEdit = False
                    cmdButtons.Revert()
                    cmdButtons.Enabled(0) = True
                    cmdButtons.Enabled(1) = False
                    cmdButtons.Enabled(2) = False
                    cmdButtons.Enabled(5) = False
                    Call RefreshForm()
                    If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    End If
                Else
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                Call PrintToReport()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Call CLEARVAR()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        On Error GoTo ErrHandler
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        Dim strsql As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            strsql = "select a.*,b.* from cust_ord_hdr a,cust_ord_dtl b where a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.unit_Code=b.unit_Code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.unit_Code= '" & gstrUNITID & "'"
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strsql = "select * from cust_ord_hdr  where Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "' and unit_Code= '" & gstrUNITID & "'"
        End If
        m_pstrSql = strsql
        rssalesTerms = New ClsResultSetDB
        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PY' and unit_Code= '" & gstrUNITID & "'"
        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rssalesTerms.GetNoRows > 0 Then
            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PR' and unit_Code= '" & gstrUNITID & "'"
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows > 0 Then
                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PK' and unit_Code= '" & gstrUNITID & "'"
                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalesTerms.GetNoRows > 0 Then
                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='FR' and unit_Code= '" & gstrUNITID & "'"
                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalesTerms.GetNoRows > 0 Then
                        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='TR' and unit_Code= '" & gstrUNITID & "'"
                        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rssalesTerms.GetNoRows > 0 Then
                            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='OC' and unit_Code= '" & gstrUNITID & "'"
                            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rssalesTerms.GetNoRows > 0 Then
                                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='MO' and unit_Code= '" & gstrUNITID & "'"
                                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                If rssalesTerms.GetNoRows > 0 Then
                                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='DL' and unit_Code= '" & gstrUNITID & "'"
                                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rssalesTerms.GetNoRows > 0 Then
                                        Select Case cmdButtons.Mode
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                                frmMKTTRN0010.formload("MODE_ADD")
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                                                frmMKTTRN0010.formload("MODE_EDIT")
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                                                frmMKTTRN0010.formload("MODE_VIEW")
                                        End Select
                                        Call frmMKTTRN0010.Show()
                                    Else
                                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                        cmdchangetype.Focus()
                                        Exit Sub
                                    End If
                                Else
                                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    cmdchangetype.Focus()
                                    Exit Sub
                                End If
                            Else
                                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                cmdchangetype.Focus()
                                Exit Sub
                            End If
                        Else
                            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdchangetype.Focus()
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        cmdchangetype.Focus()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    cmdchangetype.Focus()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                cmdchangetype.Focus()
                Exit Sub
            End If
        Else
            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmdchangetype.Focus()
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdchangetype_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdchangetype.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            System.Windows.Forms.SendKeys.SendWait("{TAB}")
        End If
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim varRetVal As Object
        Dim strhelp1() As String
        Dim strMessage As String
        On Error GoTo ErrHandler
        blnCheckforSave = False
        Dim strAmend, strString As String
        Select Case Index
            '*****Customer Code Help
            Case 0
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Customer_code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code = b.Customer_code and a.Unit_Code = b.Unit_Code AND ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", , , , , , " a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code = b.Customer_code and a.Unit_Code = b.Unit_Code AND ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", , , , , , " a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                            .Focus()
                        End With
                End Select
            Case 1
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtReferenceNo
                            If Len(.Text) = 0 Then
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & " ,isnull(InternalSONo,'') as InternalSoNo from cust_ord_hdr where Account_Code='" & Me.txtCustomerCode.Text & "' and Active_Flag='A' and Unit_Code = '" & gstrUNITID & "'")
                            Else
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & ",isnull(InternalSONo,'') as InternalSoNo from cust_ord_hdr where Account_Code='" & Me.txtCustomerCode.Text & "' and Active_Flag='A' and cust_ref= '" & Me.txtReferenceNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")
                            End If
                            If Not (UBound(strhelp1) = -1) Then
                                If (Len(strhelp1(0)) >= 1) And strhelp1(0) = "0" Then
                                    If (Len(LTrim(RTrim(Me.txtReferenceNo.Text))) > 0) Then
                                        strMessage = "No job Order  are Defined with Prefix [" & Me.txtReferenceNo.Text & "]"
                                        strMessage = strMessage & vbCrLf & "To View the list, Clear the Text and try Again."
                                        MsgBox(strMessage, MsgBoxStyle.Information, ResolveResString(100))
                                    Else
                                        MsgBox("No job order are Defined", MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                    Me.txtReferenceNo.Focus()
                                    Exit Sub
                                Else
                                    Me.txtReferenceNo.Text = strhelp1(0)
                                    blnCheckforSave = True
                                End If
                            End If
                            .Focus()
                        End With

                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtReferenceNo
                            If Len(.Text) = 0 Then
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & " ,InternalSONo from cust_ord_hdr where Account_Code='" & Me.txtCustomerCode.Text & "'  and Amendment_No ='' and Unit_Code = '" & gstrUNITID & "'")
                            Else
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select Cust_Ref,'Order_date' =  " & DateColumnNameInShowList("Order_date") & " ,InternalSONo from cust_ord_hdr where Account_Code='" & Me.txtCustomerCode.Text & "' and Amendment_No ='' and cust_ref= '" & Me.txtReferenceNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")
                            End If
                            If Not (UBound(strhelp1) = -1) Then
                                If (Len(strhelp1(0)) >= 1) And strhelp1(0) = "0" Then
                                    If (Len(LTrim(RTrim(Me.txtReferenceNo.Text))) > 0) Then
                                        strMessage = "No Reference No. are Defined with Prefix [" & Me.txtReferenceNo.Text & "]"
                                        strMessage = strMessage & vbCrLf & "To View the list, Clear the Text and try Again."
                                        MsgBox(strMessage, MsgBoxStyle.Information, ResolveResString(100))
                                    Else
                                        MsgBox("No Reference No. are Defined", MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                    Me.txtReferenceNo.Focus()
                                    Exit Sub
                                Else
                                    Me.txtReferenceNo.Text = strhelp1(0)
                                End If
                            End If
                            .Focus()
                        End With
                End Select
                '******Currency Code Help
            Case 2
                With txtCurrencyType
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
                '******Amendment No Help
            Case 3
                If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    'Changed by Arshad to include Amendment Date in Help
                    With txtAmendmentNo
                        strString = txtAmendmentNo.Text & "%"
                        If txtAmendmentNo.Text <> "" Then strAmend = " where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No like '" & strString & "' and Unit_Code = '" & gstrUNITID & "'" Else strAmend = " where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                        varRetVal = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Amendment_No,'Amendment_date'= " & DateColumnNameInShowList("Amendment_date") & ", Cust_Ref  FROM cust_ord_hdr " & strAmend & " ", "List of All Amendments", 1)
                        If UBound(varRetVal) = "-1" Then Exit Sub
                        If varRetVal(0) = "0" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal(0)
                        End If
                        .Focus()
                    End With
                End If
                'SHARJAH CHANGES 01 JAN 2018
            Case 7
                With txtSTax
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST','VAT') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With

            Case 5
                With txtCreditTerms
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", "and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", "and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 6
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtConsCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Customer_code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtConsCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code = b.Customer_code and a.Unit_Code = b.Unit_Code AND ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", , , , , , " a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code = b.Customer_code and a.Unit_Code = b.Unit_Code AND ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", , , , , , " a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                            .Focus()
                        End With
                End Select

        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdHelp_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdHelp.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Select Case Index
            Case 1
                m_blnHelpFlag = True
            Case 3
                m_blnHelpFlag = True
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0051_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0051_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If KeyAscii = System.Windows.Forms.Keys.Escape Then
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub SSPOentry_Advance(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles SSPOentry.Advance
        On Error GoTo ErrHandler
        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call ADDRow()
        Else
        End If
        Application.DoEvents()
        With SSPOentry
            .Col = 1
            .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : If .Enabled Then .Focus()
        End With
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub SSPOentry_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SSPOentry.ButtonClicked
        On Error GoTo ErrHandler
        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParametere As New ClsResultSetDB
        Dim varHelpItem As Object
        Dim strSOEntry() As String
        Dim strtest As String
        rsSalesParametere.GetResult("Select ItemRateLink from Sales_parameter where Unit_Code = '" & gstrUNITID & "'")
        If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for currency type check
        If e.col = 3 Then
            If txtAmendmentNo.Enabled = False Then
                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                    strtest = "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b where a.Account_Code='" & txtCustomerCode.Text & "' and a.Unit_Code = '" & gstrUNITID & "' and a.Unit_Code = b.Unit_Code and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code and i1.Unit_Code = '" & gstrUNITID & "') and datediff(mm,convert(datetime,'" & Format(DTDate.Value, "dd mmm yyyy") & "'),convert(datetime,b.DateFrom))<=0 and CustVend_Flg = 'C'"
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b, Item_MST as C where A.Item_code=C.Item_Code and A.Unit_Code = B.Unit_Code and A.Unit_Code = C.Unit_Code and C.Status='A'  AND a.Active=1 and C.Hold_Flag=0 and A.Unit_Code = '" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code and i1.Unit_Code = '" & gstrUNITID & "') and datediff(mm,'" & FormatDateTime(DTDate.Value, vbLongDate) & "',b.DateFrom)<=0 and CustVend_Flg = 'C'")
                Else
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0  AND a.Active=1 and Account_Code='" & txtCustomerCode.Text & "' and A.Unit_Code = B.Unit_Code and A.Unit_Code = '" & gstrUNITID & "'")
                End If
            Else
                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0  AND a.Active=1 and Account_Code='" & txtCustomerCode.Text & "' and A.Unit_Code = B.Unit_Code and A.Unit_Code = '" & gstrUNITID & "'")
            End If
            If UBound(strSOEntry) <= 0 Then Exit Sub
            If strSOEntry(0) = "0" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            Else
                Call SSPOentry.SetText(2, SSPOentry.ActiveRow, strSOEntry(0))
                Call SSPOentry.SetText(4, SSPOentry.ActiveRow, strSOEntry(1))
                lblCustPartDesc.Text = strSOEntry(2)
            End If
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Dim rsitem As ClsResultSetDB
                If Len(Trim(m_Item_Code)) > 0 Then
                    m_strSql = "SElect * from Cust_ord_dtl where Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1 and Unit_Code = '" & gstrUNITID & "'"
                    rsitem = New ClsResultSetDB
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
                        If rsitem.GetValue("OpenSO") = False Then
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = False : SSPOentry.BlockMode = False
                        Else
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = True : SSPOentry.BlockMode = False
                        End If
                        Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                        Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsitem.GetValue("Item_Code "))
                        Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsitem.GetValue("Order_Qty"))
                        Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                        Call SSPOentry.SetText(6, SSPOentry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                        Call SSPOentry.SetText(17, SSPOentry.MaxRows, rsitem.GetValue("Despatch_Qty"))
                    Else
                        '****Same Data is in keyDown events
                        If txtAmendmentNo.Enabled = False Then
                            If Len(Trim(m_Item_Code)) > 0 Then
                                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                                    m_strSql = " select * from ITemRate_Mst where and Unit_Code = '" & gstrUNITID & "' and  Serial_No = (select max(serial_no) from itemrate_mst where Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & DTDate.Value & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C' and Unit_Code = '" & gstrUNITID & "')"
                                    rsitem = New ClsResultSetDB
                                    rsitem.GetResult(m_strSql)
                                    If rsitem.GetNoRows > 0 Then
                                        'change account Plug in
                                        If chkOpenSo.Checked = False Then
                                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                                        Else
                                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                                        End If
                                        Call SSPOentry.SetText(2, SSPOentry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                        Call SSPOentry.SetText(4, SSPOentry.MaxRows, m_Item_Code)
                                        Call SSPOentry.SetText(6, SSPOentry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                                        Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                                        If UCase(Trim(cmbPOType.Text)) <> "JOB WORK" Then
                                            If SSPOentry.MaxRows < 1 Then
                                                Call ADDRow()
                                            End If
                                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 2 : SSPOentry.Focus()
                                        End If
                                        If rsitem.GetValue("Edit_flg") = False Then
                                            SSPOentry.Col = 6 : SSPOentry.Col2 = SSPOentry.MaxCols : SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.BlockMode = True : SSPOentry.Lock = True : SSPOentry.BlockMode = False
                                        Else
                                            SSPOentry.Col = 6 : SSPOentry.Col2 = SSPOentry.MaxCols : SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.BlockMode = True : SSPOentry.Lock = False : SSPOentry.BlockMode = False
                                        End If
                                    End If
                                Else
                                    If chkOpenSo.Checked = False Then
                                        SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                                    Else
                                        SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                                    End If
                                    Call SSPOentry.SetText(2, SSPOentry.MaxRows, varHelpItem)
                                    Call SSPOentry.SetText(4, SSPOentry.MaxRows, m_Item_Code)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        With SSPOentry
            If .Col <> 1 Then
                Call ssSetFocus(.Row, 2)
            End If
        End With
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub SSPOentry_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SSPOentry.ClickEvent
        On Error GoTo ErrHandler
        Dim vartext As Object
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        If chkmultipleitem() = True And SSPOentry.ActiveCol = 4 Then
            Call SetCellTypeCombo(SSPOentry.ActiveRow)
        End If
        If e.col = 1 And e.row <> 0 Then
            Call ssSetFocus(e.row, 1)
            Exit Sub
        End If
        If e.col = 0 And e.row <> 0 Then
            vartext = Nothing
            Call SSPOentry.GetText(0, e.row, vartext)
            If vartext = "*" Then
                Call SSPOentry.SetText(0, e.row, "")
                With SSPOentry
                    .BlockMode = True
                    .Row = e.row
                    .Row2 = e.row
                    .Col = 1
                    .Col2 = .MaxCols
                    If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        .Lock = False
                    End If
                    .ForeColor = Color.Black
                    .BlockMode = False
                    .Row = e.row : .Row2 = e.row : .Col = 2 : .Col2 = 4 : .BlockMode = True : .Lock = True : .BlockMode = False
                    .Row = e.row : .Col = 1
                    If .Value = True Then
                        .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = True : .BlockMode = False
                    End If
                    varDrgNo = Nothing
                    varItemCode = Nothing
                    Call .GetText(2, e.row, varDrgNo)
                    Call .GetText(4, e.row, varItemCode)
                    If (Len(Trim(varDrgNo)) > 0) And Len(Trim(varItemCode)) > 0 Then
                        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                        If rsSalesParameter.GetValue("ItemRateLink") = True Then
                            If checkforitemRate(CInt(e.row)) = False Then
                                .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = True : .BlockMode = False
                            Else
                                .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                            End If
                        Else
                            .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                        End If
                    Else
                        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                        End If
                    End If
                End With
            Else
                Call SSPOentry.SetText(0, e.row, "*")
                With SSPOentry
                    .BlockMode = True
                    .Row = e.row
                    .Row2 = e.row
                    .Col = 1
                    .Col2 = .MaxCols
                    .ForeColor = Color.Red
                    .Lock = True
                    .BlockMode = False
                End With
            End If
            '        End If
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                cmdButtons.Enabled(1) = False
                cmdButtons.Enabled(2) = False
                cmdButtons.Enabled(5) = False
            End If
        End If
    End Sub
    Private Sub txtAmendmentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If mvalid = False Then
                If txtAmendReason.Enabled Then
                    txtAmendReason.Focus()
                Else
                    cmdButtons.Focus()
                End If
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendmentNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtAmendReason_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendReason.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If DTDate.Enabled = True Then
                DTDate.Focus()
            Else
                If DTAmendmentDate.Enabled = True Then
                    DTAmendmentDate.Focus()
                Else
                    If DTEffectiveDate.Enabled Then
                        DTEffectiveDate.Focus()
                    Else
                        cmdButtons.Focus()
                    End If
                End If
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendReason_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendReason.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCreditTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.TextChanged
        Call FillLabel("CREDIT")
    End Sub
    Private Sub txtCreditTerms_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCreditTerms.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case 112
                Call cmdHelp_Click(cmdHelp.Item(5), New System.EventArgs())
        End Select
    End Sub
    Private Sub txtCreditTerms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCreditTerms.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                chkOpenSo.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCreditTerms_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.Leave
        If Len(Trim(txtCreditTerms.Text)) <> 0 Then
            m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where crtrm_TermID = '" & txtCreditTerms.Text & "' and Unit_Code = '" & gstrUNITID & "'"
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("CREDIT")
            Else
                cmdButtons.Focus()
                MsgBox("Entered Credit Term does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtCreditTerms.Text = ""
                txtCreditTerms.Focus()
            End If
        End If
    End Sub
    Private Sub txtCurrencyType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCurrencyType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCurrencyType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                If Len(Trim(txtCurrencyType.Text)) > 0 Then
                    Call txtCurrencyType_Validating(txtCurrencyType, New System.ComponentModel.CancelEventArgs(False))
                    If blnValidCurrency = False Then
                        txtCurrencyType.Focus()
                    Else
                        cmbPOType.Focus()
                    End If
                Else
                    cmbPOType.Focus()
                End If
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCurrencyType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCurrencyType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsdb As New ClsResultSetDB
        If blnmsgbox = True Then
            Exit Sub
        End If
        If Trim(txtCurrencyType.Text) = "" Or Len(Trim(txtCurrencyType.Text)) = 0 Then
            blnmsgbox = True
            GoTo EventExitSub
        End If
        blnmsgbox = False
        m_strSql = "Select * from Currency_mst where Currency_code = '" & txtCurrencyType.Text & "' and Unit_Code = '" & gstrUNITID & "'"
        Call rsdb.GetResult(m_strSql)
        If rsdb.GetNoRows = 0 Then
            'MsgBox "Currency not exist in currency master"
            blnmsgbox = True
            Call ConfirmWindow(10144, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidCurrency = False
            Cancel = True
            GoTo EventExitSub
        Else
            blnmsgbox = False
            Call SSMaxLength()
        End If
        blnValidCurrency = True
        rsdb.ResultSetClose()
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
    End Sub
    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        On Error GoTo ErrHandler
        mvalid = False
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            Select Case cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    m_strSql = "Select * from cust_Ord_hdr where Account_code ='" & Trim(txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'"
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    m_strSql = "Select * from Customer_Mst where Customer_code ='" & Trim(txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            End Select
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(1).Enabled = True
                Call FillLabel("CUSTOMER")
            Else
                cmdButtons.Focus()
                MsgBox("Customer Code does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtCustomerCode.Text = ""
                txtCustomerCode.Focus()
                mvalid = True
            End If
        End If
        m_blnCloseFlag = False
        m_blnHelpFlag = False
        If StrComp(Trim(txtCustomerCode.Text), mstrPrevAccountCode, CompareMethod.Text) <> 0 Then
            Call CLEARVAR()
        End If
        'mvalid = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo ErrHandler
        Call FillLabel("CUSTOMER")
        Call FillLabel("CURRENCY")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            m_strSql = "Select * from cust_Ord_hdr where account_code ='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(1).Enabled = True
            Else
                txtReferenceNo.Enabled = False
                txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(1).Enabled = False
            End If
        Else
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                lblCustDesc.Text = ""
                txtReferenceNo.Text = ""
                txtAmendmentNo.Text = ""
                cmdButtons.Enabled(1) = False
                cmdButtons.Enabled(2) = False
                cmdButtons.Enabled(5) = False
            End If
        End If
        SSPOentry.MaxRows = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.Enter
        mstrPrevAccountCode = Trim(txtCustomerCode.Text)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Leave(txtCustomerCode, New System.EventArgs())
            If mvalid = False Then
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus() Else cmdButtons.Focus()
            Else
                txtCustomerCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB
        With SSPOentry
            .Col = 19
            .Col2 = 19
            .ColHidden = True
            .Col = 20
            .Col2 = 20
            .ColHidden = True
        End With
        If m_blnCloseFlag = True Then
            m_blnCloseFlag = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            m_strSql = "Select * from Customer_mst where customer_Code='" & Trim(txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtCustomerCode.Text = ""
                txtReferenceNo.Text = ""
                Cancel = True
                txtCustomerCode.Focus()
                GoTo EventExitSub
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
        On Error GoTo ErrHandler
        SSPOentry.MaxRows = 0
        txtAmendmentNo.Text = ""
        txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                txtAmendmentNo.Text = ""
                txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(3).Enabled = False
                cmdButtons.Enabled(1) = False
                cmdButtons.Enabled(2) = False
                cmdButtons.Enabled(5) = False
            End If
        End If
        If m_blnHelpFlag = True Then
            m_blnHelpFlag = False
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.Enter
        mstrPrevRefNo = Trim(txtReferenceNo.Text)
    End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        blnCheckforSave = False
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If mvalid = False Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdButtons.Focus()
                Else
                    If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdButtons.Focus()
                End If
            End If
        ElseIf (KeyCode = 39) Or (KeyCode = 34) Or (KeyCode = 96) Then
            KeyCode = 0
        End If
        blnCheckforSave = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0051_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0051_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If KeyCode = System.Windows.Forms.Keys.Escape Then
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlHeader_Click(ctlHeader, New System.EventArgs())
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0051_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        prgItemDetails.Visible = False
        'Load the captions
        Call FillLabelFromResFile(Me)
        'Size the form to client workspace
        Call FitToClient(Me, fraContainer, ctlHeader, cmdButtons, 500)
        'Added by Arshad to position progress bar
        prgItemDetails.Left = fraContainer.Left
        'Disabling the controls
        Call EnableControls(False, Me, True)
        'Initialising the buttons
        cmdButtons.Revert()
        'Disabling Edit, Delete and Print buttons
        cmdButtons.Enabled(1) = False
        cmdButtons.Enabled(2) = False
        cmdButtons.Enabled(5) = False
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        txtCustomerCode.BackColor = System.Drawing.Color.White
        txtConsCode.Enabled = True
        txtConsCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdHelp(6).Enabled = True

        SSPOentry.Enabled = False
        m_strSql = "Select * from company_mst where unit_code='" & gstrUNITID & "'"
        rsGetDate.GetResult(m_strSql)
        DTDate.Format = DateTimePickerFormat.Custom
        DTDate.CustomFormat = gstrDateFormat
        DTEffectiveDate.Format = DateTimePickerFormat.Custom
        DTEffectiveDate.CustomFormat = gstrDateFormat
        DTValidDate.Format = DateTimePickerFormat.Custom
        DTValidDate.CustomFormat = gstrDateFormat
        DTAmendmentDate.Format = DateTimePickerFormat.Custom
        DTAmendmentDate.CustomFormat = gstrDateFormat
        With SSPOentry
            .Col = 7
            .Col2 = 7
            .ColHidden = True
            .Col = 8
            .Col2 = 8
            .ColHidden = True
            .Col = 9
            .Col2 = 9
            .ColHidden = True
            .Col = 10
            .Col2 = 10
            .ColHidden = True
            .Col = 11
            .Col2 = 11
            .ColHidden = True
            .Col = 12
            .Col2 = 12
            .ColHidden = True
            .Col = 13
            .Col2 = 13
            .ColHidden = True
            .Col = 14
            .Col2 = 14
            .ColHidden = True
            .Col = 15
            .Col2 = 15
            .ColHidden = True
            .Col = 16
            .Col2 = 16
            .ColHidden = True
            .Col = 17
            .Col2 = 17
            .ColHidden = True
            .Col = 18
            .Col2 = 18
            .ColHidden = False
            .Col = 19
            .Col2 = 19
            .ColHidden = True
            .Col = 20
            .Col2 = 20
            .ColHidden = True
            .Col = 21
            .Col2 = 21
            .ColHidden = True
        End With
        Call InitializeSpreed()
        DTDate.Value = GetServerDate()
        DTEffectiveDate.Value = GetServerDate()
        DTAmendmentDate.Value = GetServerDate()
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        rsGetDate.ResultSetClose()
        Call SSMaxLength()
        m_blnHelpFlag = False
        m_blnCloseFlag = False
        Call AddPOType()
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        m_strSalesTaxType = ""
        frmSearch.Enabled = True
        optPartNo.Enabled = True
        optPartNo.Checked = True
        optPartNo.Checked = True
        optItem.Enabled = True
        txtsearch.Enabled = True
        txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0051_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save the data before closing
                        Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    Else
                        gblnCancelUnload = False : gblnFormAddEdit = False
                    End If
                Else
                    'Set the global variable
                    gblnCancelUnload = True : gblnFormAddEdit = True
                End If
            End If
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = 1
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0051_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Releasing the form reference
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ValidRowData(ByVal Row As Integer, Optional ByRef Col As Integer = 0) As Boolean
        '---------------------------------------------------------------
        'Arguments      :Nil
        'Return Value   :Nil
        'Function       : To validate the row data
        '---------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsParameter As ClsResultSetDB
        rsParameter = New ClsResultSetDB
        Dim rsitem As New ClsResultSetDB
        Dim varQty, varDrawPartNo, varItemCode, varOpenSO As Object
        Dim varRate As Object
        Dim varflg, varDespatch As Object
        Dim varAbatment, varMRP, varAccessibleRateforMRP As Object
        Dim dummyVarItem As Object
        Dim varDelFlag As Object
        Dim intRow As Short ' to get the values in the grid
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        Dim emptyvar As Object
        Dim strtest As String
        'change account Plug in
        varDelFlag = Nothing
        Call SSPOentry.GetText(0, Row, varDelFlag)
        SSPOentry.Row = Row
        SSPOentry.Col = 1
        varOpenSO = SSPOentry.Value
        varOpenSO = Nothing
        Call SSPOentry.GetText(1, Row, varOpenSO)
        varDrawPartNo = Nothing
        Call SSPOentry.GetText(2, Row, varDrawPartNo)
        varItemCode = Nothing
        Call SSPOentry.GetText(4, Row, varItemCode)
        varQty = Nothing
        Call SSPOentry.GetText(5, Row, varQty)
        varRate = Nothing
        Call SSPOentry.GetText(6, Row, varRate)
        varMRP = Nothing
        Call SSPOentry.GetText(19, Row, varMRP)
        varAbatment = Nothing
        Call SSPOentry.GetText(20, Row, varAbatment)
        varAccessibleRateforMRP = Nothing
        Call SSPOentry.GetText(21, Row, varAccessibleRateforMRP)
        If varDelFlag = "*" Then
            ValidRowData = True
            Exit Function
        End If
        If Col = 0 Or Col = 2 Then ' if col is 2 or entire row
            If Len(Trim(varDrawPartNo)) = 0 Then
                ValidRowData = False
                Call SSPOentry.SetText(4, Row, "")
                Call SSPOentry.SetText(2, Row, "")
                Call ssSetFocus(Row)
                SSPOentry.Focus()
                Exit Function
            End If
            m_strSql = "Select A.* from custitem_mst as A, Item_MST as B where a.Item_code=b.item_code and Status='A' and Hold_Flag=0  and a.Active=1 and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & Trim(txtCustomerCode.Text) & "' and A.Unit_Code = B.Unit_Code and A.Unit_Code = '" & gstrUNITID & "'"
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows <= 0 Then
                MsgBox("Please check reference of this  item in  " & vbCrLf & "Item Master or Customer Item Master or in Item Rate Master")
                'Call ConfirmWindow(10154, BUTTON_OK, IMG_INFO)
                ValidRowData = False
                Call SSPOentry.SetText(2, Row, "")
                Call SSPOentry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                SSPOentry.Focus()
                Exit Function
            End If
            m_strSql = "Select * from cust_ord_dtl where Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & txtCustomerCode.Text & "' and cust_ref='" & txtReferenceNo.Text & "'and active_Flag='A' and ITem_code = '" & varItemCode & "' and amendment_no = '" & txtAmendmentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'"
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows > 1 Then
                Call ConfirmWindow(10069, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                ValidRowData = False
                Call SSPOentry.SetText(2, Row, "")
                Call SSPOentry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                SSPOentry.Focus()
                Exit Function
            End If
            For intRow = 1 To SSPOentry.MaxRows
                'change account Plug in
                dummyVarItem = Nothing
                Call SSPOentry.GetText(2, intRow, dummyVarItem)
                If dummyVarItem = varDrawPartNo And intRow <> Row Then
                    Call ConfirmWindow(10156, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    ValidRowData = False
                    Call ssSetFocus(Row)
                    SSPOentry.Focus()
                    emptyvar = ""
                    If intRow > Row Then
                        Call SSPOentry.SetText(2, intRow, emptyvar)
                        Call SSPOentry.SetText(4, intRow, emptyvar)
                    Else
                        Call SSPOentry.SetText(2, Row, emptyvar)
                        Call SSPOentry.SetText(4, Row, emptyvar)
                    End If
                    Call ssSetFocus(Row, 2)
                    Exit Function
                End If
            Next
        End If
        If Col = 0 Or Col = 4 Then ' if col is 3rd
            If Len(Trim(varItemCode)) = 0 Then
                ValidRowData = False
                Call SSPOentry.SetText(2, Row, "")
                Call ssSetFocus(Row)
                SSPOentry.Focus()
                Exit Function
            End If
            m_strSql = "Select A.* from Custitem_mst as A, Item_MST As B where A.Item_code=b.Item_code and Status='A' and A.Active=1 and Hold_Flag=0 and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and Account_Code='" & Trim(txtCustomerCode.Text) & "' and A.Item_code ='" & Trim(varItemCode) & "' and A.Unit_Code = '" & gstrUNITID & "' and A.Unit_Code = B.Unit_Code"
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows <= 0 Then
                MsgBox("Please check refrence of this  item in  " & vbCrLf & "Item Master or Customer Item Master or in Item Rate Master")
                ValidRowData = False
                Call ssSetFocus(Row)
                SSPOentry.Focus()
                Exit Function
            End If
            For intRow = 1 To SSPOentry.MaxRows
                dummyVarItem = Nothing
                Call SSPOentry.GetText(4, intRow, dummyVarItem)
            Next
        End If
        If (Col = 0 Or Col = 5) And Len(varItemCode) > 0 Then ' if col is 4
            If chkOpenSo.CheckState = 0 Then
                If varOpenSO = "" Then
                    If varQty <= 0 Or Val(Trim(varQty)) <= 0 Then
                        Call ConfirmWindow(10224, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        ValidRowData = False
                        SSPOentry.Col = 5
                        Call ssSetFocus(Row, 5)
                        SSPOentry.Focus()
                        Exit Function
                    End If
                    If Col = 5 Then
                        varDespatch = Nothing
                        Call SSPOentry.GetText(17, Row, varDespatch)
                        If Val(varDespatch) > 0 Then
                            If MsgBox("Dispatch Qty for this item is [ " & varDespatch & " ] would you like to add this Quantity You have entered.", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                                Call SSPOentry.SetText(5, Row, varQty + varDespatch)
                                varQty = Nothing
                                Call SSPOentry.GetText(5, Row, varQty)
                            End If
                        End If
                    End If
                    'CHECK FOR MEASURMENT UNIT
                    strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
                    strMeasure = strMeasure & " where b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & varItemCode & "' and a.unit_Code = b.Unit_code and a.Unit_Code = '" & gstrUNITID & "'"
                    rsMeasure = New ClsResultSetDB
                    rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
                        If System.Math.Round(varQty, 3) - Val(varQty) <> 0 Then
                            Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            ValidRowData = False
                            Call SSPOentry.SetText(5, Row, CShort(varQty))
                            SSPOentry.Col = 5
                            Call ssSetFocus(Row, 5)
                            SSPOentry.Focus()
                            Exit Function
                        End If
                    End If
                    If varQty > 9999999 Then
                        'change account Plug in
                        MsgBox("Enter value less than 9999999 OR Make it Open Item.", MsgBoxStyle.OkOnly, ResolveResString(100))
                        ValidRowData = False
                        SSPOentry.Col = 5
                        SSPOentry.Row = Row : SSPOentry.Text = CStr(0)
                        Call ssSetFocus(Row, 5)
                        SSPOentry.Focus()
                        Exit Function
                    End If
                ElseIf varOpenSO = 1 Then
                    If varQty > 0 Then
                        MsgBox("This Item (" & varDrawPartNo & ") is Open , Quantity should not be greater than 0.", MsgBoxStyle.OkOnly, "eMPro")
                        ValidRowData = False
                        Call ssSetFocus(Row, 5)
                        SSPOentry.Focus()
                        Exit Function
                    End If
                ElseIf varOpenSO = "" Then
                    If varQty <= 0 Or Val(Trim(varQty)) <= 0 Then
                        If varDrawPartNo <> "" Then
                            Call ConfirmWindow(10224, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            ValidRowData = False
                            SSPOentry.Col = 5
                            Call ssSetFocus(Row, 5)
                            SSPOentry.Focus()
                            Exit Function
                        End If
                    End If
                End If
            End If 'Flag Check
        End If
        If Col = 0 Or Col = 6 Then ' if col is 5
            If varRate = 0 Or Len(Trim(varRate)) = 0 Then
                '     Call ConfirmWindow(10224, BUTTON_OK, IMG_INFO)
                MsgBox("Enter Rate Greater than 0", MsgBoxStyle.OkOnly, ResolveResString(100))
                ValidRowData = False
                Call ssSetFocus(Row, 6)
                SSPOentry.Focus()
                Exit Function
            End If
            If varQty < 0 Then
                'change account Plug in
                MsgBox("Enter Rate Greater than 0", MsgBoxStyle.OkOnly, ResolveResString(100))
                ValidRowData = False
                SSPOentry.Col = 6
                Call ssSetFocus(Row, 6)
                SSPOentry.Focus()
                Exit Function
            End If
        End If
        If UCase(Trim(cmbPOType.Text)) = "MRP-SPARES" Then
            With SSPOentry
                .Col = 19
                .Col2 = 19
                .ColHidden = False
                .Col = 20
                .Col2 = 20
                .ColHidden = False
                .Col = 21
                .Col2 = 21
                .ColHidden = False
            End With
        End If
        If Col = 0 Or Col = 21 Then ' if col is 21
            If (varAccessibleRateforMRP = 0 Or Len(Trim(varAccessibleRateforMRP)) = 0) And varMRP > 0 Then
                MsgBox("Enter Accessible Rate more than 0", MsgBoxStyle.OkOnly, ResolveResString(100))
                ValidRowData = False
                Call ssSetFocus(Row, 21)
                SSPOentry.Focus()
                Exit Function
            End If
        End If
        If Col = 0 Or Col = 20 Then ' if col is 19
            If Len(Trim(varAbatment)) >= 1 Then
                With SSPOentry
                    rsParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID = 'ABNT' and Txrt_Rate_no = '" & varAbatment & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If rsParameter.GetNoRows = 0 Then
                        MsgBox("Invalid Abatment Code.", MsgBoxStyle.Information, ResolveResString(100))
                        .Row = .ActiveRow
                        .Col = 20 : .Text = ""
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        ValidRowData = False
                        Exit Function
                    End If
                End With
            ElseIf varMRP > 0 And Len(Trim(varAbatment)) = 0 Then
                MsgBox("Abatment Code Cannot be blank", MsgBoxStyle.Information, ResolveResString(100))
                ValidRowData = False
                SSPOentry.Col = 20
                Call ssSetFocus(Row, 20)
                SSPOentry.Focus()
                Exit Function
            End If
        End If
        ValidRowData = True
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '----------------------------------------------------------------------------
        'Argument       :   Row, Col
        'Return Value   :   Nil
        'Function       :   Set the focus according to row and col value
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        With SSPOentry
            If .Enabled = True Then
                .Row = Row
                .Col = Col
                .Action = 0
                .Focus()
            End If
        End With
    End Sub
    Public Sub FillLabel(ByRef pstrCode As String)
        On Error GoTo ErrHandler
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   fills the customer detail label
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim rsCust As New ClsResultSetDB
        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where Customer_code='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                txtCreditTerms.Text = IIf(UCase(rsCust.GetValue("Credit_Days")) = "UNKNOWN", "", rsCust.GetValue("Credit_Days"))
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where crtrm_TermID = '" & txtCreditTerms.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
            Case "CURRENCY"
                m_strSql = "select Cust_Name,currency_code from Customer_mst where Customer_code='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                txtCurrencyType.Text = IIf(UCase(rsCust.GetValue("currency_code")) = "UNKNOWN", "", rsCust.GetValue("currency_code"))
            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmbPOType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPOType.Leave
        On Error GoTo ErrHandler
        If m_blnChangeFormFlg = True Then
            frmMKTTRN0010.ShowDialog()
        End If
        If cmbPOType.Text = "" Or Len(Trim(cmbPOType.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        If Not (UCase(Trim(cmbPOType.Text)) = "OEM" Or UCase(Trim(cmbPOType.Text)) = "JOB WORK" Or UCase(Trim(cmbPOType.Text)) = "SPARES" Or UCase(Trim(cmbPOType.Text)) = "MRP-SPARES" Or UCase(Trim(cmbPOType.Text)) = "EXPORT") Then
            MsgBox("Please Enter valid P.O. Type (OEM,J,S,E,M)", MsgBoxStyle.Information, ResolveResString(100))
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        If cmbPOType.SelectedIndex = 4 Then
            cmdchangetype.Visible = False
        Else
            cmdchangetype.Visible = True
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub GetAmendmentDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        m_blnGetAmendmentDetails = True
        Dim rsAD As New ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and amendment_No='" & txtAmendmentNo.Text & "' and Unit_Code = '" & gstrUNITID & "' order by Cust_Drgno"
        rsdb.GetResult(m_strSql)
        m_strSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'"
        rsAD.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                DTDate.Value = VB6.Format(CDate(rsAD.GetValue("Order_Date")), gstrDateFormat)
            End If
            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            lblRevisionNo.Text = rsAD.GetValue("RevisionNo")
            DTAmendmentDate.Value = VB6.Format(CDate(rsAD.GetValue("Amendment_Date")), gstrDateFormat)
            DTEffectiveDate.Value = VB6.Format(CDate(rsAD.GetValue("Effect_Date")), gstrDateFormat)
            DTValidDate.Value = VB6.Format(CDate(rsAD.GetValue("Valid_Date")), gstrDateFormat)
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            With SSPOentry
                .Col = 19
                .Col2 = 19
                .ColHidden = True
                .Col = 20
                .Col2 = 20
                .ColHidden = True
                .Col = 21
                .Col2 = 21
                .ColHidden = True
            End With
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "M"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
                    With SSPOentry
                        .Col = 19
                        .Col2 = 19
                        .ColHidden = False
                        .Col = 20
                        .Col2 = 20
                        .ColHidden = False
                        .Col = 21
                        .Col2 = 21
                        .ColHidden = False
                    End With
            End Select
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtCreditTerms.Text = rsAD.GetValue("Term_Payment")
            rsdb.MoveFirst()
            SSPOentry.MaxRows = 0
            intMaxLoop = rsdb.RowCount : rsdb.MoveFirst()
            prgItemDetails.Minimum = 0 : prgItemDetails.Value = 0 : prgItemDetails.Maximum = intMaxLoop
            prgItemDetails.Visible = True
            SSPOentry.Visible = False
            SSPOentry.MaxRows = intMaxLoop
            Call SSMaxLength()
            For intLoopCounter = 1 To intMaxLoop
                If rsdb.GetValue("OpenSO") = False Then
                    SSPOentry.Row = intLoopCounter
                    SSPOentry.Col = 1
                    SSPOentry.Value = 0
                Else
                    SSPOentry.Row = intLoopCounter
                    SSPOentry.Col = 1
                    SSPOentry.Value = 1
                End If
                Call SSPOentry.SetText(2, intLoopCounter, rsdb.GetValue("Cust_DrgNo"))
                Call SSPOentry.SetText(4, intLoopCounter, rsdb.GetValue("Item_Code "))
                Call SSPOentry.SetText(5, intLoopCounter, rsdb.GetValue("Order_Qty"))
                Call SSPOentry.SetText(13, intLoopCounter, rsdb.GetValue("Rate"))
                Call SSPOentry.SetText(6, intLoopCounter, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call SSPOentry.SetText(19, intLoopCounter, rsdb.GetValue("MRP"))
                Call SSPOentry.SetText(20, intLoopCounter, rsdb.GetValue("Abantment_code"))
                Call SSPOentry.SetText(21, intLoopCounter, rsdb.GetValue("AccessibleRateforMRP"))
                rsdb.MoveNext()
                prgItemDetails.Value = prgItemDetails.Value + 1
            Next
            prgItemDetails.Visible = False
            SSPOentry.Visible = True
        Else
            Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            SSPOentry.MaxRows = 0
            Exit Sub
        End If
        With SSPOentry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .BlockMode = False
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        rsAD.ResultSetClose()
        cmdButtons.Enabled(5) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub ADDRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Adds a new row in the grid
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim inti As Short
        If SSPOentry.MaxRows > 0 Then
            If ValidRowData(SSPOentry.MaxRows, 0) = True Then
                SSPOentry.MaxRows = SSPOentry.MaxRows + 1
            Else
                Exit Sub
            End If
        Else
            SSPOentry.MaxRows = SSPOentry.MaxRows + 1
        End If
        With SSPOentry
            'change account Plug in
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .BlockMode = False
        End With
        For inti = 5 To 8
            Call SSPOentry.SetText(inti, SSPOentry.MaxRows, 0)
        Next
        For inti = 11 To 12
            Call SSPOentry.SetText(inti, SSPOentry.MaxRows, 0)
        Next
        Call SSPOentry.SetText(19, SSPOentry.MaxRows, 0)
        Call SSMaxLength()
        With SSPOentry
            .Col = 1
            If .MaxRows > 1 Then
                .Row = .MaxRows
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
            End If
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function GetFormDetails() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the additional details
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rstAD As New ClsResultSetDB
        Dim strSalesTaxType As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            m_strSql = "select a.*,b.* from cust_ord_hdr a,cust_ord_dtl b where a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.Unit_Code = b.Unit_code and a.Unit_Code = '" & gstrUNITID & "'"
            rstAD.GetResult(m_strSql)
            If rstAD.GetNoRows > 0 Then
                strSalesTaxType = rstAD.GetValue("SalesTax_Type")
            End If
            If IsDBNull(strSalesTaxType) Then
                GetFormDetails = False
            Else
                GetFormDetails = True
            End If
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            GetFormDetails = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetReferenceDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the additional details
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim intDecimal As Short
        Dim strMax As String
        Dim strMin As String
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and amendment_No = '" & txtAmendmentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'" ' and active_Flag IN('A','L') order by Cust_drgNo"
        rsAD.GetResult(m_strSql)
        m_strSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag in ('A','L') and Unit_Code = '" & gstrUNITID & "'"
        rsRefNo.GetResult(m_strSql)
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            txtAmendmentNo.Text = " "
            lblIntSONoDes.Text = rsRefNo.GetValue("InternalSONo")
            lblRevisionNo.Text = rsRefNo.GetValue("RevisionNo")
            If Len(Trim(txtAmendmentNo.Text)) > 0 Then
                DTAmendmentDate.Value = rsRefNo.GetValue("Amendment_Date")
            Else
                DTAmendmentDate.Value = GetServerDate()
            End If
            txtCreditTerms.Text = rsRefNo.GetValue("Term_Payment")
            txtSTax.Text = rsRefNo.GetValue("SALESTAX_TYPE")
            DTEffectiveDate.Value = rsRefNo.GetValue("Effect_Date")
            DTValidDate.Value = rsRefNo.GetValue("Valid_Date")
            txtCurrencyType.Text = rsRefNo.GetValue("Currency_Code")
            ctlPerValue.Text = rsRefNo.GetValue("PerValue")
            txtAmendReason.Text = rsRefNo.GetValue("Reason")
            DTDate.Value = rsRefNo.GetValue("Order_date")
            strpotype = rsRefNo.GetValue("PO_Type")
            strSOType = rsRefNo.GetValue("salestax_Type")
            SSPOentry.Col = 19
            SSPOentry.Col2 = 19
            SSPOentry.ColHidden = True
            SSPOentry.Col = 20
            SSPOentry.Col2 = 20
            SSPOentry.ColHidden = True
            SSPOentry.Col = 21
            SSPOentry.Col2 = 21
            SSPOentry.ColHidden = True
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "M"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
                    SSPOentry.Col = 19
                    SSPOentry.Col2 = 19
                    SSPOentry.ColHidden = False
                    SSPOentry.Col = 20
                    SSPOentry.Col2 = 20
                    SSPOentry.ColHidden = False
                    SSPOentry.Col = 21
                    SSPOentry.Col2 = 21
                    SSPOentry.ColHidden = False
            End Select
            If rsRefNo.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            rsAD.MoveFirst()
            SSPOentry.MaxRows = 0
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where currency_code ='" & Trim(txtCurrencyType.Text) & "' and Unit_Code = '" & gstrUNITID & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                intMaxCounter = rsAD.GetNoRows
                prgItemDetails.Value = 0 : prgItemDetails.Minimum = 0 : prgItemDetails.Maximum = intMaxCounter
                prgItemDetails.Visible = True
                rsAD.MoveFirst()
                With SSPOentry
                    For intLoopCounter = 1 To intMaxCounter
                        SSPOentry.MaxRows = SSPOentry.MaxRows + 1
                        m_custItemDesc = rsAD.GetValue("Cust_Drg_Desc")
                        If rsAD.GetValue("OpenSO") = False Then
                            .Col = 1
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            SSPOentry.Value = 0
                        Else
                            .Col = 1
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            SSPOentry.Value = 1
                        End If
                        .Col = 2
                        .Row = intLoopCounter
                        .TypeMaxEditLen = 30
                        Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsAD.GetValue("Cust_DrgNo"))
                        .Col = 4
                        .Row = intLoopCounter
                        .TypeMaxEditLen = 16
                        Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsAD.GetValue("Item_Code "))
                        .Col = 5
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = "0.00"
                        .TypeFloatMax = "9999999.99"
                        Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsAD.GetValue("Order_Qty"))
                        .Col = 13
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsAD.GetValue("Rate"))
                        .Col = 6
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call SSPOentry.SetText(6, SSPOentry.MaxRows, rsAD.GetValue("Rate") * CDbl(ctlPerValue.Text))
                        .Col = 18
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        Call SSPOentry.SetText(18, SSPOentry.MaxRows, rsAD.GetValue("Remarks"))
                        .Col = 19
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call SSPOentry.SetText(19, SSPOentry.MaxRows, rsAD.GetValue("MRP"))
                        .Col = 20
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        Call SSPOentry.SetText(20, SSPOentry.MaxRows, rsAD.GetValue("abantment_code"))
                        .Col = 21
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call SSPOentry.SetText(21, SSPOentry.MaxRows, rsAD.GetValue("AccessibleRateforMRP"))
                        rsAD.MoveNext()
                        prgItemDetails.Value = prgItemDetails.Value + 1
                    Next
                    prgItemDetails.Visible = False
                End With
            End If
        Else
            Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Exit Sub
        End If
        With SSPOentry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .BlockMode = False
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End If
        End With
        cmdButtons.Enabled(5) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub DeleteRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Deletes a row
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim rsDespatchQuantity As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrawCode As Object
        ReDim ArrDispatchQty(SSPOentry.MaxRows - 1)
        rsDespatchQuantity = New ClsResultSetDB
        For intRow = 1 To SSPOentry.MaxRows
            varItemCode = Nothing
            Call SSPOentry.GetText(4, intRow, varItemCode)
            varDrawCode = Nothing
            Call SSPOentry.GetText(2, intRow, varDrawCode)
            rsDespatchQuantity.GetResult("Select Despatch_Qty from cust_ord_dtl where Account_Code='" & txtCustomerCode.Text & "'and Cust_Ref='" & txtReferenceNo.Text & "' and Amendment_No= '" & txtAmendmentNo.Text & "' and Item_Code= '" & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0 and Unit_Code = '" & gstrUNITID & "'")
            If rsDespatchQuantity.GetNoRows > 0 Then
                ArrDispatchQty(intRow - 1) = rsDespatchQuantity.GetValue("Despatch_Qty")
            Else
                ArrDispatchQty(intRow - 1) = 0
            End If
            strsql = "delete cust_ord_dtl where Account_Code='"
            strsql = strsql & txtCustomerCode.Text & "'and Cust_Ref='"
            strsql = strsql & txtReferenceNo.Text & "' and Amendment_No= '"
            strsql = strsql & txtAmendmentNo.Text & "' and Item_Code= '"
            strsql = strsql & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0 and Unit_Code = '" & gstrUNITID & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Next
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub InsertRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Inserts a row
        'Comments       :   Nil
        'Revision History   :    Changes done by Ashutosh on 10-10-2005 , Issue Id:15876, Bug fix of SO entry form , If Item is saved as closed one but it still saved as Open.
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varDeleteFlag As Object
        Dim clsDrgDes As New ClsResultSetDB
        Dim strsql As String
        Dim varRate, varItemCode, varordqty As Object
        Dim varAbatment_code, varMRP, varAccessibleRateforMRP As Object
        Dim varDespatchQty, varCustItemCode, varCustItemDesc, varOpenSO As Object
        Dim varRemarks As Object
        Dim rsSalesParameter As New ClsResultSetDB

        For intRow = 1 To SSPOentry.MaxRows
            varDeleteFlag = Nothing
            Call SSPOentry.GetText(0, intRow, varDeleteFlag)
            If varDeleteFlag <> "*" Then 'to get the values from the grid
                varCustItemCode = Nothing
                Call SSPOentry.GetText(2, intRow, varCustItemCode)
                'Getting the Drawing No. Description
                strsql = "SELECT drg_desc FROM  custitem_mst WHERE Active=1 and Cust_drgno = '" & Trim(varCustItemCode) & "' and Unit_Code = '" & gstrUNITID & "'"
                clsDrgDes = New ClsResultSetDB
                If clsDrgDes.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsDrgDes.GetNoRows > 0 Then
                    clsDrgDes.MoveFirst()
                    m_custItemDesc = Trim(clsDrgDes.GetValue("drg_desc"))
                    clsDrgDes.ResultSetClose()
                End If
                varOpenSO = Nothing
                Call SSPOentry.GetText(1, intRow, varOpenSO)
                varItemCode = Nothing
                Call SSPOentry.GetText(4, intRow, varItemCode)
                varordqty = Nothing
                Call SSPOentry.GetText(5, intRow, varordqty)
                varRate = Nothing
                Call SSPOentry.GetText(6, intRow, varRate)
                If CDbl(ctlPerValue.Text) >= 1 Then
                    varRate = varRate / CDbl(ctlPerValue.Text)
                End If
                varRemarks = Nothing
                Call SSPOentry.GetText(18, intRow, varRemarks)
                varMRP = Nothing
                Call SSPOentry.GetText(19, intRow, varMRP)
                varAbatment_code = Nothing
                Call SSPOentry.GetText(20, intRow, varAbatment_code)
                varAccessibleRateforMRP = Nothing
                Call SSPOentry.GetText(21, intRow, varAccessibleRateforMRP)
                If varMRP = 0.0 Then varMRP = 0
                If varAccessibleRateforMRP = 0.0 Then varAccessibleRateforMRP = 0
                If cmbPOType.Text <> "MRP-SPARES" Then
                    varMRP = "0"
                    varAbatment_code = ""
                End If
                strsql = "Insert into Cust_Ord_Dtl (Account_Code, Cust_Ref, Amendment_No,InternalSONo,RevisionNo, "
                strsql = strsql & "Item_Code , Rate, Order_Qty, Despatch_Qty, "
                strsql = strsql & "Active_Flag,Cust_DrgNo,"
                strsql = strsql & "Remarks,MRP,abantment_code,AccessibleRateforMRP,"
                strsql = strsql & "Cust_Drg_Desc,"
                strsql = strsql & "Authorized_flag, openSO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId, Consignee_code,PerValue,Unit_Code )"
                strsql = strsql & " values('" & Trim(txtCustomerCode.Text) & "','"
                strsql = strsql & Trim(txtReferenceNo.Text) & "','"
                strsql = strsql & Trim(txtAmendmentNo.Text) & "','"
                strsql = strsql & Trim(lblIntSONoDes.Text) & "'," & Trim(lblRevisionNo.Text) & ",'"
                strsql = strsql & Trim(varItemCode) & "',"
                strsql = strsql & Trim(varRate) & ","
                strsql = strsql & IIf(IsNothing(varordqty), 0, varordqty) & "," & ArrDispatchQty(intRow - 1) & ",'A','"
                strsql = strsql & varCustItemCode & "','"
                strsql = strsql & Trim(varRemarks) & "',"
                strsql = strsql & Trim(varMRP) & ",'"
                strsql = strsql & Trim(varAbatment_code) & "',"
                strsql = strsql & (Trim(varAccessibleRateforMRP)) & ",'"
                strsql = strsql & Trim(m_custItemDesc) & "',0,"
                If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strsql = strsql & "1,"
                Else
                    If CStr(varOpenSO) = "" Then
                        strsql = strsql & "0,"
                    Else
                        strsql = strsql & "1,"
                    End If
                End If
                strsql = strsql & " getdate()" & ",'" & mP_User & "'," & "getdate() "
                strsql = strsql & ",'" & mP_User & "',"
                Dim strConsCode As String
                If Len(txtConsCode.Text.Trim) = 0 Then
                    strConsCode = Trim(Me.txtCustomerCode.Text)
                Else
                    strConsCode = Trim(txtConsCode.Text)
                End If

                strsql = strsql & "'" & strConsCode & "',"

                If CDbl(ctlPerValue.Text) >= 1 Then
                    strsql = strsql & ctlPerValue.Text & ",'" & gstrUNITID & "')"
                Else
                    strsql = strsql & " 1,'" & gstrUNITID & "' )"
                End If
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Next
        clsDrgDes = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub UpdateRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Updates the header table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        strsql = "update cust_ord_hdr set Order_Date='"
        strsql = strsql & getDateForDB(DTDate.Text) & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            strsql = strsql & "Amendment_Date='"
            strsql = strsql & getDateForDB(DTAmendmentDate.Text) & "',"
        End If
        strsql = strsql & "Currency_Code='"
        strsql = strsql & txtCurrencyType.Text & "',Valid_Date='"
        strsql = strsql & getDateForDB(DTValidDate.Text) & "',Effect_Date='"
        strsql = strsql & getDateForDB(DTEffectiveDate.Text) & "',Term_Payment='"
        strsql = strsql & Trim(txtCreditTerms.Text) & "',Special_Remarks='" & m_strSpecialNotes & "',Pay_Remarks='"
        strsql = strsql & m_strPaymentTerms & "',Price_Remarks='" & m_strPricesAre & "',Packing_Remarks='"
        strsql = strsql & m_strPkgAndFwd & "',Frieght_Remarks='" & m_strFreight & "',Transport_Remarks='"
        strsql = strsql & m_strTransitInsurance & "',Octorai_Remarks='" & m_strOctroi & "',Mode_Despatch='"
        strsql = strsql & m_strModeOfDespatch & "',Delivery='" & m_strDeliverySchedule & "',"
        strsql = strsql & "Reason='" & txtAmendReason.Text & "',PO_Type='"
        strsql = strsql & Mid(cmbPOType.Text, 1, 1) & "',"
        strsql = strsql & "SalesTax_Type='" & txtSTax.Text & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strsql = strsql & " OpenSO = 1,"
        Else
            strsql = strsql & " OpenSO = 0,"
        End If
        strsql = strsql & " Ent_dt="
        strsql = strsql & " getdate() " & ",Ent_UserId='" & mP_User & "',Upd_dt="
        strsql = strsql & " getdate() " & ",Upd_UserId='" & mP_User & "', Unit_code = '" & gstrUNITID & "'"
        If CDbl(ctlPerValue.Text) >= 1 Then
            strsql = strsql & ", PerValue = " & ctlPerValue.Text & " where Account_Code='"
        Else
            strsql = strsql & ", PerValue = 1 where Account_Code='"
        End If
        strsql = strsql & txtCustomerCode.Text & "'and Cust_Ref='"
        strsql = strsql & txtReferenceNo.Text & "'and Amendment_No='" & txtAmendmentNo.Text & "' and Unit_code = '" & gstrUNITID & "'"
        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub InsertRowCustOrdHdr()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Inserts Row in the header table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varExiseDuty As Object
        Dim varSalestax As Object
        Dim varSurchargeSalesTax As Object
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            lblIntSONoDes.Text = GenerateDocumentNumber("Cust_ord_hdr", "InternalSONo", "ent_dt", getDateForDB(GetServerDate()))
        End If
        strsql = "Insert into Cust_Ord_Hdr (Account_Code, Cust_Ref, Amendment_No,InternalSONo,RevisionNo, Order_Date, "
        strsql = strsql & "Amendment_Date, Active_Flag, "
        strsql = strsql & " Currency_Code, Valid_Date,"
        strsql = strsql & "Effect_Date, Term_Payment, Special_Remarks, Pay_Remarks, "
        strsql = strsql & "Price_Remarks, Packing_Remarks, Frieght_Remarks, Transport_Remarks,"
        strsql = strsql & "Octorai_Remarks, Mode_Despatch, Delivery, First_Authorized,"
        strsql = strsql & "Second_Authorized, Third_Authorized, Authorized_Flag, Reason, "
        strsql = strsql & "PO_Type, SalesTax_Type,OpenSO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId, Consignee_Code,PerValue,Unit_Code)"
        strsql = strsql & " Values('" & Trim(txtCustomerCode.Text) & "','" & Trim(txtReferenceNo.Text) & "','" & Trim(txtAmendmentNo.Text) & "',"
        strsql = strsql & "'" & lblIntSONoDes.Text & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            lblRevisionNo.Text = CStr(GenerateRevisionNo())
            strsql = strsql & lblRevisionNo.Text & ",'"
        Else
            lblRevisionNo.Text = "0"
            strsql = strsql & "0,'"
        End If
        strsql = strsql & getDateForDB(DTDate.Text) & "','" & IIf(Len(Me.txtAmendmentNo.Text) = 0, System.DBNull.Value, getDateForDB(DTAmendmentDate.Text)) & "','A' ,'"
        strsql = strsql & Trim(txtCurrencyType.Text) & "','" & getDateForDB(DTValidDate.Text) & "','" & getDateForDB(DTEffectiveDate.Text) & "','"
        strsql = strsql & IIf(Len(Trim(txtCreditTerms.Text)) = 0, 0, Trim(txtCreditTerms.Text)) & "','"
        strsql = strsql & Trim(m_strSpecialNotes) & "','"
        strsql = strsql & Trim(m_strPaymentTerms) & "','" & Trim(m_strPricesAre) & "','" & Trim(m_strPkgAndFwd) & "','" & Trim(m_strFreight) & "','"
        strsql = strsql & Trim(m_strTransitInsurance) & "','" & Trim(m_strOctroi) & "','" & Trim(m_strModeOfDespatch) & "','" & Trim(m_strDeliverySchedule) & "','',"
        strsql = strsql & "'','','','" & Trim(txtAmendReason.Text) & "','" & Mid(cmbPOType.Text, 1, 1) & "','" & txtSTax.Text & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strsql = strsql & "1,"
        Else
            strsql = strsql & "0,"
        End If
        strsql = strsql & " getdate() " & ",'" & mP_User & "'," & " getdate() " & ",'" & mP_User & "'"

        Dim strConsCode As String
        If Len(txtConsCode.Text.Trim) = 0 Then
            strConsCode = Me.txtCustomerCode.Text.Trim
        Else
            strConsCode = txtConsCode.Text.Trim
        End If
        strsql = strsql & ",'" & strConsCode & "'"

        If Val(ctlPerValue.Text) >= 1 Then
            strsql = strsql & "," & ctlPerValue.Text & ",'" & gstrUNITID & "')"
        Else
            strsql = strsql & ", 1,'" & gstrUNITID & "')"
        End If
        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function CheckFormDetails() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Check The Additional Details that they have been entered or not
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rstAD As New ClsResultSetDB
        If Len(Trim(m_strPaymentTerms)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                m_strSql = "select a.*,b.* from cust_ord_hdr a,cust_ord_dtl b where a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.consignee_code = b.consignee_code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.unit_Code = b.Unit_code and a.unit_Code = '" & gstrUNITID & "'"
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                m_strSql = "select * from cust_ord_hdr  where Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "' and consignee_code='" & txtConsCode.Text.Trim & "' and unit_Code = '" & gstrUNITID & "'"
            End If
            rstAD.GetResult(m_strSql)
            If rstAD.GetNoRows > 0 Then
                m_strSpecialNotes = rstAD.GetValue("Special_Remarks")
                m_strPaymentTerms = rstAD.GetValue("Pay_Remarks")
                m_strPricesAre = rstAD.GetValue("Price_Remarks")
                m_strPkgAndFwd = rstAD.GetValue("Packing_Remarks")
                m_strFreight = rstAD.GetValue("Frieght_Remarks")
                m_strTransitInsurance = rstAD.GetValue("Transport_Remarks")
                m_strOctroi = rstAD.GetValue("Octorai_Remarks")
                m_strModeOfDespatch = rstAD.GetValue("Mode_Despatch")
                m_strDeliverySchedule = rstAD.GetValue("Delivery")
                CheckFormDetails = True
            Else
                CheckFormDetails = False
                Exit Function
            End If
        Else
            CheckFormDetails = True
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetailsForAuthorizedPO()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        m_blnGetAmendmentDetails = True
        Dim rsAD As New ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and consignee_code='" & txtConsCode.Text.Trim & "' and  Authorized_Flag=1 and unit_Code = '" & gstrUNITID & "' order by Cust_drgNo"
        rsdb.GetResult(m_strSql)
        m_strSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and consignee_code='" & txtConsCode.Text.Trim & "' and authorized_Flag=1 and unit_Code = '" & gstrUNITID & "'"
        rsAD.GetResult(m_strSql)
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            lblRevisionNo.Text = rsAD.GetValue("RevisionNo")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "M"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
            End Select
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtCreditTerms.Text = rsAD.GetValue("Term_Payment")
            rsdb.MoveFirst()
            SSPOentry.MaxRows = 0
            Do While Not rsdb.EOFRecord
                SSPOentry.MaxRows = SSPOentry.MaxRows + 1
                If rsdb.GetValue("OpenSO") = False Then
                    SSPOentry.Row = SSPOentry.MaxRows
                    SSPOentry.Col = 1
                    SSPOentry.Value = 0
                Else
                    SSPOentry.Row = SSPOentry.MaxRows
                    SSPOentry.Col = 1
                    SSPOentry.Value = 1
                End If
                Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsdb.GetValue("Cust_DrgNo"))
                Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsdb.GetValue("Item_Code "))
                Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsdb.GetValue("Order_Qty"))
                Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsdb.GetValue("Rate"))
                Call SSPOentry.SetText(6, SSPOentry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                rsdb.MoveNext()
            Loop
        Else
            Call ConfirmWindow(10130, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            SSPOentry.MaxRows = 0
            Exit Sub
        End If
        With SSPOentry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .BlockMode = False
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        rsAD.ResultSetClose()
        cmdButtons.Enabled(5) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub TxtReferenceNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtReferenceNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Sub UpdateActiveFlag()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   fills the customer detail label
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As New ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim AmendmentNo As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intmaxitems As Short
        Dim rsCustOrdHdr As ClsResultSetDB
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and consignee_code='" & txtConsCode.Text.Trim & "' and active_Flag='A' and unit_Code = '" & gstrUNITID & "'"
        rsAD.GetResult(m_strSql)
        intmaxitems = rsAD.GetNoRows
        intMaxLoop = SSPOentry.MaxRows
        ReDim ArrDispatchQty(intMaxLoop - 1)
        For intLoopCounter = 1 To intMaxLoop
            varItemCode = Nothing
            Call SSPOentry.GetText(4, intLoopCounter, varItemCode)
            varDrgNo = Nothing
            Call SSPOentry.GetText(2, intLoopCounter, varDrgNo)
            m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and consignee_code='" & txtConsCode.Text.Trim & "' and active_Flag='A' and Item_Code ='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "' and unit_Code = '" & gstrUNITID & "'"
            rsAD.GetResult(m_strSql)
            If rsAD.GetNoRows >= 1 Then
                ArrDispatchQty(intLoopCounter - 1) = rsAD.GetValue("Despatch_qty")
                m_strSql = "update cust_ord_dtl set Active_Flag='O' where Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and Item_Code='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "' and unit_Code = '" & gstrUNITID & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Else
                ArrDispatchQty(intLoopCounter - 1) = 0
            End If
        Next
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function ConfirmDeletion() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   confirms the deletion if some rows are marked for deletion
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim blndelrow As Boolean
        Dim strAns As MsgBoxResult
        Dim vardelrow As Object
        ConfirmDeletion = True
        For intRow = 1 To SSPOentry.MaxRows
            vardelrow = Nothing
            Call SSPOentry.GetText(0, intRow, vardelrow)
            If vardelrow = "*" Then
                blndelrow = True
                Exit Function
            Else
                blndelrow = False
            End If
        Next
        If blndelrow = True Then
            Call ConfirmWindow(10101, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            strAns = ConfirmWindow(10099, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
            If strAns = MsgBoxResult.Yes Then
                Exit Function
            ElseIf strAns = MsgBoxResult.No Then
                For intRow = 1 To SSPOentry.MaxRows
                    Call SSPOentry.SetText(0, intRow, "")
                Next
                Exit Function
            Else
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                ConfirmDeletion = False
                Exit Function
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub SSMaxLength()
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim strMin As String
        Dim strMax As String
        Dim intDecimal As Short
        Dim intLoopCounter As Short
        Dim rscurrency As ClsResultSetDB
        With Me.SSPOentry
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = 2
                .TypeMaxEditLen = 30
                .Col = 3
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
                .Col = 4
                .TypeMaxEditLen = 16
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 2
                .TypeFloatMin = "0.00"
                .TypeFloatMax = "9999999.99"
                If chkOpenSo.CheckState = 1 Then
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 5
                    .Col2 = 5
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                Else
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 5
                    .Col2 = 5
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End If

                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where currency_code ='" & Trim(txtCurrencyType.Text) & "' and unit_Code = '" & gstrUNITID & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                .Row = intRow
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 6
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 19
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 21
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
            Next intRow
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub RefreshForm()
        Dim rsGetDate As ClsResultSetDB
        On Error GoTo Err_Handler
        txtCurrencyType.Text = ""
        txtAmendReason.Text = ""
        lblIntSONoDes.Text = ""
        lblRevisionNo.Text = ""
        lblCustPartDesc.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        DTDate.Value = GetServerDate()
        DTAmendmentDate.Value = GetServerDate()
        DTEffectiveDate.Value = GetServerDate()
        rsGetDate = New ClsResultSetDB
        rsGetDate.GetResult("select * from Company_Mst where Unit_Code = '" & gstrUNITID & "'")
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        SSPOentry.MaxRows = 0
        m_strSalesTaxType = ""
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub AddPOType()
        On Error GoTo Err_Handler
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "MRP-SPARES")
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidRecord() As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        blnInvalidData = False
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Customer Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCustomerCode
        End If
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Reference No"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtReferenceNo
        End If
        If txtAmendmentNo.Enabled Then
            If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Amendment No "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtAmendmentNo
            End If
        End If
        If Len(Trim(txtCurrencyType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Currency Type"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCurrencyType
        End If
        If Len(Trim(cmbPOType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "SO Type"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = cmbPOType
        End If
        If Len(Trim(txtCreditTerms.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Credit Terms "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCreditTerms
        End If
        If cmbPOType.SelectedIndex <> 4 Then
            If CheckFormDetails() = False Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Sales Terms "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = cmdchangetype
            End If
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & " ."
        lNo = lNo + 1
        blnValidAmendDate = True
        If Len(Trim(txtAmendmentNo.Text)) >= 1 Then
            If DTAmendmentDate.Value > DTValidDate.Value Then
                Call ConfirmWindow(10146, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
            If DTAmendmentDate.Value < DTDate.Value Then
                Call ConfirmWindow(10147, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
            If DTAmendmentDate.Value > GetServerDate() Then
                MsgBox("Amendment Date Can not be Greater Than Current Date", MsgBoxStyle.Information, ResolveResString(100))
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
        End If
        If blnInvalidData = True And blnValidAmendDate = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = True
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub CLEARVAR()
        m_strSalesTaxType = ""
        m_intCreditDays = ""
        m_strSpecialNotes = ""
        m_strPaymentTerms = ""
        m_strPricesAre = ""
        m_strPkgAndFwd = ""
        m_strFreight = ""
        m_strTransitInsurance = ""
        m_strOctroi = ""
        m_strModeOfDespatch = ""
        m_strDeliverySchedule = ""
    End Sub
    Public Sub PrintToReport()
        '*********************************************'
        'Author:                Ananya Nath
        'Arguments:             None
        'Return Value   :       None
        'Description    :       Used to print currently selected/entered sales Order.
        '*********************************************'
        Dim strReportName As String
        On Error GoTo ErrHandler
        '<<<<CR11 Code Starts>>>>
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlHeader.HeaderString
        '<<<<CR11 Code Ends>>>>
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.AppStarting)
        strReportName = GetPlantName()
        strReportName = "\Reports\rptSOPrinting_" & strReportName & ".rpt"
        If Not CheckFile(strReportName) Then
            strReportName = "\Reports\rptSOPrinting.rpt"
        End If
        With objRpt
            'load the report
            .Load(My.Application.Info.DirectoryPath & strReportName)
            .DataDefinition.FormulaFields("Comp_name").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("Comp_address").Text = "'" & gstr_WRK_ADDRESS1 & "'"
        End With
        strsql = ""
        If Len(Trim(Me.txtAmendmentNo.Text)) = 0 Then
            strsql = " {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text & "' and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text & "' and  {cust_ord_hdr.amendment_no} = ''" 'Initialising Sql Query.
        Else
            strsql = " {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text & "'  and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text & "' and  {cust_ord_hdr.amendment_no} = '" & txtAmendmentNo.Text & "'" 'Initialising Sql Query.
        End If
        objRpt.RecordSelectionFormula = strsql & " and {cust_ord_hdr.UNIT_CODE} = '" & gstrUNITID & "'"
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        frmReportViewer.Zoom = 120
        frmReportViewer.Show()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub
    Public Function checkforitemRate(ByRef pintRow As Short) As Boolean
        Dim varDrgNo As Object
        Dim strItemRate As String
        Dim rsItemRate As ClsResultSetDB
        With SSPOentry
            varDrgNo = Nothing
            Call .GetText(2, pintRow, varDrgNo)
            strItemRate = "select Edit_flg from ITemRate_Mst where Unit_Code = '" & gstrUNITID & "' and Serial_No = (select max(serial_no) from "
            strItemRate = strItemRate & " itemrate_mst where Party_code = '" & txtCustomerCode.Text
            strItemRate = strItemRate & "' and item_code = '" & varDrgNo
            strItemRate = strItemRate & "' and datediff(mm,convert(varchar(10),'" & DTDate.Value & "',103),convert(varchar(10),DateFrom,103))<=0 "
            strItemRate = strItemRate & " and custVend_Flg ='C' and Unit_Code = '" & gstrUNITID & "')"
            rsItemRate = New ClsResultSetDB
            rsItemRate.GetResult(strItemRate)
            checkforitemRate = rsItemRate.GetValue("Edit_Flg")
        End With
    End Function
    Public Sub SetCellTypeCombo(ByRef intRow As Short)
        Dim strcustdtl As String
        Dim StrItemCode As Object
        Dim strDrgNo As Object
        Dim FinalstrItemCode As String
        Dim rsitem As ClsResultSetDB
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call SSPOentry.GetText(2, intRow, strDrgNo)
        strcustdtl = "SElect * from custITem_Mst where Active=1 and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
        rsitem.GetResult(strcustdtl)
        With SSPOentry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
            rsitem.MoveFirst()
            FinalstrItemCode = ""
            While Not rsitem.EOFRecord
                StrItemCode = IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code"))
                FinalstrItemCode = FinalstrItemCode & StrItemCode & Chr(9) '& "[V]alue":
                rsitem.MoveNext()
            End While
            FinalstrItemCode = VB.Left(FinalstrItemCode, Len(FinalstrItemCode) - 1)
            .TypeComboBoxList = FinalstrItemCode
        End With
    End Sub
    Public Sub SetCellStatic(ByRef intRow As Integer)
        With SSPOentry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
        End With
    End Sub
    Public Function chkmultipleitem() As Boolean
        Dim StrItemCode As Object
        Dim strDrgNo As Object
        Dim rsitem As ClsResultSetDB
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call SSPOentry.GetText(2, SSPOentry.ActiveRow, strDrgNo)
        StrItemCode = "Select * from custITem_Mst where  Active=1 AND cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
        rsitem.GetResult(StrItemCode)
        If rsitem.GetNoRows > 1 Then
            chkmultipleitem = True
        Else
            chkmultipleitem = False
        End If
    End Function
    Public Sub RowDetailsfromKeyBoard(ByRef pstrItemCode As Object, ByRef pstrDrgno As Object)
        Dim rsitem As ClsResultSetDB
        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Len(Trim(pstrItemCode)) > 0 Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    m_strSql = "Select * from Cust_ord_dtl where Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' and cust_ref ='" & txtReferenceNo.Text & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A' and Unit_Code = '" & gstrUNITID & "'"
                Else
                    m_strSql = "Select * from Cust_ord_dtl where Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1 and Unit_Code = '" & gstrUNITID & "'"
                End If
                rsitem = New ClsResultSetDB
                rsitem.GetResult(m_strSql)
                If rsitem.GetNoRows > 0 Then
                    Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                    Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsitem.GetValue("Item_Code "))
                    lblCustPartDesc.Text = rsitem.GetValue("Cust_drg_desc")
                    Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsitem.GetValue("Order_Qty"))
                    Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                    Call SSPOentry.SetText(6, SSPOentry.MaxRows, rsitem.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Else
                    If txtAmendmentNo.Enabled = False Then
                        If Len(Trim(pstrItemCode)) > 0 Then
                            m_strSql = " select * from ITemRate_Mst where Unit_Code = '" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & pstrDrgno & "' and datediff(mm,convert(varchar(10),'" & DTDate.Value & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C' and Unit_Code = '" & gstrUNITID & "')"
                            rsitem = New ClsResultSetDB
                            rsitem.GetResult(m_strSql)
                            If rsitem.GetNoRows > 0 Then
                                If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Col = 1
                                    SSPOentry.Col2 = 1
                                    SSPOentry.Value = System.Windows.Forms.CheckState.Unchecked
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Row2 = SSPOentry.MaxRows
                                    SSPOentry.Col = 5
                                    SSPOentry.Col2 = 5
                                    SSPOentry.BlockMode = True
                                    SSPOentry.Lock = False
                                    SSPOentry.BlockMode = False
                                Else
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Col = 1
                                    SSPOentry.Col2 = 1
                                    SSPOentry.Value = System.Windows.Forms.CheckState.Checked
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Row2 = SSPOentry.MaxRows
                                    SSPOentry.Col = 5
                                    SSPOentry.Col2 = 5
                                    SSPOentry.BlockMode = True
                                    SSPOentry.Lock = True
                                    SSPOentry.BlockMode = False
                                End If
                                Call SSPOentry.SetText(2, SSPOentry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                Call SSPOentry.SetText(4, SSPOentry.MaxRows, pstrItemCode)
                                Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                                Call SSPOentry.SetText(6, SSPOentry.MaxRows, rsitem.GetValue("Rate") * CDbl(ctlPerValue.Text))
                                If rsitem.GetValue("Edit_flg") = False Then
                                    SSPOentry.Col = 6
                                    SSPOentry.Col2 = SSPOentry.MaxCols
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Row2 = SSPOentry.MaxRows
                                    SSPOentry.BlockMode = True
                                    SSPOentry.Lock = True
                                    SSPOentry.BlockMode = False
                                Else
                                    SSPOentry.Col = 6
                                    SSPOentry.Col2 = SSPOentry.MaxCols
                                    SSPOentry.Row = SSPOentry.MaxRows
                                    SSPOentry.Row2 = SSPOentry.MaxRows
                                    SSPOentry.BlockMode = True
                                    SSPOentry.Lock = False
                                    SSPOentry.BlockMode = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Call ssSetFocus(SSPOentry.MaxRows, 3)
    End Sub
    Public Function GenerateDocumentNumber(ByVal pstrTableName As String, ByVal pstrDocNofield As String, ByRef pstrDateFieldName As String, ByVal pstrWantedDate As String) As String
        On Error GoTo ErrHandler
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim NewTempSeries As String 'Generate a NEW Series
        Dim rsDocumentNoSO As ClsResultSetDB
        rsDocumentNoSO = New ClsResultSetDB
        If Len(Trim(pstrWantedDate)) > 0 Then 'For Post Dated Docs
            'No need to check for Previously made documents for After Dates
            mP_Connection.Execute("Set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            rsDocumentNoSO.GetResult("Select DocNo = Max(convert(int,substring(" & pstrDocNofield & ",9,7))) from " & pstrTableName & " Where datePart(mm,ent_dt) = datePart(mm,'" & pstrWantedDate & "') and datePart(yyyy,ent_dt) = datePart(yyyy,'" & pstrWantedDate & "') and Unit_Code = '" & gstrUNITID & "'")
            strCheckDOcNo = rsDocumentNoSO.GetValue("DocNo")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Period
            strTempSeries = CStr(CDbl(strCheckDOcNo) + 1)
            If Val(strTempSeries) < 9999 Then
                strTempSeries = New String("0", 4 - Len(strTempSeries)) & strTempSeries 'Concatenate Zeroes before the Number
            End If
            strCheckDOcNo = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & strTempSeries
            GenerateDocumentNumber = strCheckDOcNo
        Else 'The Document has not been made for that Period
            NewTempSeries = NewTempSeries & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & "0001"
            GenerateDocumentNumber = NewTempSeries 'The Number Is Generated
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
    End Function
    Public Function GenerateRevisionNo() As Short
        Dim rsRevisionNo As ClsResultSetDB
        rsRevisionNo = New ClsResultSetDB
        rsRevisionNo.GetResult("Select Revision = Max(RevisionNo) from Cust_ord_hdr where cust_ref = '" & txtReferenceNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")
        GenerateRevisionNo = IIf(IsDBNull(rsRevisionNo.GetValue("Revision")), 1, Val(rsRevisionNo.GetValue("Revision")) + 1)
    End Function
    Public Sub InsertPreviousSODetails(ByRef pstrAccountCode As String, ByRef pstrRef As String, ByRef pstrAmendment As String, ByRef pstrInternalSONo As String, ByRef pintRevisionNo As Short)
        '*********************************************'
        'Author:                Nisha Rai
        'Arguments:             Account_code , CustRef,Amendment_no , IntSONo,RevisionNo
        'Return Value   :       None
        'Description    :       To Insert active item details from base SO & its amendment which are not there in Grid.
        '*********************************************'
        Dim strsql As String
        Dim strDrgItem As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim VarDelete As Object
        Dim rsCustOrdDtl As ClsResultSetDB
        On Error GoTo ErrHandler
        strsql = "insert into cust_ord_dtl (Account_Code,Cust_Ref, Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,"
        strsql = strsql & " Cust_DrgNo,Cust_Drg_Desc,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,"
        strsql = strsql & " OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo, Consignee_Code,Unit_Code)"
        strsql = strsql & " (Select Account_Code,Cust_Ref, Amendment_No = '" & pstrAmendment & "',Item_Code,Rate,Order_Qty,Despatch_Qty = 0 ,"
        strsql = strsql & " Active_Flag ,Cust_DrgNo,Cust_Drg_Desc,Authorized_flag = 0 "
        strsql = strsql & " ,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,InternalSONo = '" & pstrInternalSONo & "',"
        strsql = strsql & " RevisionNo = " & pintRevisionNo & ",consignee_code,'" & gstrUNITID & "' from Cust_ord_dtl where Account_code = '" & pstrAccountCode & "' "
        strsql = strsql & " and cust_ref = '" & pstrRef & "' and Active_flag = 'A' and authorized_flag = 1 "
        strsql = strsql & " and amendment_no <> '" & pstrAmendment & "' and Unit_Code = '" & gstrUNITID & "' and Consignee_Code= '" & Me.txtConsCode.Text.Trim & "'"
        If SSPOentry.MaxRows > 0 Then
            intMaxLoop = SSPOentry.MaxRows
            strDrgItem = ""
            For intLoopCounter = 1 To intMaxLoop
                With SSPOentry
                    .Row = intLoopCounter
                    VarDelete = Nothing
                    Call .GetText(0, intLoopCounter, VarDelete)
                    varDrgNo = Nothing
                    Call .GetText(2, intLoopCounter, varDrgNo)
                    varItemCode = Nothing
                    Call .GetText(4, intLoopCounter, varItemCode)
                    If VarDelete <> "*" Then
                        If Len(Trim(strDrgItem)) > 0 Then
                            strDrgItem = Trim(strDrgItem) & " and (Cust_drgNo <> '" & varDrgNo & "' or Item_code <> '" & varItemCode & "')"
                        Else
                            strDrgItem = " and (Cust_drgNo <> '" & varDrgNo & "' or Item_code <> '" & varItemCode & "')"
                        End If
                    End If
                End With
            Next
            If Len(Trim(strDrgItem)) > 0 Then
                strsql = strsql & strDrgItem & ")"
            End If
        End If
        mP_Connection.BeginTrans()
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Public Function ToSetcustdrgDesc(ByRef plngRow As Integer) As String
        Dim varItemCode As Object
        Dim rsCustDrgDesc As ClsResultSetDB
        Dim varCPartCode As Object
        rsCustDrgDesc = New ClsResultSetDB
        With SSPOentry
            varCPartCode = Nothing
            Call .GetText(2, plngRow, varCPartCode)
            varItemCode = Nothing
            Call .GetText(4, plngRow, varItemCode)
            If Len(Trim(varCPartCode)) > 0 Then
                If Len(Trim(varItemCode)) Then
                    rsCustDrgDesc.GetResult("Select Drg_desc from CustItem_Mst where Active=1 and account_code =  '" & Trim(txtCustomerCode.Text) & "' and ITem_code = '" & Trim(varItemCode) & "' and Cust_drgNo = '" & Trim(varCPartCode) & "' and Unit_Code = '" & gstrUNITID & "'")
                    If rsCustDrgDesc.GetNoRows > 0 Then
                        rsCustDrgDesc.MoveFirst()
                        lblCustPartDesc.Text = rsCustDrgDesc.GetValue("Drg_desc")
                        ToSetcustdrgDesc = rsCustDrgDesc.GetValue("Drg_desc")
                    Else
                        lblCustPartDesc.Text = ""
                        ToSetcustdrgDesc = ""
                    End If
                End If
            End If
        End With
    End Function
    Private Sub SearchItem()
        '---------------------------------------------------------------------
        'Created By     -   Arshad Ali
        '---------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        With SSPOentry
            .Row = -1
            .Col = -1
            .Font = VB6.FontChangeBold(.Font, False)
            If optPartNo.Checked Then
                .Col = 2
            End If
            If optItem.Checked Then
                .Col = 4
            End If
            For intCount = 1 To .MaxRows
                .Row = intCount
                If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                    .TopRow = .Row
                    .Col = -1
                    .Font = VB6.FontChangeBold(.Font, True)
                    Exit Sub
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub InitializeSpreed()
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With Me.SSPOentry
            .Row = 0
            .Col = 19 : .Text = "MRP"
            .Row = 0
            .Col = 20 : .Text = "Abatment"
            .Row = 0
            .Col = 21 : .Text = "Accessible Rate"
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdButtons_MouseDown(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.MouseDownEventArgs) Handles cmdButtons.MouseDown
        On Error GoTo ErrHandler
        Select Case e.Index
            Case 0
                m_blnCloseFlag = True
            Case 4
                m_blnCloseFlag = True
            Case 6
                m_blnCloseFlag = True
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlPerValue_Change1(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.Change
        Dim intLoopCounter As Short
        Dim rsSalesParameter As New ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varRate As Object
        Dim varCustSupp As Object
        Dim varToolCost As Object
        Dim varOthers As Object
        With SSPOentry
            If Len(Trim(ctlPerValue.Text)) = 0 Then ctlPerValue.Text = 1
            If Val(ctlPerValue.Text) > 1 Then
                .Row = 0
                .Col = 6
                .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
            Else
                .Row = 0
                .Col = 6 : .Text = "Rate (Per Unit)"
            End If
            rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
            If rsSalesParameter.GetValue("ItemRateLink") = True Then
                If (Len(Trim(txtAmendmentNo.Text)) = 0) And (txtAmendmentNo.Enabled = False) Then
                    For intLoopCounter = 1 To SSPOentry.MaxRows
                        varDrgNo = Nothing
                        Call .GetText(2, intLoopCounter, varDrgNo)
                        varItemCode = Nothing
                        Call .GetText(4, intLoopCounter, varItemCode)
                        If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                            varRate = Nothing
                            Call .GetText(13, intLoopCounter, varRate)
                            Call .SetText(6, intLoopCounter, varRate * CDbl(ctlPerValue.Text))
                        End If
                    Next
                End If
            End If
        End With
    End Sub
    Private Sub ctlPerValue_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlPerValue.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                If cmbPOType.SelectedIndex = 4 Then
                    cmdButtons.Focus()
                Else
                    cmdchangetype.Focus()
                End If
        End Select
    End Sub
    Private Sub DTAmendmentDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTAmendmentDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = Keys.Enter Then
            Call DTAmendmentDate_Validating(DTAmendmentDate, New System.ComponentModel.CancelEventArgs(False))
            If blnValidAmendDate = True Then
                DTEffectiveDate.Focus()
            End If
        End If
        Select Case e.KeyCode
            Case 39, 34, 96
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTAmendmentDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTAmendmentDate.Validating
        On Error GoTo ErrHandler
        Application.DoEvents()
        If blnCheck = True Then
            blnCheck = False
            Exit Sub
        End If
        If blnLeavetxt = True Then
            blnLeavetxt = False
            Exit Sub
        End If
        If DTAmendmentDate.Value > DTValidDate.Value Then
            Call ConfirmWindow(10146, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            e.Cancel = False
            blnCheck = True
            Me.DTEffectiveDate.Focus()
            Exit Sub
        Else
            blnValidAmendDate = True
        End If
        If DTAmendmentDate.Value < DTDate.Value Then
            Call ConfirmWindow(10147, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            e.Cancel = True
            blnCheck = True
            Me.DTEffectiveDate.Focus()
            Exit Sub
        Else
            blnValidAmendDate = True
        End If
        'on requirement of MATE for Back Date SO Entry
        If DTAmendmentDate.Value > GetServerDate() Then
            MsgBox("Amendment Date Can not be Greater Than Current Date", vbInformation, ResolveResString(100))
            blnValidAmendDate = False
            e.Cancel = True
            blnCheck = True
            Me.DTEffectiveDate.Focus()
            Exit Sub
        Else
            blnValidAmendDate = True
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = Keys.Enter Then
            If DTAmendmentDate.Enabled = True Then
                DTAmendmentDate.Focus()
            Else
                If DTEffectiveDate.Enabled Then
                    DTEffectiveDate.Focus()
                Else
                    cmdButtons.Focus()
                End If
            End If
        End If
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTDate.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case keyascii
            Case Keys.Enter
            Case 39, 34, 96
                keyascii = 0
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTDate.LostFocus
        On Error GoTo ErrHandler
        m_strSql = "select * from company_mst where Unit_Code = '" & gstrUNITID & "'"
        rsdb.GetResult(m_strSql)
        m_blnDateFlag = False
        If DTDate.Value > DateValue(FinancialYearDates(FinancialYearDatesEnum.DATE_END)) Then
            Call ConfirmWindow(10074, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            DTDate.Focus()
            Exit Sub
        End If
        If DTDate.Value > GetServerDate() Then
            MsgBox("Date Can not be greater than Current Date")
            m_blnDateFlag = True
            DTDate.Focus()
            Exit Sub
        End If
        dtSODate = DTDate.Value
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTDate.ValueChanged
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
        'parameter check added by nisha on 16/16/2003
        If rsSalesParameter.GetValue("ItemRateLink") = True Then
            If DateDiff("d", dtSODate, DTDate.Value) <> 0 Then
                If MsgBox("Change in SO Date will remove the all Item Details from Grid.", vbYesNo, ResolveResString(100)) = vbYes Then
                    SSPOentry.MaxRows = 0
                    Call ADDRow()
                    DTDate.Focus()
                Else
                    DTDate.Value = GetServerDate() : DTDate.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub DTEffectiveDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTEffectiveDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = Keys.Enter Then
            If DTValidDate.Enabled = True Then
                DTValidDate.Focus()
            Else
                txtAmendReason.Focus()
            End If
        End If
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTEffectiveDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTEffectiveDate.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case keyascii
            Case 39, 34, 96
                keyascii = 0
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTEffectiveDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTEffectiveDate.Leave
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
        If Not cmdButtons.GetActiveButton Is Nothing Then
            If cmdButtons.GetActiveButton.Text.ToUpper = "close".ToUpper Then
                Exit Sub
            End If
        End If
        If blnCheck = True Then
            blnCheck = False
            blnLeavetxt = True
            Exit Sub
        End If
        m_strSql = "Select * from company_mst where Unit_Code = '" & gstrUNITID & "'"
        rsGetDate.GetResult(m_strSql)
        If DTEffectiveDate.Value < DTDate.Value And m_blnDateFlag <> True Then
            MsgBox("Effective Date Cannot Be Less than SO Date", vbInformation, ResolveResString(100))
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        If DTEffectiveDate.Value > DateValue(FinancialYearDates(FinancialYearDatesEnum.DATE_END)) Then
            MsgBox("Effective Date Cannot Be Greater than Financial End Date", vbInformation, ResolveResString(100))
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTValidDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = Keys.Enter Then
            If txtCurrencyType.Enabled = True Then
                txtCurrencyType.Focus()
            Else
                txtAmendReason.Focus()
            End If
        End If
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTValidDate.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case keyascii
            Case 39, 34, 96
                keyascii = 0
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTValidDate.Leave
        Call DTValidDate_Validating(DTValidDate, New System.ComponentModel.CancelEventArgs(False))
    End Sub
    Private Sub DTValidDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTValidDate.Validating
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
        If blnCheck = True Then
            Exit Sub
        End If
        m_strSql = "Select * from company_mst where Unit_Code = '" & gstrUNITID & "'"
        rsGetDate.GetResult(m_strSql)
        If DTValidDate.Value < FinancialYearDates(FinancialYearDatesEnum.DATE_START) Then
            Call ConfirmWindow(10073, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            e.Cancel = True
            DTValidDate.Value = FinancialYearDates(FinancialYearDatesEnum.DATE_START)
            blnCheck = True
            DTValidDate.Focus()
            blnCheck = False
            Exit Sub
        End If
        If DTValidDate.Value > FinancialYearDates(FinancialYearDatesEnum.DATE_END) Then
            Call ConfirmWindow(10074, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            e.Cancel = True
            DTValidDate.Value = FinancialYearDates(FinancialYearDatesEnum.DATE_END)
            blnCheck = True
            DTValidDate.Focus()
            blnCheck = False
            Exit Sub
        End If
        If DTValidDate.Value < GetServerDate() Then
            MsgBox("Valid Date Cannot be Less than Current Date.", vbOKOnly, ResolveResString(100))
            e.Cancel = True
            DTValidDate.Value = GetServerDate()
            blnCheck = True
            DTValidDate.Focus()
            blnCheck = False
            Exit Sub
        End If
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSPOentry_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SSPOentry.GotFocus
        If blnmsgbox1 = True Then
            blnmsgbox1 = False
            Exit Sub
        End If
        If Len(Trim(txtCurrencyType.Text)) = 0 Then ' for currency type Check
            If txtCurrencyType.Enabled = True Then
                txtCurrencyType.Focus()
                blnmsgbox1 = True
                MsgBox("Please Define Currency Code", vbInformation, "eMPro")
                txtCurrencyType.Focus()
                Exit Sub
            End If
            Exit Sub
        End If
    End Sub
    Private Sub SSPOentry_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SSPOentry.KeyDownEvent
        On Error GoTo ErrHandler
        Dim varHelpItem As Object
        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim rsAbatmentRate As New ClsResultSetDB
        Dim inti As Integer
        Dim varMRP, varAbatment, varAccessibleRateforMRP As Object
        Dim strSOEntry() As String
        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
        If (e.shift = 2 And e.keyCode = Keys.N) Then
            If ValidRowData(SSPOentry.ActiveRow, 0) Then Call ADDRow()
            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 2 : SSPOentry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End If
        If SSPOentry.ActiveCol = 2 Or SSPOentry.ActiveCol = 4 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If txtAmendmentNo.Enabled = False Then
                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a, Itemrate_Mst b, Item_MST as C where A.Item_code=C.Item_code and Status='A' and Hold_Flag=0 and a.Active=1 and a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code and i1.Unit_Code = '" & gstrUNITID & "') and datediff(mm,'" & FormatDateTime(DTDate.Value, vbLongDate) & "',b.DateFrom)<=0 and CustVend_Flg = 'C' AND a.Unit_code = c.Unit_Code and a.Unit_code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "'")
                    Else
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.Item_code=B.Item_code and B.Status='A' and B.hold_flag=0 and a.Active=1 and Account_Code='" & txtCustomerCode.Text & "' and a.Unit_code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "'")
                    End If
                Else
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst where  Active=1 and Account_Code='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'")
                End If
                If UBound(strSOEntry) <= 0 Then Exit Sub
                If strSOEntry(0) = "0" Then
                    Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    Call SSPOentry.SetText(2, SSPOentry.ActiveRow, strSOEntry(0))
                    Call SSPOentry.SetText(4, SSPOentry.ActiveRow, strSOEntry(1))
                    lblCustPartDesc.Text = strSOEntry(2)
                End If
                If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Dim rsitem As ClsResultSetDB
                    If Len(Trim(m_Item_Code)) > 0 Then
                        m_strSql = "SElect * from Cust_ord_dtl where Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1 and Unit_Code = '" & gstrUNITID & "'"
                        rsitem = New ClsResultSetDB
                        rsitem.GetResult(m_strSql)
                        If rsitem.GetNoRows > 0 Then
                            'change account Plug in
                            Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                            Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsitem.GetValue("Item_Code "))
                            Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsitem.GetValue("Order_Qty"))
                            Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                            Call SSPOentry.SetText(6, SSPOentry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                            Call SSPOentry.SetText(19, SSPOentry.MaxRows, rsitem.GetValue("MRP"))
                            Call SSPOentry.SetText(20, SSPOentry.MaxRows, rsitem.GetValue("abantment_code"))
                            Call SSPOentry.SetText(21, SSPOentry.MaxRows, rsitem.GetValue("AccessibleRateforMRP"))
                        Else
                            If txtAmendmentNo.Enabled = False Then
                                If Len(Trim(m_Item_Code)) > 0 Then
                                    'm_strSql = "SElect * from ITemRate_Mst where Item_code ='" & varHelpItem & " Party_Code ='" & txtCustomerCode.Text & "' and CustVend_Flag ='C'"
                                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                                        m_strSql = " select * from ITemRate_Mst where Serial_No = (select max(serial_no) from itemrate_mst where Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & DTDate.Value & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C' and Unit_Code = '" & gstrUNITID & "') and  Unit_Code = '" & gstrUNITID & "'"
                                        rsitem = New ClsResultSetDB
                                        rsitem.GetResult(m_strSql)
                                        If rsitem.GetNoRows > 0 Then
                                            'change account Plug in
                                            If chkOpenSo.Checked = False Then
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = False : SSPOentry.BlockMode = False
                                            Else
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = True : SSPOentry.BlockMode = False
                                            End If
                                            Call SSPOentry.SetText(2, SSPOentry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                            Call SSPOentry.SetText(4, SSPOentry.MaxRows, m_Item_Code)
                                            Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                                            Call SSPOentry.SetText(6, SSPOentry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                                            If rsitem.GetValue("Edit_flg") = False Then
                                                SSPOentry.Col = 6 : SSPOentry.Col2 = SSPOentry.MaxCols : SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.BlockMode = True : SSPOentry.Lock = True : SSPOentry.BlockMode = False
                                            Else
                                                SSPOentry.Col = 6 : SSPOentry.Col2 = SSPOentry.MaxCols : SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.BlockMode = True : SSPOentry.Lock = False : SSPOentry.BlockMode = False
                                            End If
                                        Else
                                            'If itemratelink in salesparameter is false then
                                            If chkOpenSo.Checked = False Then
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = False : SSPOentry.BlockMode = False
                                            Else
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Row2 = SSPOentry.MaxRows : SSPOentry.Col = 5 : SSPOentry.Col2 = 5 : SSPOentry.BlockMode = True : SSPOentry.Lock = True : SSPOentry.BlockMode = False
                                            End If
                                            Call SSPOentry.SetText(2, SSPOentry.MaxRows, varHelpItem)
                                            Call SSPOentry.SetText(4, SSPOentry.MaxRows, m_Item_Code)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf SSPOentry.ActiveCol = 20 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                With SSPOentry
                    .Row = .ActiveRow : .Col = .ActiveCol
                    varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='ABNT' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If varHelpItem = "" Or varHelpItem = "-1" Then
                        MsgBox("Abatment Code Does Not Exist", vbInformation, ResolveResString(100))
                    Else
                        Call SSPOentry.SetText(20, SSPOentry.ActiveRow, varHelpItem)
                        '''Accssible Rate to be calculated
                        varMRP = Nothing
                        varAbatment = Nothing
                        Call SSPOentry.GetText(19, SSPOentry.ActiveRow, varMRP)
                        Call SSPOentry.GetText(20, SSPOentry.ActiveRow, varAbatment)
                        m_strSql = "select txrt_percentage from Gen_TaxRate where Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        rsAbatmentRate = New ClsResultSetDB
                        rsAbatmentRate.GetResult(m_strSql)
                        If rsAbatmentRate.GetNoRows > 0 Then
                            varAbatment = Val(rsAbatmentRate.GetValue("txrt_percentage"))
                            varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                            Call SSPOentry.SetText(21, SSPOentry.ActiveRow, varAccessibleRateforMRP)
                        End If
                    End If
                End With
                rsAbatmentRate.ResultSetClose()
                rsAbatmentRate = Nothing
            End If
            If e.keyCode = 13 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                With SSPOentry
                    varMRP = Nothing
                    varAbatment = Nothing
                    Call SSPOentry.GetText(19, SSPOentry.ActiveRow, varMRP)
                    Call SSPOentry.GetText(20, SSPOentry.ActiveRow, varAbatment)
                    m_strSql = "select txrt_percentage from Gen_TaxRate where Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                    rsAbatmentRate = New ClsResultSetDB
                    rsAbatmentRate.GetResult(m_strSql)
                    If rsAbatmentRate.GetNoRows > 0 Then
                        varAbatment = Val(rsAbatmentRate.GetValue("txrt_percentage"))
                        varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                        Call SSPOentry.SetText(21, SSPOentry.ActiveRow, varAccessibleRateforMRP)
                    Else
                        If varMRP > 0 Then
                            MsgBox("Abatment Code Does Not Exist", vbInformation, ResolveResString(100))
                            Call SSPOentry.SetText(21, SSPOentry.ActiveRow, "")
                            'Call ssSetFocus(SSPOentry.ActiveRow, 20)
                            SSPOentry.Focus()
                            Exit Sub
                        End If
                    End If
                End With
                rsAbatmentRate.ResultSetClose()
                rsAbatmentRate = Nothing
            End If
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub SSPOentry_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SSPOentry.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96
                e.keyAscii = 0
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub SSPOentry_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SSPOentry.LeaveCell
        On Error GoTo ErrHandler
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        Dim rsAbatmentRate As New ClsResultSetDB
        Dim varMRP, varAbatment, varAccessibleRateforMRP As Object
        If e.newRow < 1 Then Exit Sub
        If ValidRowData(e.row, e.col) = True Then
            If (e.col = 2) Or (e.col = 4) Then
                With SSPOentry
                    .Col = 2 : .Row = e.row
                    If Len(Trim(.Text)) > 0 Then
                        .Col = 4 : .Row = e.row
                        If Len(Trim(.Text)) > 0 Then
                            If UCase(Trim(cmbPOType.Text)) <> "JOB WORK" Then
                                If SSPOentry.MaxRows < 1 Then
                                    Call ADDRow()
                                End If
                                SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 2 : SSPOentry.Focus()
                            End If
                        End If
                    End If
                End With
            End If
            If e.col = 19 Then
                varMRP = Nothing
                varAbatment = Nothing
                Call SSPOentry.GetText(19, e.row, varMRP)
                Call SSPOentry.GetText(20, e.row, varAbatment)
                If varAbatment <> "" Then
                    m_strSql = "select txrt_percentage from Gen_TaxRate where Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and Unit_Code = '" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                    rsAbatmentRate = New ClsResultSetDB
                    rsAbatmentRate.GetResult(m_strSql)
                    If rsAbatmentRate.GetNoRows > 0 Then
                        varAbatment = CLng(rsAbatmentRate.GetValue("txrt_percentage"))
                        varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                        Call SSPOentry.SetText(21, e.row, varAccessibleRateforMRP)
                    End If
                End If
            End If
        Else
            With SSPOentry
                'change account Plug in
                .Row = intRow : .Row2 = intRow : .Col = 11 : .Col2 = 11 : .BlockMode = True : .Lock = False : .BlockMode = False
            End With
        End If
        If (e.col = 1) Then
            With SSPOentry
                .Row = e.row : .Col = 1
                If .Value = 1 Then
                    .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Text = 0 : .Lock = True : .BlockMode = False
                Else
                    .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = False : .BlockMode = False
                End If
            End With
        End If
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If (e.col = 2) Or (e.col = 4) Then
                Dim strDrgNo As Object
                Dim GetDetails As Boolean
                Dim rsitem As ClsResultSetDB
                Dim strcustdtl As String
                Dim strItemCode As Object
                rsitem = New ClsResultSetDB
                If e.col = 2 Then
                    strDrgNo = Nothing
                    Call SSPOentry.GetText(e.col, SSPOentry.MaxRows, strDrgNo)
                    strcustdtl = "Select * from custITem_Mst where  Active=1 and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                    rsitem.GetResult(strcustdtl)
                    If rsitem.GetNoRows > 1 Then
                        GetDetails = False
                        MsgBox("This Part code has more then two Items linked, Please select one from Item ListBox", vbInformation, ResolveResString(100))
                        SetCellTypeCombo(e.row)
                        Call ssSetFocus(e.row, 4)
                        Exit Sub
                    Else
                        GetDetails = True
                        strItemCode = IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code"))
                        lblCustPartDesc.Text = rsitem.GetValue("Drg_desc")
                        Call SSPOentry.SetText(4, SSPOentry.MaxRows, strItemCode)
                        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
                        If rsSalesParameter.GetValue("ItemRateLink") = True Then
                            Call RowDetailsfromKeyBoard(strItemCode, strDrgNo)
                        End If
                    End If
                End If
                If e.col = 4 Then
                    Call SetCellStatic(e.row)
                End If
                If GetDetails = True Then
                    If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        m_strSql = "Select * from Cust_ord_dtl where Item_code ='" & strItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1 and  Unit_Code = '" & gstrUNITID & "'"
                    Else
                        m_strSql = "Select * from Cust_ord_dtl where Item_code ='" & strItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & txtReferenceNo.Text & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A' and Unit_Code = '" & gstrUNITID & "'"
                    End If
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
                        If rsitem.GetValue("OpenSO") = False Then
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 0
                        Else
                            SSPOentry.Row = SSPOentry.MaxRows : SSPOentry.Col = 1 : SSPOentry.Col2 = 1 : SSPOentry.Value = 1
                        End If
                        Call SSPOentry.SetText(2, SSPOentry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                        lblCustPartDesc.Text = rsitem.GetValue("Cust_Drg_desc")
                        Call SSPOentry.SetText(4, SSPOentry.MaxRows, rsitem.GetValue("Item_Code "))
                        Call SSPOentry.SetText(5, SSPOentry.MaxRows, rsitem.GetValue("Order_Qty"))
                        Call SSPOentry.SetText(13, SSPOentry.MaxRows, rsitem.GetValue("Rate"))
                        Call SSPOentry.SetText(6, SSPOentry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                        Call SSPOentry.SetText(17, SSPOentry.MaxRows, rsitem.GetValue("Despatch_qty"))
                        Call SSPOentry.SetText(19, SSPOentry.MaxRows, rsitem.GetValue("MRP"))
                        Call SSPOentry.SetText(20, SSPOentry.MaxRows, rsitem.GetValue("Abantment_code"))
                        Call SSPOentry.SetText(21, SSPOentry.MaxRows, rsitem.GetValue("AccessibleRateforMRP"))
                        Call ssSetFocus(SSPOentry.MaxRows, 3)
                    End If
                End If
            End If
        End If
        Call ToSetcustdrgDesc(e.newRow)
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtReferenceNo.Validating
        On Error GoTo ErrHandler
        Dim rsRefNo As New ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim strAns As MsgBoxResult
        Dim varPOFlg As Object
        Dim intAnswer As Short
        Dim intbase As Short
        mvalid = False
        If m_blnCloseFlag = True Then 'incase close button is clicked then exit
            m_blnCloseFlag = False
            Exit Sub
        End If
        If m_blnHelpFlag = True Then 'incase help button is clicked then exit
            m_blnHelpFlag = False
            Exit Sub
        End If
        If Len(Trim(txtReferenceNo.Text)) > 0 Then 'if reference no is not blank
            m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'"
            Call rsRefNo.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If rsRefNo.GetNoRows = 1 Then
                    intbase = 1
                End If
                If rsRefNo.GetNoRows > 1 Then
                    intAnswer = MsgBox("Would You Like to View Base SO", MsgBoxStyle.YesNo, "eMPower")
                    If intAnswer = 6 Then
                        txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        intbase = 1
                    Else
                        txtAmendmentNo.Enabled = True : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        intbase = rsRefNo.GetNoRows
                    End If
                End If
            Else
                intbase = rsRefNo.GetNoRows
            End If
            If intbase = 1 Then 'if only one record for the reference no is existing
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and (Authorized_flag=1 or Future_SO =1) and Unit_Code = '" & gstrUNITID & "'"
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        strAns = ConfirmWindow(10131, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        If strAns = MsgBoxResult.Yes Then
                            txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            cmdHelp(3).Enabled = False
                            DTAmendmentDate.Value = GetServerDate()
                            Call GetReferenceDetails()
                            SSPOentry.MaxRows = 1
                            Call SSMaxLength()
                            cmdchangetype.Enabled = True
                            If txtAmendmentNo.Enabled Then txtAmendmentNo.Focus()
                            mvalid = False
                            Exit Sub
                        Else
                            Call GetReferenceDetails()
                            cmdButtons.Enabled(0) = True
                            cmdButtons.Enabled(1) = False
                            cmdButtons.Enabled(2) = False
                            cmdButtons.Enabled(3) = True
                            cmdButtons.Enabled(4) = False
                            cmdButtons.Enabled(5) = True
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10132, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetReferenceDetails()
                        mvalid = True
                        cmdButtons.Revert()
                        cmdButtons.Enabled(0) = True
                        cmdButtons.Enabled(1) = False
                        cmdButtons.Enabled(2) = False
                        cmdButtons.Enabled(3) = False
                        cmdButtons.Enabled(4) = False
                        cmdButtons.Enabled(5) = False
                        cmdButtons.Enabled(6) = True
                        cmdForms.Enabled = False
                        Exit Sub
                    End If
                ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No ='' and Unit_Code = '" & gstrUNITID & "'"
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and Authorized_flag=0 and amendment_No ='' and Unit_Code = '" & gstrUNITID & "'"
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then
                            m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and future_SO=1 and amendment_No ='' and Unit_Code = '" & gstrUNITID & "'"
                            Call rsRefNo.GetResult(m_strSql)
                            If rsRefNo.GetNoRows > 0 Then
                                MsgBox("This Is Future SO(AUTHORISED)", MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            Call GetReferenceDetails()
                            cmdButtons.Enabled(0) = True
                            cmdButtons.Enabled(1) = True
                            cmdButtons.Enabled(2) = True
                            cmdButtons.Enabled(3) = True
                            cmdButtons.Enabled(4) = False
                            cmdButtons.Enabled(5) = True
                        Else
                            Call ConfirmWindow(10133, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            Call GetReferenceDetails()
                            mvalid = True
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10134, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetReferenceDetails()
                        mvalid = True
                        Exit Sub
                    End If
                End If
                ' If No of records for the reference no is more than 1 that means amendment no exists
            ElseIf intbase > 1 Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rsSalesParameter.GetResult("Select AppendSOItem from Sales_parameter where Unit_Code = '" & gstrUNITID & "'")
                    If rsSalesParameter.GetValue("AppendSOItem") = True Then
                        'incase an amendment already exists and is not authorized
                        m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and (Authorized_flag=0 and future_SO = 0) and Unit_Code = '" & gstrUNITID & "'"
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then 'incase a not authorized amendment exists
                            cmdButtons.Focus()
                            Call ConfirmWindow(10135, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            txtReferenceNo.Text = ""
                            txtReferenceNo.Focus()
                            mvalid = True
                            Exit Sub
                        Else
                            Me.txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            Call GetReferenceDetails()
                            SSPOentry.MaxRows = 1
                            Call SSMaxLength()
                            Exit Sub
                        End If
                    Else
                        Me.txtAmendmentNo.Enabled = True
                        txtAmendmentNo.BackColor = System.Drawing.Color.White
                        Call GetReferenceDetails()
                        SSPOentry.MaxRows = 1
                        Call SSMaxLength()
                        Exit Sub
                    End If
                Else
                    txtAmendmentNo.Enabled = True
                    txtAmendmentNo.BackColor = System.Drawing.Color.White
                    cmdHelp(3).Enabled = True
                    If txtAmendmentNo.Enabled Then
                        txtAmendmentNo.Focus()
                    End If
                    Exit Sub
                End If
            Else 'If There are no records existing
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then 'If Mode is new,then
                    With SSPOentry
                        .Col = 19
                        .Col2 = 19
                        .ColHidden = True
                        .Col = 20
                        .Col2 = 20
                        .ColHidden = True
                        .Col = 21
                        .Col2 = 21
                        .ColHidden = True
                    End With
                    DTDate.Enabled = True
                    DTValidDate.Enabled = True
                    DTEffectiveDate.Enabled = True
                    txtCurrencyType.Enabled = True
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(2).Enabled = True
                    cmbPOType.Enabled = True
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    SSPOentry.Enabled = True
                    cmdchangetype.Enabled = True
                    txtAmendmentNo.Enabled = False
                    txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(3).Enabled = False
                    txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True
                    cmdHelp(7).Enabled = True
                    cmdForms.Enabled = True
                    ctlPerValue.Enabled = True
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ctlPerValue.Text = 1
                    chkOpenSo.Enabled = True
                    Call ADDRow()
                    DTDate.Focus()
                    Exit Sub
                Else
                    DTDate.Enabled = False
                    Call ConfirmWindow(10136, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    If txtReferenceNo.Enabled Then txtReferenceNo.Focus()
                    mvalid = True
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        mvalid = False
        If StrComp(mstrPrevRefNo, Trim(txtReferenceNo.Text), CompareMethod.Text) <> 0 Then
            Call CLEARVAR()
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        On Error GoTo ErrHandler
        Dim inti As Short
        Dim rsAmend As New ClsResultSetDB
        mvalid = False
        If m_blnHelpFlag = True Then
            m_blnHelpFlag = False
            Exit Sub
        End If
        If m_blnCloseFlag = True Then
            m_blnCloseFlag = False
            Exit Sub
        End If
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = "Select * from cust_ord_hdr where Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "' "
            rsAmend.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If rsAmend.GetNoRows > 0 Then
                    cmdButtons.Focus()
                    Call ConfirmWindow(10141, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    mvalid = True
                    txtAmendmentNo.Focus()
                    Exit Sub
                Else
                    DTAmendmentDate.Enabled = True
                    DTEffectiveDate.Enabled = True
                    DTValidDate.Enabled = True
                    txtAmendReason.Enabled = True
                    txtAmendReason.BackColor = System.Drawing.Color.White
                    cmdchangetype.Enabled = True
                    cmdForms.Enabled = True
                    txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True : cmdHelp(7).Enabled = True

                    chkOpenSo.Enabled = True
                    With Me.SSPOentry
                        .Enabled = True
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 5
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                End If
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                m_strSql = "Select * from cust_ord_hdr where Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'"
                rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then
                    m_strSql = "Select * from cust_ord_hdr where Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' and Unit_Code = '" & gstrUNITID & "' "
                    rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then
                        m_strSql = "Select * from cust_ord_hdr where Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' and (authorized_Flag=1)  and Unit_Code = '" & gstrUNITID & "'"
                        rsAmend.GetResult(m_strSql)
                        If rsAmend.GetNoRows > 0 Then
                            Call ConfirmWindow(10142, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdButtons.Focus()
                            Call GetAmendmentDetails()
                            mvalid = True
                            cmdButtons.Focus()
                            Exit Sub
                        Else
                            m_strSql = "Select * from cust_ord_hdr where Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' and (future_so =1) and Unit_Code = '" & gstrUNITID & "'"
                            rsAmend.GetResult(m_strSql)
                            If rsAmend.GetNoRows > 0 Then
                                MsgBox("This is Future SO(AUTHORISED).", MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            Call GetAmendmentDetails()
                            With SSPOentry
                                .Col = 1
                                .Col2 = .MaxCols
                                .BlockMode = True
                                .Lock = True
                                .BlockMode = False
                            End With
                            cmdButtons.Enabled(0) = True
                            cmdButtons.Enabled(1) = True
                            cmdButtons.Enabled(2) = True
                            cmdButtons.Enabled(3) = True
                            cmdButtons.Enabled(4) = False
                            cmdButtons.Enabled(5) = False
                        End If
                    Else
                        cmdButtons.Focus()
                        Call ConfirmWindow(10143, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetAmendmentDetails()
                        mvalid = True
                        cmdButtons.Focus()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    mvalid = True
                    txtAmendmentNo.Focus()
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        mvalid = False
        cmdButtons.Enabled(5) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub


    Private Sub txtConsCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB
        If Len(Trim(txtConsCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            'issue id : 10224351
            m_strSql = "Select * from Customer_mst where customer_Code='" & Trim(txtConsCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtConsCode.Text = ""
                Cancel = True
                txtConsCode.Focus()
                GoTo EventExitSub
            Else
                txtConsCode.Text = IIf(UCase(rsCD.GetValue("customer_Code")) = "UNKNOWN", "", rsCD.GetValue("customer_Code"))
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        e.Cancel = Cancel
    End Sub

    Private Sub _cmdHelp_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _cmdHelp_6.Click

    End Sub

    Private Sub txtSTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTax.Leave
        If Len(Trim(txtSTax.Text)) <> 0 Then
            m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "'"
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("STAX")
            Else
                MsgBox("Entered S.Tax Code does not exist", MsgBoxStyle.Information, "empower")
                txtSTax.Text = ""
                txtSTax.Focus()
            End If
            rsdb.ResultSetClose()
        End If
    End Sub

    Private Sub txtSTax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTax.TextChanged
        Call FillLabel("STAX")
    End Sub

End Class