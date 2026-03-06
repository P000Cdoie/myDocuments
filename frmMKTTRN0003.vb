Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0003
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'
	'File Name          :   frmMKTTRN0003.frm
	'Function           :   Customer PO Lock
	'Created By         :   Meenu Gupta
	'Created on         :   9, April 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 2
	'30/10/2001 CHANGED TO ALLOW UNLOCK
	'17-01-2002 internal issue log no = 55,56 - checked out form no =4009
	'25-01-2002 done to allow 4 decimal places in Rate for MSSL-ED - checked out form no =4018
	'28-02-02 Changed lable Surcharge % on formNo 4052 in grid control
	'14/05/2002 for check if no item select & pressed lock
	'13/09/2002 changed by nisha for accounts Plugin
	'08/11/2002 Changed by nisha to add
	'AutoGeneration No in SO
	'12/12/2002 Changed by nisha to add
	'1. Sorting on Customer Part Code
	'26/05/2003 changes done by nisha
	'to lock Base PO if its amendment is already there
	'****Field Added by Ajay on 21/07/2003    for Surcharge on S.Tax
	'20/07/2004 Changes done by Brij Bohara
	'1) to remove the error for HILEX----Can't set Value to NULL when CheckBox property = FALSE.
	'2) and to fixing the bug of lblIntSONo disappearing when text cleared from txtReferenceNo text box
	
	'28/12/2004 Changes done by Brij Bohara
	'to display Remarks in Grid
	'10/01/2005 Does't ALLOW UNLOCK  Change Done BY Rajani Kant
	
    '---------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    05/05/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------

	Dim m_blnCloseButton As Boolean
    Dim rsdb As ClsResultSetDB
	Dim mintFormIndex, intRow As Short
	Dim mstrCode As String
	Dim m_blnChangeFormFlg As Boolean
	Dim m_strSql, strSQL As String
	Dim m_blnGetAmendmentDetails As Boolean
    Dim rsRefNo As ClsResultSetDB
	Dim m_blnhelp As Boolean
    Dim m_ItemDesc, m_custItemDesc As String

    Private Sub chkSelect_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSelect.CheckStateChanged
        Dim Index As Short = chkSelect.GetIndex(eventSender)
        On Error GoTo errHandler
        Select Case Index
            Case 0
                ssPOEntry.Col = 16
                If chkSelect(0).CheckState = 1 Then
                    chkSelect(1).CheckState = System.Windows.Forms.CheckState.Unchecked
                    For intRow = 1 To ssPOEntry.MaxRows
                        ssPOEntry.Row = intRow
                        If ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox Then
                            Call ssPOEntry.SetText(16, intRow, 1)
                        End If
                    Next
                End If
            Case 1
                If chkSelect(1).CheckState = 1 Then
                    chkSelect(0).CheckState = System.Windows.Forms.CheckState.Unchecked
                    ssPOEntry.Col = 16
                    For intRow = 1 To ssPOEntry.MaxRows
                        ssPOEntry.Row = intRow
                        If ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox Then
                            Call ssPOEntry.SetText(16, intRow, 0)
                        End If
                    Next
                End If
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        On Error GoTo errHandler
        m_pstrSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND  Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'"
        frmMKTTRNAdditionalDetails.ShowDialog()
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        On Error GoTo errHandler
        Dim varRetVal As Object
        Select Case Index
            Case 0
                With Me.txtCustomerCode
                    If Len(.Text) = 0 Then
                        'for accounts Plug in customer code change
                        varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 1
                With Me.txtReferenceNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag = 'A'")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag = 'A' ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 2
                With txtCurrencyType
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 3
                With txtAmendmentNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No <> ' ' and Active_Flag = 'A'")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and  cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No <> ' ' and Active_Flag = 'A'")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10013, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdHelp_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdHelp.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        Select Case Index
            Case 1
                m_blnhelp = True
            Case 3
                m_blnhelp = True
        End Select
    End Sub
    Private Sub frmMKTTRN0003_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0003_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlHeader_Click(ctlHeader, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0003_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            '     SendKeys "{Tab}"
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAmendmentNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendmentNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            If cmdchangetype.Enabled Then cmdchangetype.Focus() Else cmdAuthorize.Focus()
        ElseIf KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtAmendmentNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
    End Sub
    Private Sub txtCreditTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.TextChanged
        Call FillLabel("CREDIT")
    End Sub
    Private Sub txtCurrencyType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus() Else cmdAuthorize.Focus()
        End If
    End Sub
    Private Sub txtCustomerCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo errHandler
        Dim rsCD As New ClsResultSetDB
        If m_blnCloseButton = True Then
            m_blnCloseButton = False
            GoTo EventExitSub
        End If
        If m_blnhelp = True Then
            m_blnhelp = False
            GoTo EventExitSub
        End If
        If Me.ActiveControl.Name = cmdAuthorize.Name Then Exit Sub
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            MsgBox("Customer Code can not be blank", MsgBoxStyle.OkOnly, ResolveResString(100))
            txtCustomerCode.Focus()
            Cancel = True
            GoTo EventExitSub
        Else
            m_strSql = "Select TOP 1 1 from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                'MsgBox ("Customer Code Does Not Exist")
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                GoTo EventExitSub
            End If
        End If
        cmdHelp(1).Enabled = True
        Call FillLabel("Customer")
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub TxtReferenceNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtReferenceNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            If txtAmendmentNo.Enabled = True Then
                txtAmendmentNo.Focus()
            Else
                cmdAuthorize.Focus()
            End If
        ElseIf KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtReferenceNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
    End Sub

    Private Sub frmMKTTRN0003_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0003_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim dtServerdate As Date
        DTDate.Format = DateTimePickerFormat.Custom
        DTDate.CustomFormat = gstrDateFormat

        DTValidDate.Format = DateTimePickerFormat.Custom
        DTValidDate.CustomFormat = gstrDateFormat

        DTAmendmentDate.Format = DateTimePickerFormat.Custom
        DTAmendmentDate.CustomFormat = gstrDateFormat

        DTEffectiveDate.Format = DateTimePickerFormat.Custom
        DTEffectiveDate.CustomFormat = gstrDateFormat

        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        Call FillLabelFromResFile(Me)
        'Size the form to client workspace
        Call FitToClient(Me, fraContainer, ctlHeader, cmdAuthorize, 400)
        'Disabling the controls
        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 0)
        Call EnableControls(False, Me, True)
        'Initialising the buttons
        cmdAuthorize.Caption(0) = "Lock"
        'Disabling Authorize, Refresh  buttons
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        txtCustomerCode.BackColor = System.Drawing.Color.White
        ssPOEntry.Enabled = False
        m_blnhelp = False
        Call AddPOType()
        With ssPOEntry
            .Col = 13
            .Col2 = 13
            .ColHidden = False
            .Col = 14
            .Col2 = 14
            .ColHidden = False
            .Row = 0
            .Col = 16
            .Col2 = 16 : .Text = "LOCK"
        End With
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        dtServerdate = GetServerDate()
        Me.DTDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTAmendmentDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTEffectiveDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTValidDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
    End Sub
    Private Sub frmMKTTRN0003_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo errHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = True
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0003_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Releasing the form reference
        Me.Dispose()
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
    End Sub
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '----------------------------------------------------------------------------
        'Argument       :   Row, Col
        'Return Value   :   Nil
        'Function       :   Set the focus according to row and col value
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        With ssPOEntry
            .Row = Row
            .Col = Col
            .Action = 0
        End With
    End Sub
    Public Function FillLabel(ByRef pstrCode As Object) As Object
        On Error GoTo errHandler
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   fills the customer detail label
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim rsCust As ClsResultSetDB

        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where UNIT_CODE='" & gstrUNITID & "' AND  Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                rsCust.ResultSetClose()
            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where UNIT_CODE='" & gstrUNITID & "' AND Txrt_Rate_No = '" & txtSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
                '1.Surcharge on S.Tax
            Case "SSCHTAX"
                m_strSql = "SELECT TxRt_Rate_no, TxRt_RateDesc FROM gen_taxrate where UNIT_CODE='" & gstrUNITID & "' AND Txrt_Rate_No= '" & txtSChSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSChSTaxDesc.ForeColor = System.Drawing.Color.White

                lblSChSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
                '------------------->>
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where UNIT_CODE='" & gstrUNITID & "' AND crtrm_TermID = '" & txtCreditTerms.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
                rsCust.ResultSetClose()
        End Select
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsAD As New ClsResultSetDB
        Dim varLockFlag As Object
        Dim rscurrency As ClsResultSetDB
        Dim strpotype As String
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and authorized_Flag=1"
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty from cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Authorized_Flag=1 order by cust_drgno"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date") 'VB6.Format(rsAD.GetValue("Order_Date"), "dd/mm/yyyy")
            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date") 'VB6.Format(rsAD.GetValue("Amendment_Date"), "dd/mm/yyyy")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date") 'VB6.Format(rsAD.GetValue("Effect_Date"), "dd/mm/yyyy")
            DTValidDate.Value = rsAD.GetValue("Valid_Date") 'VB6.Format(rsAD.GetValue("Valid_Date"), "dd/mm/yyyy")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            With ssPOEntry
                .Col = 13
                .Col2 = 13
                .ColHidden = True
                .Col = 14
                .Col2 = 14
                .ColHidden = True
                .Col = 15
                .Col2 = 15
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
                    With ssPOEntry
                        .Col = 13
                        .Col2 = 13
                        .ColHidden = False
                        .Col = 14
                        .Col2 = 14
                        .ColHidden = False
                        .Col = 15
                        .Col2 = 15
                        .ColHidden = False
                    End With
            End Select
            'to show the details of Sales Tax,Credit Days,AddCustSupplied Flag,Open SO Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            ssPOEntry.MaxRows = 0
            Do While Not rsdb.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                ssPOEntry.Col = 1
                ssPOEntry.Col2 = 1
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                ssPOEntry.BlockMode = False
                ssPOEntry.Col = 10
                ssPOEntry.Col2 = 10
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                'Changed to add open Item Falg in Grid
                If rsdb.GetValue("OpenSO") = False Then
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsdb.GetValue("Cust_DrgNo"))
                Call ssPOEntry.SetText(3, ssPOEntry.MaxRows, rsdb.GetValue("Item_Code "))
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Drg_Desc"))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb.GetValue("Order_Qty"))

                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
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
                For intLoopCounter = 6 To 9
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMax = strMax
                        .TypeFloatMin = strMin
                    End With
                Next
                With Me.ssPOEntry
                    .Row = .MaxRows
                    .Col = 13
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 14
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 15
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 16
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                End With

                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
                varLockFlag = rsdb.GetValue("Active_Flag")
                If varLockFlag = "L" Then
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, 1)
                End If
                rsdb.MoveNext()
            Loop
        End If
        With ssPOEntry
            .BlockMode = True
            .Col = 1
            .Col2 = 12
            .Row = 1
            .Row2 = .MaxRows
            .Lock = True
            .BlockMode = False
            .Enabled = True
            .ColsFrozen = 3
        End With
        cmdchangetype.Enabled = True
        chkSelect(0).Enabled = True
        chkSelect(1).Enabled = True
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub GetDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is no amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim strAuthFlg, strpotype As String
        Dim varLockFlag As Object
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and authorized_Flag= 1 and amendment_no = '" & Trim(txtAmendmentNo.Text) & "'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and authorized_Flag = 1 and amendment_no = '" & Trim(txtAmendmentNo.Text) & "' Order by cust_drgNo"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date") ' VB6.Format(rsAD.GetValue("Order_Date"), "dd/mm/yyyy")
            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            If rsAD.GetValue("Amendment_Date") = "" Or IsDBNull(rsAD.GetValue("Amendment_Date")) = True Then
            Else
                DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date") ' VB6.Format(rsAD.GetValue("Amendment_Date"), "dd/mm/yyyy")
            End If
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date") ' VB6.Format(rsAD.GetValue("Effect_Date"), "dd/mm/yyyy")
            DTValidDate.Value = rsAD.GetValue("Valid_Date") 'VB6.Format(rsAD.GetValue("Valid_Date"), "dd/mm/yyyy")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            With ssPOEntry
                .Col = 13
                .Col2 = 13
                .ColHidden = True
                .Col = 14
                .Col2 = 14
                .ColHidden = True
                .Col = 15
                .Col2 = 15
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
                    With ssPOEntry
                        .Col = 13
                        .Col2 = 13
                        .ColHidden = False
                        .Col = 14
                        .Col2 = 14
                        .ColHidden = False
                        .Col = 15
                        .Col2 = 15
                        .ColHidden = False
                    End With
            End Select
            'to show the details of Sales Tax,Credit Days,AddCustSupplied Flag,Open SO Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            Do While Not rsdb.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                ssPOEntry.Col = 1
                ssPOEntry.Col2 = 1
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                ssPOEntry.BlockMode = False
                ssPOEntry.Col = 10
                ssPOEntry.Col2 = 10
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                'Changed to add open Item Falg in Grid
                If rsdb.GetValue("OpenSO") = False Then
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsdb.GetValue("Cust_DrgNo"))
                Call ssPOEntry.SetText(3, ssPOEntry.MaxRows, rsdb.GetValue("Item_Code"))
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Drg_Desc"))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb.GetValue("Order_Qty"))
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                    rscurrency.ResultSetClose()
                End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                For intLoopCounter = 6 To 9
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMax = strMax
                        .TypeFloatMin = strMin
                    End With
                Next
                With Me.ssPOEntry
                    .Row = .MaxRows
                    .Col = 13
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 14
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 15
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 16
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                End With
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
                varLockFlag = rsdb.GetValue("Active_Flag")
                If varLockFlag = "L" Then
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, 1)
                End If
                rsdb.MoveNext()
            Loop
        End If
        With ssPOEntry
            .BlockMode = True
            .Col = 1
            .Col2 = 12
            .Row = 1
            .Row2 = .MaxRows
            .Lock = True
            .BlockMode = False
            .Enabled = True
            .ColsFrozen = 3
        End With
        cmdchangetype.Enabled = True
        chkSelect(0).Enabled = True
        chkSelect(1).Enabled = True
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            Call RefreshForm()
            txtAmendmentNo.Text = ""
            txtSTax.Text = "" : txtCreditTerms.Text = "" : txtSChSTax.Text = ""
            chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdHelp(3).Enabled = False
            txtAmendmentNo.Enabled = False
            txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
    End Sub
    Public Sub RefreshForm()
        Dim dtServerdate As Date
        lblIntSONoDes.Text = ""
        txtCurrencyType.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        dtServerdate = GetServerDate()
        Me.DTDate.Value = dtServerdate ' VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTAmendmentDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTEffectiveDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTValidDate.Value = dtServerdate 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        ssPOEntry.MaxRows = 0
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        txtSTax.Text = "" : txtCreditTerms.Text = "" : txtSChSTax.Text = ""
        chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            Call RefreshForm()
        End If
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo errHandler
        Call FillLabel("Customer")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            txtReferenceNo.Enabled = True
            txtReferenceNo.BackColor = System.Drawing.Color.White
            cmdHelp(1).Enabled = True
        Else
            Call RefreshForm()
            lblCustDesc.Text = ""
            txtReferenceNo.Text = ""
            cmdHelp(1).Enabled = False
            cmdHelp(3).Enabled = False
            txtReferenceNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtAmendmentNo.Text = ""
            txtAmendmentNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        ssPOEntry.MaxRows = 0
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub AddPOType()
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "MRP-SPARES")
    End Sub
    Private Sub txtSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSTax.TextChanged
        Call FillLabel("STAX")
    End Sub
    Private Sub txtSChSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSChSTax.TextChanged
        Call FillLabel("SSCHTAX")
    End Sub
    Private Sub cmdAuthorize_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdAuthorize.ButtonClick
        On Error GoTo errHandler
        Dim strSQL As String
        Dim strErrMsg As String
        Dim blnUpdateHdrTable As Boolean
        Dim strAns As MsgBoxResult
        Dim vartext, varItemCode, varCustDrgNo As Object
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim blnItemSelectedtoLock As Boolean
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'Lock PO
                enmValue = ConfirmWindow(10164, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Are You Sure To Lock the PO
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    '******  to check atleast one Item Should be selected
                    For intRow = 1 To ssPOEntry.MaxRows
                        vartext = Nothing
                        Call ssPOEntry.GetText(16, intRow, vartext)
                        If Val(vartext) = 1 Then
                            blnItemSelectedtoLock = True
                            Exit For
                        End If
                    Next
                    If blnItemSelectedtoLock = True Then
                        blnUpdateHdrTable = True
                        For intRow = 1 To ssPOEntry.MaxRows
                            vartext = Nothing
                            Call ssPOEntry.GetText(16, intRow, vartext)
                            varItemCode = Nothing
                            Call ssPOEntry.GetText(3, intRow, varItemCode)
                            varCustDrgNo = Nothing
                            Call ssPOEntry.GetText(2, intRow, varCustDrgNo)
                            If Val(vartext) = 1 Then 'Update Cust_ord_dtl table
                                m_strSql = "Update cust_ord_dtl set Active_Flag ='L' where unit_code='" & gstrUNITID & "' and "
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Item_Code='"
                                m_strSql = m_strSql & varItemCode & "' and cust_drgno='" & varCustDrgNo & "'  and authorized_Flag=1"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        Next
                        '******* to check if any item in grid is not selected to lock
                        For intRow = 1 To ssPOEntry.MaxRows
                            vartext = Nothing
                            Call ssPOEntry.GetText(16, intRow, vartext)
                            If Val(vartext) <> 1 Then
                                blnUpdateHdrTable = False
                                Exit For
                            End If
                        Next
                        If blnUpdateHdrTable = True Then 'Update Cust_ord_hdr table
                            m_strSql = "Update cust_ord_hdr set Active_Flag ='L' where unit_code='" & gstrUNITID & "' and "
                            m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                            m_strSql = m_strSql & " and Account_Code='"
                            m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                            m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                            mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        Call EnableControls(False, Me, True)
                        txtCustomerCode.Enabled = True
                        txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdHelp(0).Enabled = True
                        chkSelect(0).Enabled = False
                        chkSelect(1).Enabled = False
                        lblCustDesc.Text = ""
                        ssPOEntry.MaxRows = 0
                        txtCustomerCode.Focus()
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                    Else
                        MsgBox("No Item Selected for Lock", MsgBoxStyle.Information, "empower")
                        chkSelect(0).Focus()
                        If chkSelect(0).Enabled Then chkSelect(0).Focus()
                        Exit Sub
                    End If
                Else
                    txtCustomerCode.Focus()
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH 'Refresh Screen
                Call EnableControls(False, Me, True)
                ssPOEntry.MaxRows = 0
                txtAmendmentNo.Enabled = True
                cmdHelp(0).Enabled = True
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.Color.White
                cmdHelp(0).Enabled = True
                lblCustDesc.Text = ""
                txtCustomerCode.Focus()
                cmdAuthorize.Enabled(0) = False
                cmdAuthorize.Enabled(1) = False
                cmdAuthorize.Enabled(2) = False
                cmdAuthorize.Enabled(3) = True
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'Close
                Me.Close()
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub cmdAuthorize_MouseDown(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.MouseDownEventArgs) Handles cmdAuthorize.MouseDown
        Select Case e.Index
            Case 3
                m_blnCloseButton = True
        End Select
    End Sub
    Private Sub ctlPerValue_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.Change
        With ssPOEntry
            If Len(Trim(ctlPerValue.Text)) = 0 Then ctlPerValue.Text = 1
            If Val(ctlPerValue.Text) > 1 Then
                .Row = 0
                .Col = 6
                .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 7
                .Text = "Cust Supp Mat (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 8
                .Text = "Tool Cost (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 11
                .Text = "Others (Per " & Val(ctlPerValue.Text) & ")"
            Else
                .Row = 0
                .Col = 6 : .Text = "Rate (Per Unit)"
                .Row = 0
                .Col = 7 : .Text = "Cust Supp Mat. (Per Unit)"
                .Row = 0
                .Col = 8 : .Text = "Tool Cost (Per Unit)"
                .Row = 0
                .Col = 11 : .Text = "Others (Per Unit)"
            End If

            .Row = 0
            .Col = 13 : .Text = "MRP"
            .Row = 0
            .Col = 14 : .Text = "Abantment"
            .Row = 0
            .Col = 15 : .Text = "Accessible Rate"
        End With

    End Sub
    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("salesorderlock_trans(mkt).htm")
    End Sub
    Private Sub txtReferenceNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtReferenceNo.Validating
        On Error GoTo errHandler
        Dim rsBasePO As ClsResultSetDB
        Dim intAns As Short
        If m_blnCloseButton = True Then
            m_blnCloseButton = False
            Exit Sub
        End If
        If m_blnhelp = True Then
            m_blnhelp = False
            Exit Sub
        End If
        Dim inti As Short
        If Len(Trim(txtReferenceNo.Text)) > 0 Then
            ' Check if records for the entered reference no exist or not
            m_strSql = " Select 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'"
            rsRefNo = New ClsResultSetDB
            Call rsRefNo.GetResult(m_strSql)
            If rsRefNo.GetNoRows = 1 Then ' If there are records existing for the entered reference no
                rsRefNo.ResultSetClose()
                ' check whether the PO is Authorized or not
                rsRefNo = New ClsResultSetDB
                m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Authorized_Flag =1"
                Call rsRefNo.GetResult(m_strSql)
                If rsRefNo.GetNoRows > 0 Then
                    rsRefNo.ResultSetClose()
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='L' and Authorized_Flag =1"
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made and is already authorized and locked
                        rsRefNo.ResultSetClose()
                        Call GetDetails()
                        Call ConfirmWindow(10196, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        chkSelect(0).Enabled = True
                        chkSelect(1).Enabled = True
                        ssPOEntry.Focus()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        Exit Sub
                    Else
                        rsRefNo.ResultSetClose()
                        Call GetDetails()
                        chkSelect(0).Enabled = True
                        chkSelect(1).Enabled = True
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        cmdAuthorize.Focus()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10197, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    txtReferenceNo.Focus()
                    rsRefNo.ResultSetClose()
                    Exit Sub
                End If
            ElseIf rsRefNo.GetNoRows > 1 Then  ' If An amendment exists for the reference no
                rsRefNo.ResultSetClose()
                m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and amendment_No = '' and active_Flag = 'A'"
                rsBasePO = New ClsResultSetDB
                rsBasePO.GetResult(m_strSql)
                If rsBasePO.GetNoRows > 0 Then
                    rsBasePO.ResultSetClose()
                    intAns = MsgBox("Would you like to lock Base SO ?", MsgBoxStyle.YesNo, "empower")
                    If intAns = 6 Then
                        txtAmendmentNo.Enabled = False
                        txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        cmdHelp(3).Enabled = False
                        m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Authorized_Flag =1 and amendment_no = ''"
                        rsBasePO = New ClsResultSetDB
                        rsBasePO.GetResult(m_strSql)
                        If rsBasePO.GetNoRows > 0 Then
                            rsBasePO.ResultSetClose()
                            Call GetDetails()
                            chkSelect(0).Enabled = True
                            chkSelect(1).Enabled = True
                            cmdAuthorize.Enabled(0) = True
                            cmdAuthorize.Enabled(1) = True
                            cmdAuthorize.Enabled(2) = False
                            cmdAuthorize.Enabled(3) = True
                            cmdAuthorize.Focus()
                            Exit Sub
                        Else
                            Call ConfirmWindow(10197, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            txtReferenceNo.Text = ""
                            txtReferenceNo.Focus()
                            rsBasePO.ResultSetClose()
                            Exit Sub
                        End If
                    Else
                        ssPOEntry.MaxRows = 0
                        txtAmendmentNo.Text = ""
                        txtAmendmentNo.Enabled = True
                        txtAmendmentNo.BackColor = System.Drawing.Color.White
                        cmdHelp(3).Enabled = True
                        txtAmendmentNo.Focus()
                        Exit Sub
                    End If
                Else
                    txtAmendmentNo.Enabled = True
                    txtAmendmentNo.BackColor = System.Drawing.Color.White
                    cmdHelp(3).Enabled = True
                    txtAmendmentNo.Focus()
                    rsBasePO.ResultSetClose()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtReferenceNo.Text = ""
                txtReferenceNo.Focus()
                rsRefNo.ResultSetClose()
                Exit Sub
            End If
        ElseIf Len(Trim(txtReferenceNo.Text)) = 0 Then
            Exit Sub
        Else
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub txtAmendmentNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        On Error GoTo errHandler
        If m_blnCloseButton = True Then
            m_blnCloseButton = False
            Exit Sub
        End If
        If m_blnhelp = True Then
            m_blnhelp = False
            Exit Sub
        End If
        Dim inti As Short
        Dim rsAmend As ClsResultSetDB
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and amendment_No='" & txtAmendmentNo.Text & "'"
            rsAmend = New ClsResultSetDB
            Call rsAmend.GetResult(m_strSql)
            If rsAmend.GetNoRows > 0 Then ' If there are records existing for the entered Amendment no
                rsAmend.ResultSetClose()
                ' check whether the PO is Authorized or not
                rsAmend = New ClsResultSetDB
                m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Authorized_Flag= 1 and amendment_No='" & txtAmendmentNo.Text & "'"
                Call rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then ' If An amendment exists for the reference no
                    rsAmend.ResultSetClose()
                    m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='L' and amendment_No='" & txtAmendmentNo.Text & "' and Authorized_Flag= 1"
                    rsAmend = New ClsResultSetDB
                    Call rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then ' If the amendment is Already Locked
                        rsAmend.ResultSetClose()
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        MsgBox(" This Amendment has all/some Locked items")
                        'call ConfirmWindow (,BUTTON_OK ,IMG_INFO )
                        chkSelect(0).Enabled = True
                        chkSelect(1).Enabled = True
                        Exit Sub
                    Else
                        rsAmend.ResultSetClose()
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        cmdAuthorize.Focus()
                        chkSelect(0).Enabled = True
                        chkSelect(1).Enabled = True
                        Exit Sub
                    End If
                Else
                    MsgBox("This Amendment Is Not Authorized") ' If PO is not Authorized
                    txtAmendmentNo.Text = ""
                    txtAmendmentNo.Focus()
                    rsAmend.ResultSetClose()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtAmendmentNo.Text = ""
                txtAmendmentNo.Focus()
                rsAmend.ResultSetClose()
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        rsAmend.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
End Class