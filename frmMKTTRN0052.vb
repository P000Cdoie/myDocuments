Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0052
	Inherits System.Windows.Forms.Form
    '*****************************************************************************************************
	'(C) 2001 MIND, All rights reserved
    'File Name          :   frmMKTTRN0052.frm
	'Function           :   Sales Order Authorization
	'Created By         :   Manoj Kr. vaish
    'Created on         :   27 April 2007
    '---------------------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   17/05/2011
    'Modified to support MultiUnit functionality
    '---------------------------------------------------------------------------------------
    '************************************************************************************************
    Dim rsdb As ClsResultSetDB
	Dim m_CloseButton As Boolean
	Dim m_blnhelp As Boolean
	Dim mintFormIndex, intRow As Short
	Dim m_strSql, strsql As String
    Dim rsRefNo As ClsResultSetDB
	Dim m_ItemDesc, m_custItemDesc As String
	Dim strpotype As String
	Dim blnValidAmend As Boolean
	Dim blnValidCust As Boolean
    Dim blnValidref As Boolean
    Dim blnLeaveRef As Boolean = False
    Dim blnAuth As Boolean = False
	
	Private Sub cmbPOType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbPOType.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		On Error GoTo errHandler
		If cmbPOType.Text = "MRP-SPARES" Then
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 13
            ssPOEntry.Col2 = 13
            ssPOEntry.ColHidden = False
            ssPOEntry.Col = 14
            ssPOEntry.Col2 = 14
            ssPOEntry.ColHidden = False
            ssPOEntry.BlockMode = False
            ssPOEntry.Col = 15
            ssPOEntry.Col2 = 15
            ssPOEntry.ColHidden = False
            ssPOEntry.BlockMode = False
        Else
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 13
            ssPOEntry.Col2 = 13
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 14
            ssPOEntry.Col2 = 14
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 15
            ssPOEntry.Col2 = 15
            ssPOEntry.ColHidden = True
            ssPOEntry.BlockMode = False
		End If
		GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
		GoTo EventExitSub
		
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
	
    Private Sub cmdAuthorize_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdAuthorize.ButtonClick
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Functionality of Authorize/Refresh/Close
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim strsql As String
        Dim strErrMsg As String
        Dim strAns As MsgBoxResult
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varCustDrgNo As Object
        Dim dblDespatch As Double
        Dim rsDespatch As New ClsResultSetDB
        Select Case eventArgs.button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'Authorize PO
                If ValidAuthorize() = True Then
                    enmValue = ConfirmWindow(10163, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Are You Sure To Authorize the PO
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        If DTEffectiveDate.Value > GetServerDate() Then
                            Call ConfirmWindow(10198, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else

                            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FormatDateTime(DTEffectiveDate.Value, DateFormat.LongDate)), CDate(FormatDateTime(GetServerDate, DateFormat.LongDate))) >= 0 Then
                                'Update Cust_ord_hdr table
                                m_strSql = "Update cust_ord_hdr set First_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Second_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Third_Authorized ='"
                                m_strSql = m_strSql & mP_User & "', Authorized_Flag =1 where"
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                'Update Cust_ord_dtl table
                                m_strSql = "Update cust_ord_dtl set Authorized_Flag =1 where"
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                If ssPOEntry.MaxRows > 0 Then
                                    intMaxLoop = ssPOEntry.MaxRows
                                    For intLoopCounter = 1 To intMaxLoop
                                        With ssPOEntry
                                            varCustDrgNo = Nothing
                                            varItemCode = Nothing
                                            Call .GetText(2, intLoopCounter, varCustDrgNo)
                                            Call .GetText(3, intLoopCounter, varItemCode)
                                            m_strSql = "Select Despatch_Qty from cust_ord_dtl Where Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                                            m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'"
                                            m_strSql = m_strSql & " and amendment_no <> '" & Trim(txtAmendmentNo.Text) & "' and active_flag = 'A'"
                                            m_strSql = m_strSql & " and Authorized_flag =1 and ITem_code = '" & Trim(varItemCode) & "'"
                                            m_strSql = m_strSql & " and Cust_drgno = '" & Trim(varCustDrgNo) & "'"
                                            m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                            rsDespatch = New ClsResultSetDB
                                            rsDespatch.GetResult(m_strSql)
                                            If rsDespatch.GetNoRows > 0 Then
                                                rsDespatch.MoveFirst()
                                                dblDespatch = rsDespatch.GetValue("Despatch_Qty")
                                                m_strSql = "update cust_ord_dtl set Despatch_qty = " & dblDespatch & " Where Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                                                m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'"
                                                m_strSql = m_strSql & " and amendment_no = '" & Trim(txtAmendmentNo.Text) & "' and active_flag = 'A'"
                                                m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                                m_strSql = m_strSql & " and ITem_code = '" & Trim(varItemCode) & "' and Cust_drgno = '" & Trim(varCustDrgNo) & "'"
                                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                            rsDespatch.ResultSetClose()
                                            rsDespatch = Nothing
                                        End With
                                    Next
                                End If
                                m_strSql = "Update cust_ord_dtl set Active_Flag = 'O' where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                m_strSql = m_strSql & " and Account_Code = '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No <> '" & Trim(Me.txtAmendmentNo.Text) & "' and authorized_flag =1 and Active_Flag <> 'L' "
                                m_strSql = m_strSql & " and cust_drgno in(select cust_drgno from   cust_ord_dtl  where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code= '" & gstrUNITID & "' and Account_Code= '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No='" & Trim(Me.txtAmendmentNo.Text) & "')"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Call UpdateHdrActiveFlag()
                            Else
                                m_strSql = "Update cust_ord_hdr set First_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Second_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Third_Authorized ='"
                                m_strSql = m_strSql & mP_User & "', future_so =1 where"
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                m_strSql = m_strSql & " and unit_code='" & gstrUNITID & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '************
                            End If
                        End If
                        MsgBox("SO Authorized Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                        blnAuth = True
                        Call EnableControls(False, Me, True)
                        blnAuth = False
                        txtCustomerCode.Enabled = True
                        cmdHelp(0).Enabled = True
                        ssPOEntry.MaxRows = 0
                        Me.txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        Call RefreshForm()
                        txtCustomerCode.Focus()
                    Else
                        If txtAmendmentNo.Enabled = True Then
                            txtAmendmentNo.Focus()
                            Exit Sub
                        Else
                            txtReferenceNo.Focus()
                            Exit Sub
                        End If
                    End If
                Else
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH 'Refresh Screen
                Call EnableControls(False, Me, True)
                ssPOEntry.MaxRows = 0
                cmdHelp(0).Enabled = True
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
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
    Private Sub cmdAuthorize_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.MouseDownEventArgs) Handles cmdAuthorize.MouseDown
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Value of Public Variable Used in Diffrent
        '                    - Cases.
        '-----------------------------------------------------------------------
        Select Case eventArgs.Index
            Case 3
                m_CloseButton = True
        End Select
    End Sub
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Sub Form for Payment Terms
        '-----------------------------------------------------------------------
        m_pstrSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' and Unit_code='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'"
        frmMKTTRNAdditionalDetails.ShowDialog()
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show the Help Form
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim varRetVal As Object
        Select Case Index
            Case 0
                With Me.txtCustomerCode
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 1
                With Me.txtReferenceNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", "order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) and Rtrim(Active_Flag) Not In('L','O')")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", "order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and future_so = 0 and (authorized_flag = 0 or future_so = 0) And Rtrim(Active_Flag) In ('L','O')")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 3
                With txtAmendmentNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No <>' ' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) and isnull(Active_Flag,'') Not In('L','O') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and  cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & txtCustomerCode.Text & "'  and Amendment_No <>' ' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) And isnull(Active_Flag,'') Not In('L','O') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Value of Public Variable
        '-----------------------------------------------------------------------
        Select Case Index
            Case 0
                m_blnhelp = True
        End Select
    End Sub
    Private Sub ctlHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlHeader.Click
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Empower Help on F4 Click
        '-----------------------------------------------------------------------
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("salesorderauth_trans(mkt).htm")
    End Sub
    Private Sub ctlPerValue_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlPerValue.Change
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Change The Values Grid Labeles on Change of
        '                    - Per Value
        '-----------------------------------------------------------------------
        With ssPOEntry
            If Len(Trim(ctlPerValue.Text)) = 0 Then ctlPerValue.Text = 1
            If Val(ctlPerValue.Text) > 1 Then
                .Row = 0
                .Col = 6
                .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
            Else
                .Row = 0
                .Col = 6 : .Text = "Rate (Per Unit)"
            End If
        End With
    End Sub
    Private Sub frmMKTTRN0052_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Required code on Form Deactivate
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0052_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call empower help on F4 Click
        '-----------------------------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlHeader_ClickEvent(ctlHeader, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh Values on Change in Amendment Text Box
        '-----------------------------------------------------------------------
        txtCreditTerms.Text = ""
        chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            Call RefreshForm()
        End If
    End Sub
    Private Sub txtAmendmentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtAmendmentNo_Validating(txtAmendmentNo, New System.ComponentModel.CancelEventArgs(False))
            If blnValidAmend = True Then
                If cmdchangetype.Enabled Then cmdchangetype.Focus() Else cmdAuthorize.Focus()
            End If
        End If
    End Sub
    Private Sub txtCreditTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.TextChanged
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh Label of Credit Description
        '-----------------------------------------------------------------------
        Call FillLabel("CREDIT")
    End Sub
    Private Sub txtCurrencyType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh The Data on Change of Customer Code
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Call FillLabel("CUSTOMER")
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
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code to Validate Entered Customer Code
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            If blnValidCust = True Then
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus() Else cmdAuthorize.Focus()
            End If
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate Customer Code Entered by User
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsCD As New ClsResultSetDB
        blnValidCust = False
        If m_CloseButton = True Then
            m_CloseButton = False
            GoTo EventExitSub
        End If
        If m_blnhelp = True Then
            m_blnhelp = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            txtReferenceNo.Enabled = True
            txtReferenceNo.BackColor = System.Drawing.Color.White
            cmdHelp(1).Enabled = True
            Call FillLabel("CUSTOMER")
            m_strSql = "Select * from Customer_Mst where Unit_Code = '" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCD = New ClsResultSetDB
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                txtCustomerCode.Text = ""
                blnValidCust = False
                GoTo EventExitSub
            Else
                blnValidCust = True
            End If
            rsCD.ResultSetClose()
            rsCD = Nothing
        End If
        blnValidCust = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the Data on Changeing Referance
        '                    - No.
        '-----------------------------------------------------------------------
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            Call RefreshForm()
            lblCustDesc.Text = ""
            txtAmendmentNo.Text = ""
            txtCreditTerms.Text = ""
            chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdHelp(3).Enabled = False
            txtAmendmentNo.Enabled = False
            txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
    End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdAuthorize.Focus()
        End If
    End Sub
    Private Sub txtReferenceNo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.Leave
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Dispaly Details on Lost Focus of Referance No
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        If blnAuth = True Then
            Exit Sub
        End If
        If blnLeaveRef = True Then
            Exit Sub
        End If
        Dim inti As Short
        If m_CloseButton = True Then
            m_CloseButton = False
            Exit Sub
        End If
        blnValidref = False
        If Len(Trim(txtReferenceNo.Text)) > 0 Then
            ' Check if records for the entered reference no exist or not
            m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'"
            'rsRefNo.ResultSetClose()
            rsRefNo = New ClsResultSetDB
            Call rsRefNo.GetResult(m_strSql)
            If rsRefNo.GetNoRows = 1 Then ' If there are records existing for the entered reference no
                ' check whether the PO is active or not
                m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Unit_Code='" & gstrUNITID & "' and Active_Flag='A'"
                rsRefNo.ResultSetClose()
                rsRefNo = New ClsResultSetDB
                Call rsRefNo.GetResult(m_strSql)
                If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                    m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and (Authorized_Flag =1 or future_so = 1) "
                    rsRefNo.ResultSetClose()
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                        Call GetDetails()
                        Call ConfirmWindow(10161, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        blnLeaveRef = True
                        txtReferenceNo.Focus()
                        blnLeaveRef = False
                        blnValidref = False
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        rsRefNo.ResultSetClose()
                        Exit Sub
                    Else
                        Call GetDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        blnLeaveRef = True
                        cmdAuthorize.Focus()
                        blnLeaveRef = False
                        blnValidref = True
                        rsRefNo.ResultSetClose()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10162, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    blnLeaveRef = True
                    txtReferenceNo.Focus()
                    blnLeaveRef = False
                    blnValidref = False
                    rsRefNo.ResultSetClose()
                    Exit Sub
                End If
            ElseIf rsRefNo.GetNoRows > 1 Then  ' If An amendment exists for the reference no
                txtAmendmentNo.Enabled = True
                txtAmendmentNo.BackColor = System.Drawing.Color.White
                cmdHelp(3).Enabled = True
                If txtAmendmentNo.Enabled = True Then
                    blnLeaveRef = True
                    txtAmendmentNo.Focus()
                    blnLeaveRef = False
                End If
                blnValidref = True
            Else
                Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtReferenceNo.Text = ""
                blnLeaveRef = True
                txtReferenceNo.Focus()
                blnLeaveRef = False
                blnValidref = False
                rsRefNo.ResultSetClose()
                Exit Sub
            End If
            rsRefNo.ResultSetClose()
        End If
        blnValidref = True
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0052_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Activate
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0052_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Load
        '-----------------------------------------------------------------------
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        'Load the captions
        Call FillLabelFromResFile(Me)
        'Size the form to client workspace
        Call FitToClient(Me, fraContainer, ctlHeader, cmdAuthorize, 400)
        'Disabling the controls
        Call EnableControls(False, Me, True)
        'Initialising the buttons
        'Disabling Authorize, Refresh  buttons
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        txtCustomerCode.BackColor = System.Drawing.Color.White
        Call AddPOType()
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        Me.DTDate.Format = DateTimePickerFormat.Custom
        Me.DTDate.CustomFormat = gstrDateFormat
        Me.DTDate.Value = GetServerDate()
        Me.DTAmendmentDate.Format = DateTimePickerFormat.Custom
        Me.DTAmendmentDate.CustomFormat = gstrDateFormat
        Me.DTAmendmentDate.Value = GetServerDate()
        Me.DTEffectiveDate.Format = DateTimePickerFormat.Custom
        Me.DTEffectiveDate.CustomFormat = gstrDateFormat
        Me.DTEffectiveDate.Value = GetServerDate()
        Me.DTValidDate.Format = DateTimePickerFormat.Custom
        Me.DTValidDate.CustomFormat = gstrDateFormat
        Me.DTValidDate.Value = GetServerDate()
        ssPOEntry.Enabled = False
        m_CloseButton = False
        Call InitializeSpreed()
        With ssPOEntry
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
    End Sub
    Private Sub frmMKTTRN0052_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form_QueryUnload
        '-----------------------------------------------------------------------
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            Cancel = 1
        End If
        eventArgs.Cancel = Cancel
        Me.Dispose()
    End Sub
    Private Sub frmMKTTRN0052_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form_Unload
        '-----------------------------------------------------------------------
        'Releasing the form reference
        frmMKTTRN0002 = Nothing
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
    End Sub
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - Set the focus according to row and col value
        '-----------------------------------------------------------------------
        With ssPOEntry
            .Row = Row
            .Col = Col
            .Action = 0
        End With
    End Sub
    Private Sub txtAmendmentNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Details of Amendment No on Validate
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim inti As Short
        Dim rsAmend As New ClsResultSetDB
        blnValidAmend = False
        If m_CloseButton = True Then
            m_CloseButton = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and amendment_No='" & txtAmendmentNo.Text & "'"
            Call rsAmend.GetResult(m_strSql)
            If rsAmend.GetNoRows > 0 Then ' If there are records existing for the entered Amendment no
                ' check whether the PO is active or not
                m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "'"
                rsAmend.ResultSetClose()
                rsAmend = New ClsResultSetDB
                Call rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then ' If An amendment exists for the reference no
                    m_strSql = " Select * from cust_ord_hdr where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "' and (Authorized_Flag= 1 or future_so =1)"
                    rsAmend.ResultSetClose()
                    rsAmend = New ClsResultSetDB
                    Call rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then ' If the amendment is Already Authorized
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        Call ConfirmWindow(10161, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Cancel = True
                        blnValidAmend = False
                        rsAmend.ResultSetClose()
                        GoTo EventExitSub
                    Else
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        blnValidAmend = True
                        rsAmend.ResultSetClose()
                        GoTo EventExitSub
                    End If
                Else
                    Call ConfirmWindow(10160, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    Cancel = True
                    blnValidAmend = False
                    rsAmend.ResultSetClose()
                    GoTo EventExitSub
                End If
            Else
                Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtAmendmentNo.Text = ""
                Cancel = True
                blnValidAmend = False
                rsAmend.ResultSetClose()
                GoTo EventExitSub
            End If
            rsAmend.ResultSetClose()
        End If
        blnValidAmend = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function FillLabel(ByRef pstrCode As Object) As Object
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - fills the customer detail label
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsCust As New ClsResultSetDB
        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where Customer_code='" & txtCustomerCode.Text & "' and Unit_Code='" & gstrUNITID & "' "
                'rsCust.ResultSetClose()
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where crtrm_TermID = '" & txtCreditTerms.Text & "' and Unit_Code='" & gstrUNITID & "'"
                rsCust.ResultSetClose()
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
        End Select
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetails()
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - Get the details if there is an amendment
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        m_strSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and amendment_No='" & txtAmendmentNo.Text & "' and ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Active_Flag='A'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and amendment_No='" & txtAmendmentNo.Text & "' and ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Active_Flag='A' order by cust_drgno"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            strpotype = rsAD.GetValue("PO_Type")
            txtsalestaxtype.Text = rsAD.GetValue("SalesTax_Type")

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
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
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
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where currency_code ='" & Trim(txtCurrencyType.Text) & "' and Unit_Code='" & gstrUNITID & "'")
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
                For intLoopCounter = 6 To 8
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
                    .Col = 9
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 13
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Row = .MaxRows
                    .Col = 14
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                End With
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
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
        End With
        cmdchangetype.Enabled = True
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        rsdb = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub GetDetails()
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - Get the details if there is no amendment
        '-----------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select * from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date")
            If Len(Trim(rsAD.GetValue("Amendment_No"))) > 0 Then
                DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            End If
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            strpotype = rsAD.GetValue("PO_Type")
            txtsalestaxtype.Text = rsAD.GetValue("SalesTax_Type")
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
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            'Changed to add open Item Falg in Grid
            Do While Not rsdb.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                ssPOEntry.Col = 1
                ssPOEntry.Col2 = 1
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
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
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where currency_code ='" & Trim(txtCurrencyType.Text) & "' and Unit_Code='" & gstrUNITID & "'")
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
                For intLoopCounter = 6 To 8
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
                    .Col = 9
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
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
                End With
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
                rsdb.MoveNext()
            Loop
        End If
        With ssPOEntry
            'Changed to add open Item Falg in Grid
            .BlockMode = True
            .Col = 1
            .Col2 = 11
            .Row = 1
            .Row2 = .MaxRows
            .Lock = True
            .BlockMode = False
            .Enabled = True
        End With
        cmdchangetype.Enabled = True
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        rsdb = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub AddPOType()
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add SO Type in Combo
        '-----------------------------------------------------------------------
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "MRP-SPARES")
    End Sub
    Public Sub RefreshForm()
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh The Form
        '-----------------------------------------------------------------------
        txtCurrencyType.Text = ""
        txtsalestaxtype.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        Me.DTDate.Value = GetServerDate()
        Me.DTAmendmentDate.Value = GetServerDate()
        Me.DTEffectiveDate.Value = GetServerDate()
        Me.DTValidDate.Value = GetServerDate()
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        ssPOEntry.MaxRows = 0
    End Sub
    Public Function ValidAuthorize() As Boolean
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Check Before Authorization
        '-----------------------------------------------------------------------
        ValidAuthorize = False
        Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
        If blnValidCust = True Then
            Call txtReferenceNo_Leave(txtReferenceNo, New System.EventArgs())
            If blnValidref = True Then
                If Len(Trim(txtAmendmentNo.Text)) <> 0 Then
                    Call txtAmendmentNo_Validating(txtAmendmentNo, New System.ComponentModel.CancelEventArgs(False))
                    If blnValidAmend = True Then
                        ValidAuthorize = True
                    Else
                        Exit Function
                    End If
                Else
                    ValidAuthorize = True
                End If
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End Function
    Public Sub UpdateHdrActiveFlag()
        '-----------------------------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To update Active Flag To "O" on Authorization
        '-----------------------------------------------------------------------
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim intmaxitems As Short
        Dim intMaxOverItem As Short
        Dim AmendmentNo As String
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdDtl = New ClsResultSetDB
        m_strSql = "select distinct(AmendMEnt_No)from cust_ord_hdr where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsCustOrdHdr.GetResult(m_strSql)
        intMaxLoop = rsCustOrdHdr.GetNoRows
        rsCustOrdHdr.MoveFirst()
        For intLoopCounter = 1 To intMaxLoop
            AmendmentNo = Trim(rsCustOrdHdr.GetValue("Amendment_No"))
            m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "' and active_Flag='O'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intMaxOverItem = rsCustOrdDtl.GetNoRows
            m_strSql = "select * from cust_ord_dtl where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "'"
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intmaxitems = rsCustOrdDtl.GetNoRows
            If intmaxitems = intMaxOverItem Then
                m_strSql = "Update cust_ord_hdr set active_Flag='O' where Cust_ref='" & txtReferenceNo.Text & "' and Unit_Code='" & gstrUNITID & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and amendment_no ='" & AmendmentNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            rsCustOrdHdr.MoveNext()
            rsCustOrdDtl.ResultSetClose()
        Next
        rsCustOrdHdr.ResultSetClose()
    End Sub
    Private Sub InitializeSpreed()
        On Error GoTo errHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With Me.ssPOEntry
            .MaxCols = 16
            .Row = 0
            .Col = 13 : .Text = "MRP"
            .Row = 0
            .Col = 14 : .Text = "Abantment"
            .Row = 0
            .Col = 15 : .Text = "Accessible Rate"
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
errHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
End Class