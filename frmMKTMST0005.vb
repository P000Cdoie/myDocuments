Option Strict Off
Option Explicit On
Friend Class frmMKTMST0005
	Inherits System.Windows.Forms.Form
	'****************************************************
	'Copyright (c)  -  MIND
	'Name of module -  frmMKTMST0005.frm
	'Created By     -  Kapil
	'Created Date   -  24 - 05 - 2001
	'description    -  Sales Configuration
    'Revised date   -   9/01/2002, nimesh
    'Revised Date   - 15/01/2002, Nisha on Checkedout Form No =4003 on version No 2
    'Revised Date   - 28/01/2002, Nisha on Checkedout Form No =4025 on version No 3
    'Revised Date   - 15/03/2002, Nisha on Checkedout Form No =4065
    'Revised Date   - 05/04/2002, Nisha on Checkedout Form No =4080 to Add new feild Openning Balence
    '***** 05/04/2002, Nisha on Checkedout Form No feilds Pre Printed Flag & No of Copies
    '16/10/2002 changes done by Nisha to Allow "_" in ReportFile Name Feild
    'CHANGES DONE BY NISHA ON 24/03/2003 FOR FINANCIAL ROLLOVER 09/04/2003
    'Changes Done By Arshad Ali 28/04/2004 to verify existance of record in SalesChallan_dtl, function for the purpose is CheckDocNoFromSalesChallanDetail
    'Modified by Amit Rana on 25/April/2011 for multi unit change
    '****************************************************


    Dim mintIndex As Short 'Declared to hold the Form Count
    Private Const mlng_SAVEBEFOREEXIT As Short = 9 'To use if user selects Save, when Exiting
    Dim mstrDocumentno As String
    Dim mstrLocation_Code As String
    Dim mstrCategory As String
    Dim mstrInvType As String
    Dim strInsert As String
    Dim mstrSubType As String
    Dim EOU_Flag As Boolean
    Dim blnWhetherEnteredOrNOt As Boolean ' to check whether cost center detail is entered ot not
    Dim strFinancialYearStart As String
    Dim strFinancialYearEnd As String
    Dim strPreFinYear As String

    Private Sub CboFinancialYear_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CboFinancialYear.SelectedIndexChanged
        Select Case Trim(Me.CmdGrpSaleConf.lbMode)
            Case "View"
                If CboFinancialYear.Text <> strPreFinYear Then
                    txtLocationCode.Text = "" : txtInvoiceType.Text = ""
                    txtInvSubType.Text = "" : txtPurposeCode.Text = ""
                    txtRptName.Text = ""
                    txtStockLocation.Text = "" : ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Unchecked : ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    lblInvTypeDes.Text = "" : lblInvSubTypeDes.Text = ""
                    lblPurposeDes.Text = ""
                    txtCurrentNo.Text = ""
                End If
        End Select
    End Sub
    Private Sub CboFinancialYear_DropDown(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CboFinancialYear.DropDown
        strPreFinYear = CboFinancialYear.Text
    End Sub
    Private Sub CboFinancialYear_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CboFinancialYear.Enter
        strPreFinYear = CboFinancialYear.Text
    End Sub
    Private Sub CboFinancialYear_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CboFinancialYear.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "View"
                        txtLocationCode.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkPre_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPre.Enter
        Shape2.Visible = True
    End Sub
    Private Sub chkPre_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkPre.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call chkPre_Leave(chkPre, New System.EventArgs())
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkPre_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPre.Leave
        If chkPre.CheckState = System.Windows.Forms.CheckState.Checked Then
            ctlNoOfCopies.Text = 1
            ctlNoOfCopies.Enabled = False
            ctlNoOfCopies.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            If ctlNoOfCopies.Enabled = True Then
                ctlNoOfCopies.Focus()
            ElseIf ChkSameSeries.Enabled = True Then
                ChkSameSeries.Focus()
            Else
                spdCostCenter.Col = 3
                spdCostCenter.Row = 1
                spdCostCenter.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdCostCenter.EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
                spdCostCenter.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdCostCenter.Focus()
            End If
        Else
            ctlNoOfCopies.Enabled = True
            ctlNoOfCopies.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ctlNoOfCopies.Focus()
            ctlNoOfCopies.Focus()
        End If
        Shape2.Visible = False
    End Sub
    Private Sub ChkSameSeries_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkSameSeries.Enter
        Shape1.Visible = True
    End Sub
    Private Sub ChkSameSeries_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles ChkSameSeries.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        spdCostCenter.Col = 3
        spdCostCenter.Row = 1
        spdCostCenter.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        spdCostCenter.EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
        spdCostCenter.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        spdCostCenter.Focus()
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ChkSameSeries_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkSameSeries.Leave
        Shape1.Visible = False
    End Sub
    Private Sub ChkUpdatePOFlag_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkUpdatePOFlag.Enter
        Shape3.Visible = True
    End Sub
    Private Sub ChkUpdatePOFlag_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkUpdatePOFlag.Leave
        Shape3.Visible = False
    End Sub
    Private Sub ChkUpdateStockFlag_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkUpdateStockFlag.Enter
        Shape4.Visible = True
    End Sub
    Private Sub ChkUpdateStockFlag_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkUpdateStockFlag.Leave
        Shape4.Visible = False
    End Sub
    Private Sub CmdAddOne_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdAddOne.Click
        On Error GoTo ErrHandler
        txtCurrentNo.Text = Val(CStr(CDbl(txtCurrentNo.Text) + 1))
        If CheckDocNoFromSalesChallanDetail() Then
            MsgBox("An invoice exists with " & Val(txtCurrentNo.Text) + 1 & " document number." & vbCrLf & "Please skip either backward or forward.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdInvoiceType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdInvoiceType.Click
        On Error GoTo ErrHandler
        'Check Location Code whether entered or not
        If Len(txtLocationCode.Text) = 0 Then
            Call ConfirmWindow(10474, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtLocationCode.Focus()
            Exit Sub
        End If
        Dim StrCheckHelp As String
        If Len(Me.txtInvoiceType.Text) = 0 Then
            Call ToGetYearStartandYearEnd()
            StrCheckHelp = ShowList(1, (txtInvoiceType.MaxLength), "", "Invoice_Type", "Description", "SaleConf", " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtInvoiceType.Text = StrCheckHelp
            End If
        Else
            Call ToGetYearStartandYearEnd()
            StrCheckHelp = ShowList(1, (txtInvoiceType.MaxLength), txtInvoiceType.Text, "Invoice_Type", "Description", "SaleConf", " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
            'change ends here on 24/03/2003
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtInvoiceType.Text = StrCheckHelp
            End If
        End If
        Call SelectDescriptionForInvType("Description", lblInvTypeDes, StrCheckHelp, "Invoice_Type")
        txtInvoiceType.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdInvSubType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdInvSubType.Click
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   Display Help for Invoice Type from SaleConf
        '*****************************************************
        On Error GoTo ErrHandler
        'Check Location Code whether entered or not
        If Len(txtLocationCode.Text) = 0 Then
            Call ConfirmWindow(10474, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtLocationCode.Focus()
            Exit Sub
        End If
        If Len(txtInvoiceType.Text) = 0 Then
            Call ConfirmWindow(10475, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtInvoiceType.Focus()
            Exit Sub
        End If
        Dim StrCheckHelp As String
        If Len(Me.txtInvSubType.Text) = 0 Then
            Call ToGetYearStartandYearEnd()
            StrCheckHelp = ShowList(1, (txtInvSubType.MaxLength), "", "Sub_Type", "Sub_Type_Description", "SaleConf", " and Invoice_Type in('" & Trim(txtInvoiceType.Text) & "') and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtInvSubType.Focus()
                Exit Sub
            Else
                txtInvSubType.Text = StrCheckHelp
            End If
        Else
            Call ToGetYearStartandYearEnd()
            StrCheckHelp = ShowList(1, (txtInvSubType.MaxLength), txtInvSubType.Text, "Sub_Type", "Sub_Type_Description", "SaleConf", "and Invoice_Type in('" & Trim(txtInvoiceType.Text) & "') and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtInvSubType.Focus()
                Exit Sub
            Else
                txtInvSubType.Text = StrCheckHelp
            End If
        End If
        Call SelectDescriptionForInvType("Sub_Type_Description", lblInvSubTypeDes, StrCheckHelp, "Sub_Type")
        txtInvSubType.Focus()
        Call txtInvSubType_Validating(txtInvSubType, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdLocationCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocationCode.Click
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Display Help from Location_Mst for Location_Code,Description
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        If Len(Me.txtLocationCode.Text) = 0 Then
            StrCheckHelp = ShowList(1, (txtLocationCode.MaxLength), "", "Location_Code", "Description", "Location_mst", " and Loc_Type='A'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtLocationCode.Text = StrCheckHelp
            End If
        Else
            StrCheckHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "Location_Code", "Description", "Location_mst", " and Loc_Type='A'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtLocationCode.Text = StrCheckHelp
            End If
        End If
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub CmdpurposeCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdpurposeCodeHelp.Click
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Display Help from Account_Mst for Account_Code,Description
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        If Len(Me.txtPurposeCode.Text) = 0 Then
            StrCheckHelp = ShowList(1, (txtPurposeCode.MaxLength), "", "cnttab_SubCodeId", "cnttab_Value1", "Gen_ControlTable", "and cnttab_MajorId = 'fin' AND cnttab_CodeId = 'prpsCode'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtPurposeCode.Text = StrCheckHelp
            End If
        Else
            StrCheckHelp = ShowList(1, (txtPurposeCode.MaxLength), "", "cnttab_SubCodeId", "cnttab_Value1", "Gen_ControlTable", "and cnttab_MajorId = 'fin' AND cnttab_CodeId = 'prpsCode'")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtPurposeCode.Text = StrCheckHelp
            End If
        End If
        '---------------
        'Procedure Call to Select Description  from Account_mst
        'Arguments  -   AccountCode
        '           -   Label To Set the Caption
        '---------------
        Call SelectDesOfAccountCode(StrCheckHelp, lblPurposeDes)
        txtRptName.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdStockLocation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdStockLocation.Click
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Display Help from Location_Mst for Location_Code,Description
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        If Len(Me.txtStockLocation.Text) = 0 Then
            StrCheckHelp = ShowList(1, (txtStockLocation.MaxLength), "", "Location_Code", "Description", "Location_mst")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtStockLocation.Text = StrCheckHelp
            End If
        Else
            StrCheckHelp = ShowList(1, (txtStockLocation.MaxLength), txtStockLocation.Text, "Location_Code", "Description", "Location_mst")
            If StrCheckHelp = "-1" Then
                Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Else
                txtStockLocation.Text = StrCheckHelp
            End If
        End If
        '---------------
        'Procedure Call to Select Description  from Location Master
        'Arguments  -   AccountCode
        '           -   Label To Set the Caption
        '---------------
        Call CheckLocationCode((txtStockLocation.Text), True)
        txtStockLocation.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CSubOne_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CSubOne.Click
        On Error GoTo ErrHandler
        If Val(txtCurrentNo.Text) > 0 Then
            txtCurrentNo.Text = Val(txtCurrentNo.Text) - 1
            If CheckDocNoFromSalesChallanDetail() Then
                MsgBox("An invoice exists with " & Val(txtCurrentNo.Text) + 1 & " document number." & vbCrLf & "Please skip either backward or forward.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTMST0005.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTMST0005_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex 'Check Form Name In MDI Menu
        If Me.txtLocationCode.Enabled = True Then CboFinancialYear.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0005_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0005_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0005_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If User Pressed the Escape Key then Display Form In View Mode
        '*****************************************************
        Dim rsSalesConf As ClsResultSetDB
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Trim(Me.CmdGrpSaleConf.lbMode) <> "View" Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        mstrCategory = lblCategory.Text
                        mstrLocation_Code = txtLocationCode.Text
                        mstrInvType = txtInvoiceType.Text
                        rsSalesConf = New ClsResultSetDB
                        Call ToGetYearStartandYearEnd()
                        rsSalesConf.GetResult("Select Description from SaleConf where Unit_Code='" & gstrUNITID & "' And Invoice_type ='" & Trim(mstrInvType) & "' And Category ='" & Trim(mstrCategory) & "' and Location_code ='" & Trim(mstrLocation_Code) & "' and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
                        If rsSalesConf.GetNoRows > 0 Then
                            lblInvTypeDes.Text = rsSalesConf.GetValue("Description")
                            mstrSubType = txtInvSubType.Text
                            Call Me.CmdGrpSaleConf.Revert()
                            Call EnableControls(False, Me, True)
                            CboFinancialYear.SelectedIndex = 0
                            lblCategory.Text = mstrCategory
                            txtLocationCode.Text = mstrLocation_Code
                            txtInvoiceType.Text = mstrInvType
                            txtInvSubType.Text = mstrSubType
                            Call txtInvSubType_Validating(txtInvSubType, New System.ComponentModel.CancelEventArgs(False))
                            CboFinancialYear.Enabled = True : CboFinancialYear.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtInvoiceType.Enabled = True : txtInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtInvSubType.Enabled = True : txtInvSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdLocationCode.Enabled = True
                            lblCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdInvoiceType.Enabled = True : CmdInvSubType.Enabled = True
                            Me.CmdGrpSaleConf.Enabled(0) = True
                            gblnCancelUnload = False
                            gblnFormAddEdit = False
                            Me.CmdGrpSaleConf.Focus()
                        End If
                        rsSalesConf.ResultSetClose()
                    Else
                        Me.ActiveControl.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTMST0005_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rstDB As ClsResultSetDB
        Dim strSql As String
        Dim strSQLA As String
        Dim rstDBA As ClsResultSetDB
        'Add Form Name To Window List
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FillLabelFromResFile(Me)
        Call FitToClient(Me, FraSC, ctlFormHeader1, CmdGrpSaleConf, 500)
        'Procedure Call to Set Help Picture at Commands Button
        Call SetHelpPictureAtCommandButton()
        'Disable all controls
        Call EnableControls(False, Me, True)
        Call AddlabelToGrid()
        CboFinancialYear.Enabled = True : CboFinancialYear.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtInvoiceType.Enabled = True : txtInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtInvSubType.Enabled = True : txtInvSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        'To Visible off on Form Load (Openning Balence)
        ctlOpBalence.Visible = False : lblOpBalence.Visible = False
        lblCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        CmdInvoiceType.Enabled = True : CmdInvSubType.Enabled = True
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocationCode.Enabled = True
        Me.CmdGrpSaleConf.Enabled(0) = False
        CmdLocationCode.Image = My.Resources.ico111.ToBitmap
        CmdStockLocation.Image = My.Resources.ico111.ToBitmap
        Call AddDataToFinancialYear()
        Call ToGetYearStartandYearEnd()
        strSql = "select isnull(eou_flag,0) as eou_flag from sales_parameter where Unit_Code='" & gstrUNITID & "'"
        rstDB = New ClsResultSetDB
        Call rstDB.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB.RowCount > 0 Then
            If rstDB.GetValue("eou_flag") = False Then
                strSQLA = "select isnull(eou_flag,0) as eou_flag1 from company_mst where Unit_Code='" & gstrUNITID & "'"
                rstDBA = New ClsResultSetDB
                Call rstDBA.GetResult(strSQLA, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rstDBA.RowCount > 0 Then
                    If rstDBA.GetValue("eou_flag1") = False Then
                        Me.lblBond17.Visible = False
                        Me.txtBond17.Visible = False
                    End If
                End If
                rstDBA.ResultSetClose()
            End If
        End If
        rstDB.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0005_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Trim(Me.CmdGrpSaleConf.lbMode) <> "View" Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save data before saving
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        Call CmdGrpSaleConf_ButtonClick(eventSender, New UCActXCtl.UCbtngrptwo.ButtonClickEventArgs(mlng_SAVEBEFOREEXIT))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    'Set Global VAriable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpSaleConf.Focus()
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0005_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()  'Assign form to nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtInvoiceType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoiceType.TextChanged
        '**************************************************************
        'Created By     :   Kapil
        'Description    :   if text in the Field is blank then Clear all the fields and Display
        '                   Form in View Mode
        '**************************************************************
        On Error GoTo ErrHandler
        If Len(txtInvoiceType.Text) = 0 Then
            Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                Case "View"
                    txtInvSubType.Text = "" : txtPurposeCode.Text = ""
                    txtRptName.Text = ""
                    txtStockLocation.Text = "" : ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Unchecked : ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    lblInvTypeDes.Text = "" : lblInvSubTypeDes.Text = ""
                    lblPurposeDes.Text = ""
                    txtCurrentNo.Text = ""
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtInvoiceType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '                   And Check Validation Of Invoice Type
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "View"
                        If Len(txtInvoiceType.Text) > 0 Then
                            Call txtInvoiceType_Validating(txtInvoiceType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtInvSubType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoiceType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If F1 Key Is Pressed then Display Help For Invoice Type From SaleConf
        '*****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdInvoiceType.Enabled Then
                Call CmdInvoiceType_Click(CmdInvoiceType, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtInvoiceType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoiceType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If User Enters the Invoice Type then Check the Validity
        '                   from SaleConfig and if text in the Field is Blank then
        '                   Restrict User to Enter the Invoice Type b'coz Invoice SubType
        '                   is based on the Invoice Type
        '*****************************************************
        On Error GoTo ErrHandler
        If Len(txtInvoiceType.Text) > 0 Then
            'Check Location Code whether entered or not
            If Len(txtLocationCode.Text) = 0 Then
                Call ConfirmWindow(10474, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtLocationCode.Focus()
                GoTo EventExitSub
            End If
            Call ToGetYearStartandYearEnd()
            If CheckValidDataFromDB("Invoice_Type", "SaleConf", Trim(txtInvoiceType.Text), lblInvTypeDes, "Description", " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'") Then
                If UCase(Trim(txtInvoiceType.Text)) <> "EXP" Then
                    lblOpBalence.Visible = True : ctlOpBalence.Visible = True
                Else
                    lblOpBalence.Visible = False : ctlOpBalence.Visible = False
                End If
                If Len(Trim(txtInvSubType.Text)) Then
                    Call txtInvSubType_Validating(txtInvSubType, New System.ComponentModel.CancelEventArgs(False))
                Else
                    txtInvSubType.Focus()
                End If
            Else
                Call ConfirmWindow(10364, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtInvoiceType.Text = ""
                lblInvTypeDes.Text = ""
                txtInvoiceType.Focus()
            End If
        Else
            CmdInvoiceType.Focus()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtInvSubType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvSubType.TextChanged
        '**************************************************************
        'Created By     :   Kapil
        'Description    :   if text in the Field is blank then Clear all the fields and Display
        '                   Form in View Mode
        '**************************************************************
        On Error GoTo ErrHandler
        If Len(txtInvSubType.Text) = 0 Then
            Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                Case "View"
                    txtPurposeCode.Text = ""
                    lblInvSubTypeDes.Text = ""
                    txtStockLocation.Text = ""
                    ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    lblPurposeDes.Text = ""
                    txtCurrentNo.Text = ""
                    txtRptName.Text = ""
                    ctlOpBalence.Text = "0.00"
                    Me.spdCostCenter.MaxRows = 0
                    Me.CmdGrpSaleConf.Enabled(0) = False
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtInvSubType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvSubType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "View"
                        If Len(txtInvSubType.Text) > 0 Then
                            Call txtInvSubType_Validating(txtInvSubType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.CmdGrpSaleConf.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvSubType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvSubType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If F1 Key Is Pressed then Display Help For Invoice Type From SaleConf
        '*****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdInvSubType.Enabled Then
                Call CmdInvSubType_Click(CmdInvSubType, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtInvSubType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvSubType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If User Enters the Invoice SubType then Check the Validity
        '                   from SaleConfig and if text in the Field is Blank then
        '                   Restrict User to Enter the Invoice Type b'coz Invoice SubType
        '                   is based on the Invoice Type
        '*****************************************************
        On Error GoTo ErrHandler
        If Len(txtInvSubType.Text) > 0 Then
            'Check Location Code whether entered or not
            If Len(txtLocationCode.Text) = 0 Then
                Call ConfirmWindow(10474, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtLocationCode.Focus()
                GoTo EventExitSub
            End If
            If Len(txtInvoiceType.Text) = 0 Then
                Call ConfirmWindow(10475, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtInvoiceType.Focus()
                GoTo EventExitSub
            End If
            Call ToGetYearStartandYearEnd()
            If CheckValidDataFromDB("Sub_Type", "SaleConf", Trim(txtInvSubType.Text), lblInvSubTypeDes, "Sub_Type_Description", " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'", txtInvoiceType.Text) Then
                If selectDataFromSaleConf() Then
                    Me.CmdGrpSaleConf.Enabled(0) = True
                    Me.CmdGrpSaleConf.Focus()
                    GoTo EventExitSub
                Else
                    Call ConfirmWindow(10367, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtInvSubType.Text = ""
                    lblInvSubTypeDes.Text = ""
                    txtInvSubType.Focus()
                End If
            Else
                Call ConfirmWindow(10366, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtInvSubType.Text = ""
                lblInvSubTypeDes.Text = ""
                txtInvSubType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '**************************************************************
        'Created By     :   Kapil
        'Description    :   if text in the Field is blank then Clear all the fields and Display
        '                   Form in View Mode
        '**************************************************************
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) = 0 Then
            Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                Case "View"
                    txtInvoiceType.Text = "" : txtInvSubType.Text = ""
                    txtPurposeCode.Text = "" : lblInvSubTypeDes.Text = ""
                    txtStockLocation.Text = ""
                    ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
                    lblPurposeDes.Text = ""
                    txtCurrentNo.Text = ""
                    Me.CmdGrpSaleConf.Enabled(0) = False
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '                   And Check Validation Of Invoice Type
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "View"
                        If Len(txtLocationCode.Text) > 0 Then
                            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtInvoiceType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLocationCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If F1 Key Is Pressed then Display Help For Accounting Location Code From Location Master
        '*****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocationCode.Enabled Then
                Call cmdLocationCode_Click(CmdLocationCode, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   Check Validation Of Location Code
        '*****************************************************
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If CheckLocationCode((txtLocationCode.Text)) Then
                txtInvoiceType.Focus()
            Else
                Cancel = True
                Call ConfirmWindow(10434, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtLocationCode.Text = ""
                txtLocationCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtPurposeCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPurposeCode.TextChanged
        ''*****************************************************
        ''Created By     :   Kapil
        ''Description    :   If Text in the field in blank then clear the Description
        ''*****************************************************
        On Error GoTo ErrHandler
        If Len(txtPurposeCode.Text) = 0 Then
            lblPurposeDes.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtPurposeCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPurposeCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        If Len(txtPurposeCode.Text) > 0 Then
                            Call txtPurposeCode_Validating(txtPurposeCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtStockLocation.Enabled Then txtStockLocation.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPurposeCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPurposeCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '**************************************
        'Created By     :   Kapil
        'Description    :   Display Help from Account_Mst for Account_Code,Description at F1 Key Press
        '**************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdpurposeCodeHelp.Enabled Then
                Call CmdpurposeCodeHelp_Click(CmdpurposeCodeHelp, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtPurposeCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtPurposeCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtPurposeCode.Text) > 0 Then
            If CheckAccountCode((txtPurposeCode.Text), txtPurposeCode, lblPurposeDes, 5) Then
                If txtStockLocation.Enabled Then txtStockLocation.Focus()
            Else
                Cancel = True
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtRptName_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRptName.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim strtemp As String
        strtemp = "<>?,./~`!@#$%^&*()+:'" & Chr(34) & Chr(32)
        If InStr(1, strtemp, Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                Case "Edit"
                    If txtBond17.Enabled Then
                        txtBond17.Focus()
                    Else
                        If ChkUpdateStockFlag.Enabled Then ChkUpdateStockFlag.Focus()
                    End If
            End Select
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtStockLocation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStockLocation.TextChanged
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   If Text in the field in blank then clear the Description
        '*****************************************************
        On Error GoTo ErrHandler
        If Len(txtStockLocation.Text) = 0 Then
            lblStockLocationDes.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtStockLocation_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtStockLocation.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        If Len(txtStockLocation.Text) > 0 Then
                            Call txtStockLocation_Validating(txtStockLocation, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtRptName.Enabled Then txtRptName.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ChkUpdateStockFlag_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles ChkUpdateStockFlag.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        'ChkUpdatePOFlag.SetFocus
                        System.Windows.Forms.SendKeys.Send(vbTab)
                    Case "View"
                End Select
                ChkUpdateStockFlag_Leave(ChkUpdateStockFlag, New System.EventArgs())
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ChkUpdatePOFlag_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles ChkUpdatePOFlag.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        If txtCurrentNo.Enabled Then
                            txtCurrentNo.Focus()
                        ElseIf ctlOpBalence.Visible And ctlOpBalence.Enabled = True Then
                            ctlOpBalence.Focus()
                        Else
                            Me.chkPre.Focus()
                        End If
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub SelectDesOfAccountCode(ByRef pAccCode As String, ByRef pLbl As System.Windows.Forms.Control)
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Select Description Corresponding to the Account Code from
        '                   Account_Mst
        'Arguments      :   pAccCode  -  Account Code,pLbl  -  Label to set Caption
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim strDesSql As String 'To Make Select Query
        Dim rsDesSql As ClsResultSetDB
        strDesSql = "SELECT     cnttab_SubCodeId,cnttab_Value1 "
        strDesSql = strDesSql & " From Gen_ControlTable "
        strDesSql = strDesSql & "where Unit_Code='" & gstrUNITID & "'  And cnttab_MajorId = 'fin' AND "
        strDesSql = strDesSql & " cnttab_CodeId = 'prpsCode'"
        strDesSql = strDesSql & " and cnttab_SubCodeId= '" & txtPurposeCode.Text & "'"
        rsDesSql = New ClsResultSetDB
        rsDesSql.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'If Record Found
        If rsDesSql.GetNoRows > 0 Then
            pLbl.Text = rsDesSql.GetValue("cnttab_Value1")
        End If
        rsDesSql.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SelectDescriptionForInvType(ByRef pstrDes As String, ByRef pActCtl As System.Windows.Forms.Control, ByRef pstrFieldText As String, ByRef pstrFieldText1 As String)
        '********************************************************
        'Created By     -   Kapil
        'Description    -   Select Description From SaleConf as per Field Name
        '********************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Make Select Query
        Dim rsFieldDes As ClsResultSetDB
        'Select Query
        Call ToGetYearStartandYearEnd()
        strDesSql = "Select " & Trim(pstrDes) & " from SaleConf Where Unit_Code='" & gstrUNITID & "' And "
        strDesSql = strDesSql & Trim(pstrFieldText1) & "='" & Trim(pstrFieldText) & "'"
        strDesSql = strDesSql & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
        strDesSql = strDesSql & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
        rsFieldDes = New ClsResultSetDB
        rsFieldDes.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsFieldDes.GetNoRows > 0 Then
            'If Record Found
            pActCtl.Text = rsFieldDes.GetValue(Trim(pstrDes))
        Else
        End If
        rsFieldDes.ResultSetClose()
        rsFieldDes = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function CheckValidDataFromDB(ByRef pstrFName As String, ByRef pstrTName As String, ByRef pstrFText As String, ByRef pActiveCtl As System.Windows.Forms.Control, ByRef pCtrlDesc As String, ByRef pstrCondition As String, Optional ByRef strInvType As String = "") As Boolean
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   Check Validity Of Field Data in The Table
        'Arguments      :   pstrFName - Field Name,pstrTName - Table Name,pstrFText - Field Data
        '*****************************************************
        On Error GoTo ErrHandler
        CheckValidDataFromDB = False
        Dim strFieldSql As String 'Make Select Query
        Dim rsFieldVal As ClsResultSetDB
        rsFieldVal = New ClsResultSetDB
        'Select Query
        If Trim(pstrFName) = "Sub_Type" Then
            strFieldSql = "Select * from " & Trim(pstrTName) & " Where Unit_Code='" & gstrUNITID & "' And Invoice_Type='" & Trim(strInvType) & "' and "
            strFieldSql = strFieldSql & Trim(pstrFName) & "='" & Trim(pstrFText) & "'" & pstrCondition
        Else
            strFieldSql = "Select * from " & Trim(pstrTName) & " Where Unit_Code='" & gstrUNITID & "' And "
            strFieldSql = strFieldSql & Trim(pstrFName) & "='" & Trim(pstrFText) & "'" & pstrCondition
        End If
        rsFieldVal.GetResult(strFieldSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsFieldVal.GetNoRows > 0 Then
            CheckValidDataFromDB = True
            pActiveCtl.Text = rsFieldVal.GetValue(Trim(pCtrlDesc))
        Else
            CheckValidDataFromDB = False
        End If
        rsFieldVal.ResultSetClose()
        rsFieldVal = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function selectDataFromSaleConf() As Boolean
        '********************************************************
        'Created By     -   Kapil
        'Description    -   Select Data from the SaleConf Acc. to the Invoice Type,
        '               -   Invoice SubType and Category
        '********************************************************
        On Error GoTo ErrHandler
        selectDataFromSaleConf = False
        Dim strSelectSql As String 'Make Select Query
        Dim rsGetExData As ClsResultSetDB
        Dim blnStockFlag As Boolean 'To Retreive Stock Flag
        Dim blnPOFlag As Boolean 'To Retreive P.O Flag
        Dim blnSameSeriesFlag As Boolean ' TO RET. SAME SERIES FLAG
        Dim blnPrePrintFlag As Boolean 'To Retreive P.O Flag
        'Make Select Query
        Call ToGetYearStartandYearEnd()
        strSelectSql = "Select Fin_start_date,Fin_end_date,Location_Code,Category,Category_Description,Invoice_Type,Description,Sub_Type,Sub_Type_Description,Sale_Account_Code,Excise_Account_Code,Insurance_Account_Code,SaleTax_Account_Code,CustomDuty_Account_Code,Frieght_Account_Code,Others_Account_code,Amortization_Account_code,Turnovertax_Account_code,Stock_Location,updateStock_Flag,updatePO_Flag,Excise_1,Excise_2,Excise_3,Suffix,Current_No,OpenningBal,Preprinted_flag,NoCopies,inv_GLd_prpsCode,Report_filename,Single_series,SecReport_filename,SuppReport_filename,Cust_code,bond17OpeningBal,ExciseFormatReport,Prev_Year_ExpInv_Sale,Permissible_Sale,RecordsPerPage,SORequired,MultipleSOAllowed,BarCodeTrackingAllowed,Report_FilenameII,BatchTrackingAllowed , isnull(CURRENT_NO_TRF_SAMEGSTIN,0) CURRENT_NO_TRF_SAMEGSTIN from SaleConf Where Unit_Code='" & gstrUNITID & "' And Category='S' and Invoice_Type='" & Trim(txtInvoiceType.Text) & "' and "
        strSelectSql = strSelectSql & " Sub_Type='" & Trim(txtInvSubType.Text) & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        strSelectSql = strSelectSql & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
        strSelectSql = strSelectSql & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
        rsGetExData = New ClsResultSetDB
        rsGetExData.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetExData.GetNoRows > 0 Then
            selectDataFromSaleConf = True
            txtLocationCode.Text = rsGetExData.GetValue("Location_Code")
            txtPurposeCode.Text = rsGetExData.GetValue("inv_GLd_prpsCode")
            Call SelectDesOfAccountCode((txtPurposeCode.Text), lblPurposeDes)
            txtRptName.Text = IIf(IsDBNull(rsGetExData.GetValue("Report_filename")), "", rsGetExData.GetValue("Report_filename"))
            TxtReportII.Text = IIf(IsDBNull(rsGetExData.GetValue("Report_filenameII")), "", rsGetExData.GetValue("Report_filenameII"))
            txtStockLocation.Text = rsGetExData.GetValue("Stock_Location")
            Call CheckLocationCode(Trim(txtStockLocation.Text), True)
            If UCase(txtInvoiceType.Text) = "TRF" And chkDeliverychallan.Checked = True Then
                txtCurrentNo.Text = rsGetExData.GetValue("CURRENT_NO_TRF_SAMEGSTIN")
            Else
                txtCurrentNo.Text = rsGetExData.GetValue("Current_No")
            End If

            txtBond17.Text = rsGetExData.GetValue("bond17OpeningBal")
            blnStockFlag = rsGetExData.GetValue("updateStock_Flag")
            If Not blnStockFlag Then
                ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                ChkUpdateStockFlag.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            blnPOFlag = rsGetExData.GetValue("updatePO_Flag")
            If Not blnPOFlag Then
                ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                ChkUpdatePOFlag.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            ' FOR SAME SERIES FLAG
            blnSameSeriesFlag = rsGetExData.GetValue("Single_Series")
            If Not blnSameSeriesFlag Then
                ChkSameSeries.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                ChkSameSeries.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            If UCase(Trim(txtInvoiceType.Text)) <> "EXP" Then
                ctlOpBalence.Visible = True : lblOpBalence.Visible = True
                ctlOpBalence.Text = rsGetExData.GetValue("OpenningBal")
            Else
                ctlOpBalence.Visible = False : lblOpBalence.Visible = False
            End If
            blnPrePrintFlag = rsGetExData.GetValue("PrePrinted_Flag")
            If blnPrePrintFlag = True Then
                chkPre.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkPre.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            ctlNoOfCopies.Text = rsGetExData.GetValue("noCopies")
            Call DisplayDetailsinGrid()
        Else
            selectDataFromSaleConf = False
        End If
        rsGetExData.ResultSetClose()
        rsGetExData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function ValidateBeforeSave() As Boolean
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Check the Blank Fields In The Form
        '*****************************************************
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim dblTotalAllocation As Double
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim varallocation As Object
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidateBeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        If Len(Me.txtPurposeCode.Text) = 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ".  Purpose  Code "
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtPurposeCode
            End If
            ValidateBeforeSave = False
        End If
        If Len(Me.txtRptName.Text) = 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Report File Name "
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtRptName
            End If
            ValidateBeforeSave = False
        End If
        If Len(Me.txtStockLocation.Text) = 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Stock Location"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtStockLocation
            End If
            ValidateBeforeSave = False
        End If
        If Len(txtCurrentNo.Text) = 0 Then
            txtCurrentNo.Text = "0"
        End If
        intMaxCounter = spdCostCenter.MaxRows
        For intLoopCounter = 1 To intMaxCounter
            '***To Check Blank  total of cost center distribution
            varallocation = Nothing
            Call spdCostCenter.GetText(3, intLoopCounter, varallocation)
            dblTotalAllocation = dblTotalAllocation + Val(varallocation)
        Next
        If Val(CStr(dblTotalAllocation)) <> 0 Then
            If Val(CStr(dblTotalAllocation)) <> 100 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Total of Cost Center Distribution must be 100 % "
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = spdCostCenter
                    With spdCostCenter
                        .Row = intLoopCounter
                        .Col = 3
                    End With
                End If
                ValidateBeforeSave = False
            End If
        End If
        If Not ValidateBeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            gblnCancelUnload = True
            If TypeOf lctrFocus Is System.Windows.Forms.TextBox Then
                lctrFocus.Focus()
            Else
                DirectCast(lctrFocus, AxFPSpreadADO.AxfpSpread).Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function CheckAccountCode(ByRef pstrText As String, ByRef pActiveCtrl As System.Windows.Forms.Control, ByRef pActiveLbl As System.Windows.Forms.Control, ByRef pintNatute As Short, Optional ByRef pstrOtherCon As String = "") As Boolean
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Check Accound Code In Account_Mst whether It Exists or Not
        'Arguments      :   pstrText  -  Account Code,aActiveCtrl  -  Control Name
        '--------------------------------------
        On Error GoTo ErrHandler
        CheckAccountCode = False
        Dim strAccSql As String
        Dim rsAccSql As ClsResultSetDB
        If Len(pstrText) > 0 Then
            strAccSql = "SELECT     cnttab_SubCodeId,cnttab_Value1 " & " From Gen_ControlTable " & " Where Unit_Code='" & gstrUNITID & "' And   cnttab_MajorId = 'fin' " & " AND cnttab_CodeId = 'prpsCode' " & "  and cnttab_SubCodeId ='" & Me.txtPurposeCode.Text & "'"
            rsAccSql = New ClsResultSetDB
            rsAccSql.GetResult(strAccSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsAccSql.GetNoRows > 0 Then
                CheckAccountCode = True
                pActiveLbl.Text = rsAccSql.GetValue("cnttab_Value1")
            Else
                CheckAccountCode = False
                Call ConfirmWindow(10247, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                lblPurposeDes.Text = ""
                pActiveCtrl.Text = ""
                pActiveCtrl.Focus()
            End If
            rsAccSql.ResultSetClose()
            rsAccSql = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub txtStockLocation_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtStockLocation.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdStockLocation.Enabled Then
                Call CmdStockLocation_Click(CmdStockLocation, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function CheckLocationCode(ByRef pstrLocCode As String, Optional ByRef blnflag As Boolean = False) As Boolean
        '************************************
        'Check Location Code In The Location Master
        '************************************
        On Error GoTo ErrHandler
        CheckLocationCode = False
        Dim strLocCode As String
        Dim rsLocCode As ClsResultSetDB
        strLocCode = "Select Location_Code,Description from Location_Mst Where Unit_Code='" & gstrUNITID & "' And Location_Code='" & Trim(pstrLocCode) & "'"
        rsLocCode = New ClsResultSetDB
        rsLocCode.GetResult(strLocCode, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsLocCode.GetNoRows > 0 Then
            CheckLocationCode = True
            If blnflag = True Then
                lblStockLocationDes.Text = rsLocCode.GetValue("Description")
            End If
        Else
            CheckLocationCode = False
        End If
        rsLocCode.ResultSetClose()
        rsLocCode = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub txtStockLocation_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtStockLocation.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtStockLocation.Text) > 0 Then
            If CheckLocationCode((txtStockLocation.Text), True) Then
                If txtRptName.Enabled Then txtRptName.Focus()
            Else
                Cancel = True
                Call ConfirmWindow(10402, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtStockLocation.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub SetHelpPictureAtCommandButton()
        '******************************************
        'To Set the Help Image at Command Button
        '******************************************
        On Error GoTo ErrHandler
        CmdpurposeCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdStockLocation.Image = My.Resources.ico111.ToBitmap
        CmdInvoiceType.Image = My.Resources.ico111.ToBitmap
        CmdInvSubType.Image = My.Resources.ico111.ToBitmap
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function CheckMaxDocNoFromSalesChallanDetail() As Boolean
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Select Max(Document No.)From Sales Challan Details Table
        '               :   b'coz Invoice Upto that Number has been made so that
        '               :   Current Number Can't be less then that Document Number
        '--------------------------------------
        CheckMaxDocNoFromSalesChallanDetail = False
        Dim strDocNoSql As String 'Declared To Make Select Query
        Dim strFinStartDate As String
        Dim strFinEndDate As String
        Dim dblsuffix As Double
        Dim rsDocNo As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        strFinancialYearStart = Trim(Mid(CboFinancialYear.Text, 1, InStr(1, CboFinancialYear.Text, "-") - 1))
        strFinancialYearEnd = Trim(Mid(CboFinancialYear.Text, InStr(1, CboFinancialYear.Text, "-") + 1))
        strDocNoSql = "Select max(Doc_No) as Doc_No from SalesChallan_Dtl Where Unit_Code='" & gstrUNITID & "' And doc_no < 99000000"
        strDocNoSql = strDocNoSql & " and Invoice_Type='" & Trim(txtInvoiceType.Text) & "'"
        If ChkSameSeries.CheckState <> 1 Then
            strDocNoSql = strDocNoSql & " and sub_category='" & Trim(txtInvSubType.Text) & "'"
        End If
        strDocNoSql = strDocNoSql & " and datediff(dd,invoice_date,'" & getDateForDB(VB6.Format(strFinancialYearStart, gstrDateFormat)) & "')<=0  and datediff(dd,'" & getDateForDB(VB6.Format(strFinancialYearEnd, gstrDateFormat)) & "',invoice_date)<=0"
        rsDocNo = New ClsResultSetDB
        rsDocNo.GetResult(strDocNoSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDocNo.GetNoRows > 0 Then
            mstrDocumentno = rsDocNo.GetValue("Doc_No")
            rsSaleConf = New ClsResultSetDB
            rsSaleConf.GetResult("select suffix from saleconf Where Unit_Code='" & gstrUNITID & "' And convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "' and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'")
            dblsuffix = rsSaleConf.GetValue("suffix")
            rsSaleConf.ResultSetClose()
            If dblsuffix > 0 Then
                If Len(Trim(CStr(dblsuffix))) > 0 Then
                    mstrDocumentno = Mid(mstrDocumentno, Len(Trim(CStr(dblsuffix))) + 1)
                End If
            End If
            If Val(txtCurrentNo.Text) < Val(mstrDocumentno) Then
                CheckMaxDocNoFromSalesChallanDetail = True
            Else
                CheckMaxDocNoFromSalesChallanDetail = False
            End If
        End If
        rsDocNo.ResultSetClose()
        rsDocNo = Nothing
    End Function
    Public Sub AddlabelToGrid()
        On Error GoTo ErrHandler
        spdCostCenter.MaxRows = 0
        spdCostCenter.MaxCols = 3
        With spdCostCenter
            .PrintRowHeaders = False
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .set_RowHeight(.Row, 300)
            '.RowHeight(0) = 13
            .Row = 0
            .Col = 1 : .Text = "Cost Center Code"
            '''.ColWidth(1) = 1500
            .set_ColWidth(1, 1500)
            .Row = 0
            .Col = 2 : .Text = "Description"
            .set_ColWidth(2, 3500)
            .Row = 0
            .Col = 3 : .Text = "Allocation % "
            .set_ColWidth(3, 1800)
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
            .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsVertical
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Sub AddNewRowType()
        On Error GoTo ErrHandler
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        With spdCostCenter
            .Row = .MaxRows
            .Col = 1
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeMaxEditLen = 13
            .set_ColWidth(1, 1500)
            .Row = .MaxRows
            .Col = 2
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeMaxEditLen = 60
            .set_ColWidth(2, 3500)
            .Row = .MaxRows
            .Col = 3
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMin = 0
            .TypeFloatDecimalPlaces = 0
            .TypeFloatMax = 100
            .set_ColWidth(3, 1800)
        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub addNewInSpread()
        '****************************************************
        'Created By     -  rishi
        'Description    -  Add Row At Enter Key Press Of Last Column Of Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        Dim varCurrency As Object
        With Me.spdCostCenter
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .set_RowHeight(.Row, 300)
            If .MaxRows > 6 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Call AddNewRowType()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Function DisplayDetailsinGrid() As Object
        On Error GoTo ErrHandler
        Dim rsCustItem As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim strItemcode As String
        Dim NoOfRed As Short
        Dim strSql As String
        rsCustItem = New ClsResultSetDB
        rsCustItem.GetResult("Select Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage from invcc_dtl Where Unit_Code='" & gstrUNITID & "' And location_code ='" & txtLocationCode.Text & "' and  Invoice_Type ='" & txtInvoiceType.Text & "' and Sub_Type ='" & txtInvSubType.Text & "'  ORDER BY ccM_ccCode")
        If rsCustItem.GetNoRows > 0 Then
            blnWhetherEnteredOrNOt = True
        Else
            ' cost center allocation not done
            blnWhetherEnteredOrNOt = False
            MsgBox("Cost Center Allocation For this  location code & invoice type and subtype not  available " & vbCrLf & " Please Define Cost Center Allocation", MsgBoxStyle.Information, "Empower")
            If CmdGrpSaleConf.Mode = "E" Then
                With spdCostCenter
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 3
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End With
            End If
        End If
        rsCustItem.ResultSetClose()
        strSql = "SELECT b.ccM_ccCode,b.ccM_ccDesc,a.ccM_cc_percentage  FROM (" & " Select UNIT_CODE,Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage from invcc_dtl Where Unit_Code='" & gstrUNITID & "' And location_code ='" & txtLocationCode.Text & "' and  Invoice_Type ='" & txtInvoiceType.Text & "' and Sub_Type ='" & txtInvSubType.Text & "') a" & " full OUTER JOIN FIN_CCMASTER b ON a.ccM_ccCode=b.ccM_ccCode And a.Unit_Code=B.Unit_code and b.ccM_transTag='1' where a.UNIT_CODE='" & gstrUNITID & "' and b.UNIT_CODE='" & gstrUNITID & "'"
        rsCustItem = New ClsResultSetDB
        rsCustItem.GetResult(strSql)
        If rsCustItem.GetNoRows > 0 Then
            ' cost center entries done
            intMaxCounter = rsCustItem.GetNoRows
            rsCustItem.MoveFirst()
            With spdCostCenter
                For intLoopCounter = 1 To intMaxCounter
                    If spdCostCenter.MaxRows < intLoopCounter Then
                        Call addNewInSpread()
                    End If
                    Call spdCostCenter.SetText(1, intLoopCounter, rsCustItem.GetValue("ccM_ccCode"))
                    Call spdCostCenter.SetText(3, intLoopCounter, rsCustItem.GetValue("ccM_cc_percentage"))
                    strItemcode = rsCustItem.GetValue("ccM_ccCode")
                    rsItemMst = New ClsResultSetDB
                    rsItemMst.GetResult("Select ccM_ccDesc from fin_ccMaster Where Unit_Code='" & gstrUNITID & "' And ccM_ccCode ='" & strItemcode & "'")
                    Call spdCostCenter.SetText(2, intLoopCounter, rsItemMst.GetValue("ccM_ccDesc"))
                    rsItemMst.ResultSetClose()
                    rsCustItem.MoveNext()
                Next
                .Enabled = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = 3
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        End If
        rsCustItem.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Sub InsertData(ByRef intRow As Short)
        On Error GoTo ErrHandler
        Dim varlocCode As Object
        Dim varinvoicetype As String
        Dim varinvoicesubtype As String
        Dim varcostcentercode As Object
        Dim varcenterallocation As Object
        varlocCode = Trim(txtLocationCode.Text)
        varinvoicetype = Trim(txtInvoiceType.Text)
        varinvoicesubtype = Trim(txtInvSubType.Text)
        With spdCostCenter
            varcostcentercode = Nothing
            Call .GetText(1, intRow, varcostcentercode)
            varcenterallocation = Nothing
            Call .GetText(3, intRow, varcenterallocation)
        End With
        If Val(varcenterallocation) = 0 Then
            varcenterallocation = "0.0"
        End If
        strInsert = strInsert & vbCrLf & "Insert into invcc_dtl(Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage, Ent_Dt, Ent_UserId,upd_Dt,Upd_userId,Unit_Code) Values ( '" & varlocCode & "','" & varinvoicetype & "','" & varinvoicesubtype & "','" & varcostcentercode & "','" & varcenterallocation & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "')"
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function lockSameSeriesFlag() As Boolean
        Dim rsbillflag As ClsResultSetDB
        Dim strFinStartDate As String
        Dim strFinEndDate As String
        strFinancialYearStart = Trim(Mid(CboFinancialYear.Text, 1, InStr(1, CboFinancialYear.Text, "-") - 1))
        strFinancialYearEnd = Trim(Mid(CboFinancialYear.Text, InStr(1, CboFinancialYear.Text, "-") + 1))
        rsbillflag = New ClsResultSetDB
        Dim strSql As String
        strSql = "SELECT Location_Code,Doc_No,Suffix,Transport_Type,Vehicle_No,From_Station,To_Station,Invoice_Date,Account_Code,Cust_Ref,Amendment_No,Bill_Flag,Print_DateTime,Form3,Form3Date,Carriage_Name,Year,Insurance,Frieght_Tax,Invoice_Type,Ref_Doc_No,Cust_Name,Sales_Tax_Amount,Surcharge_Sales_Tax_Amount,Frieght_Amount,Packing_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,SalesTax_FormValue,Annex_no,Currency_Code,Nature_of_Contract,OriginStatus,Ctry_Destination_Goods,Delivery_Terms,Payment_Terms,Pre_Carriage_By,Receipt_Precarriage_at,Vessel_Flight_number,Port_Of_Loading,Port_Of_Discharge,Final_destination,Mode_Of_Shipment,Dispatch_mode,Buyer_description_Of_Goods,Invoice_description_of_EPC,Exchange_Date,Buyer_Id,Exchange_Rate,total_quantity,total_amount,TurnoverTax_per,Turnover_amt,other_ref,FIFO_flag,Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,Ent_dt,Ent_UserId,Upd_dt,Upd_Userid,Print_Flag,Cancel_flag,pervalue,remarks,dataPosted,ftp,Excise_Type,SRVDINO,SRVLocation,ExciseExumpted,LoadingChargeTaxType,LoadingChargeTaxAmount,LoadingChargeTax_Per,ConsigneeContactPerson,ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3,ConsigneeECCNo,ConsigneeLST,ServiceInvoiceformatExport,CustBankID,Discount_Type,Discount_Amount,Discount_Per,RejectionPosting,USLOC,SchTime,TCSTax_Type,TCSTax_Per,TCSTaxAmount,To_Location,ECESS_Type,ECESS_Per,ECESS_Amount,FOC_Invoice,From_Location,PrintExciseFormat,FreshCrRecd,SRCESS_Type,SRCESS_Per,SRCESS_Amount,CVDCESS_Type,CVDCESS_Per,CVDCESS_Amount,Excise_Percentage,invoice_time,Permissible_Limit,TurnOverTaxType,TotalInvoiceAmtRoundOff_diff,NRGPNOIncaseOfServiceInvoice,Trans_Parameter_Flag,SDTax_Type,SDTax_Per,SDTax_Amount,InvoiceAgainstMultipleSO,TextFileGenerated,sameunitloading,ServiceTax_Type,ServiceTax_Per,ServiceTax_Amount,Prev_Yr_ExportSales,Permissible_Limit_SmpExport,varGeneralRemarks,SECESS_Type,SECESS_Per,SECESS_Amount,CVDSECESS_Type,CVDSECESS_Per,CVDSECESS_Amount,SRSECESS_Type,SRSECESS_Per,SRSECESS_Amount,postingFlag,CheckSheetNo,MULTIPLESO,ISCHALLAN,ISCONSOLIDATE,Tot_Add_Excise_Amt,Tot_Add_Excise_PER,CONSIGNEE_CODE,Lorry_No,OTL_No,RefChallan,price_bases,LorryNo_date,dataposted_fin,ConsInvString,bond17OpeningBal,barCodeImage,invoicepicking_status,Ecess_TotalDuty_Type,Ecess_TotalDuty_Per,Ecess_TotalDuty_Amount,SEcess_TotalDuty_Type,SEcess_TotalDuty_Per,SEcess_TotalDuty_Amount,ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount From SalesChallan_Dtl  Where Unit_Code='" & gstrUNITID & "' And  (Location_Code = '" & txtLocationCode.Text & "') AND (Bill_Flag = 1) "
        strSql = strSql & " and datediff(dd,invoice_date,'" & getDateForDB(VB6.Format(strFinancialYearStart, gstrDateFormat)) & "')<=0  and datediff(dd,'" & getDateForDB(VB6.Format(strFinancialYearEnd, gstrDateFormat)) & "',invoice_date)<=0"
        rsbillflag.GetResult(strSql)
        If rsbillflag.GetNoRows > 0 Then
            lockSameSeriesFlag = True
        Else
            lockSameSeriesFlag = False
        End If
        rsbillflag.ResultSetClose()
    End Function
    Public Sub AddDataToFinancialYear()
        Dim strYear As String
        Dim intLoopCounter As Short
        Dim intmaxLoop As Short
        Dim rsSaleConf As ClsResultSetDB
        rsSaleConf = New ClsResultSetDB
        strYear = "select distinct Convert(char(10),Fin_start_date,103) + ' - ' + convert(char(10),Fin_end_date,103) AS FinancialYear from saleConf Where Unit_Code='" & gstrUNITID & "' order by FinancialYear desc"
        rsSaleConf.GetResult(strYear)
        intmaxLoop = rsSaleConf.GetNoRows
        rsSaleConf.MoveFirst()
        CboFinancialYear.Items.Clear()
        For intLoopCounter = 1 To intmaxLoop
            CboFinancialYear.Items.Add(rsSaleConf.GetValue("FinancialYear"))
            rsSaleConf.MoveNext()
        Next
        rsSaleConf.ResultSetClose()
        CboFinancialYear.SelectedIndex = 0
    End Sub
    Public Sub ToGetYearStartandYearEnd()
        Dim strString As String
        strFinancialYearStart = ""
        strFinancialYearEnd = ""
        strString = CboFinancialYear.Text
        strFinancialYearStart = Mid(strString, 1, InStr(1, strString, "-") - 2)
        strFinancialYearEnd = Mid(strString, InStr(1, strString, "-") + 2)
    End Sub
    Private Function CheckDocNoFromSalesChallanDetail() As Boolean
        '--------------------------------------
        'Created By     :   Arshad Ali
        'Description    :   Select Document No. From Sales Challan Details Table
        '--------------------------------------
        CheckDocNoFromSalesChallanDetail = False
        Dim strDocNoSql As String 'Declared To Make Select Query
        Dim strFinStartDate As String
        Dim strFinEndDate As String
        Dim rsDocNo As ClsResultSetDB
        strFinancialYearStart = Trim(Mid(CboFinancialYear.Text, 1, InStr(1, CboFinancialYear.Text, "-") - 1))
        strFinancialYearEnd = Trim(Mid(CboFinancialYear.Text, InStr(1, CboFinancialYear.Text, "-") + 1))
        strDocNoSql = "Select Doc_No as Doc_No from SalesChallan_Dtl Where Unit_Code='" & gstrUNITID & "' And doc_no < 99000000"
        strDocNoSql = strDocNoSql & " and Invoice_Type='" & Trim(txtInvoiceType.Text) & "'"
        If ChkSameSeries.CheckState <> 1 Then
            strDocNoSql = strDocNoSql & " and sub_category='" & Trim(txtInvSubType.Text) & "'"
        End If
        strDocNoSql = strDocNoSql & " and right(doc_no,6)=" & Val(Trim(txtCurrentNo.Text)) + 1
        strDocNoSql = strDocNoSql & " and datediff(dd,invoice_date,'" & getDateForDB(VB6.Format(strFinancialYearStart, gstrDateFormat)) & "')<=0  and datediff(dd,'" & getDateForDB(VB6.Format(strFinancialYearEnd, gstrDateFormat)) & "',invoice_date)<=0"
        rsDocNo = New ClsResultSetDB
        rsDocNo.GetResult(strDocNoSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDocNo.GetNoRows > 0 Then
            CheckDocNoFromSalesChallanDetail = True
        Else
            CheckDocNoFromSalesChallanDetail = False
        End If
        rsDocNo.ResultSetClose()
        rsDocNo = Nothing
    End Function
    Private Sub CmdGrpSaleConf_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtngrptwo.ButtonClickEventArgs) Handles CmdGrpSaleConf.ButtonClick
        On Error GoTo ErrHandler
        Dim rsCompany As ClsResultSetDB
        Dim strUpdateCurrentNo As String
        Dim strUpdateNoOfCopies As String
        Dim intmaxrows As Short
        Dim intLoopCounter As Short
        rsCompany = New ClsResultSetDB
        Dim strUpdateSql As String 'Declared to make Update Query
        Select Case Me.CmdGrpSaleConf.Mode
            Case "E"
                If e.ControlIndex = mlng_SAVEBEFOREEXIT Then GoTo SaveBeforeExit
                rsCompany.GetResult("Select EOU_Flag from Company_Mst Where Unit_Code='" & gstrUNITID & "'")
                EOU_Flag = rsCompany.GetValue("EOU_Flag")
                rsCompany.ResultSetClose()
                Call EnableControls(True, Me)
                If EOU_Flag = False Then
                    ctlOpBalence.Enabled = False
                    ctlOpBalence.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If
                If EOU_Flag And UCase(Trim(txtInvoiceType.Text)) = "EXP" And Val(txtBond17.Text) = 0 Then
                    txtBond17.Enabled = True
                    txtBond17.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Else
                    txtBond17.Enabled = False
                    txtBond17.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If
                CboFinancialYear.Enabled = False : CboFinancialYear.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblCategory.Enabled = False : lblCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtInvoiceType.Enabled = False : txtInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtInvSubType.Enabled = False : txtInvSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtLocationCode.Enabled = False : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdInvoiceType.Enabled = False : CmdInvSubType.Enabled = False : CmdLocationCode.Enabled = False
                If Val(txtCurrentNo.Text) = 0 Then
                    txtCurrentNo.Enabled = True
                    txtCurrentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Else
                    txtCurrentNo.Enabled = False
                    txtCurrentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdAddOne.Enabled = True : CSubOne.Enabled = True
                End If
                If lockSameSeriesFlag() = True Then
                    ChkSameSeries.Enabled = False
                End If

                If UCase(txtInvoiceType.Text) = "TRF" Then
                    chkDeliverychallan.Enabled = True
                Else
                    chkDeliverychallan.Enabled = False
                End If

                txtPurposeCode.Focus()
                With spdCostCenter
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 3
                    .Col2 = 3
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End With
            Case "S"
SaveBeforeExit:
                '***Check Blank Field In The Form
                If Not ValidateBeforeSave() Then
                    Exit Sub
                End If
                If CheckDocNoFromSalesChallanDetail() Then
                    MsgBox("An invoice exists with " & Val(txtCurrentNo.Text) + 1 & " document number." & vbCrLf & "Please skip either backward or forward and then press Update.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                    Exit Sub
                End If
                Call ToGetYearStartandYearEnd()
                strUpdateSql = "Update SaleConf set inv_GLd_prpsCode='" & Trim(txtPurposeCode.Text) & "',"
                strUpdateSql = strUpdateSql & "Report_filename='" & Trim(txtRptName.Text) & "',"
                strUpdateSql = strUpdateSql & "Stock_Location='" & Trim(txtStockLocation.Text) & "',"
                strUpdateSql = strUpdateSql & "updateStock_Flag=" & IIf(ChkUpdateStockFlag.CheckState, 1, 0) & ","
                strUpdateSql = strUpdateSql & "updatePO_Flag=" & IIf(ChkUpdatePOFlag.CheckState, 1, 0) & ","
                strUpdateSql = strUpdateSql & "Report_FilenameII='" & Trim(TxtReportII.Text) & "',"
                strUpdateSql = strUpdateSql & "bond17OpeningBal ='" & Trim(txtBond17.Text) & "',"
                strUpdateSql = strUpdateSql & " Single_Series=" & IIf(ChkSameSeries.CheckState, 1, 0) & ","
                strUpdateSql = strUpdateSql & "Upd_dt=getdate(),Upd_UserId='" & Trim(mP_User) & "'"
                strUpdateSql = strUpdateSql & " Where Unit_Code='" & gstrUNITID & "' And Category='" & Trim(lblCategory.Text) & "' and "
                strUpdateSql = strUpdateSql & " Invoice_Type='" & Trim(txtInvoiceType.Text) & "' and "
                strUpdateSql = strUpdateSql & " Sub_Type='" & Trim(txtInvSubType.Text) & "' and "
                strUpdateSql = strUpdateSql & "Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strUpdateSql = strUpdateSql & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
                strUpdateSql = strUpdateSql & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
                strUpdateCurrentNo = ""
                If EOU_Flag = True Then
                    If UCase(txtInvoiceType.Text) = "EXP" Then
                        Call ToGetYearStartandYearEnd()
                        strUpdateCurrentNo = "Update saleConf set Current_No=" & Val(txtCurrentNo.Text) & " ,OpenningBal =" & Val(ctlOpBalence.Text)
                        strUpdateCurrentNo = strUpdateCurrentNo & " Where Unit_Code='" & gstrUNITID & "' And "
                        strUpdateCurrentNo = strUpdateCurrentNo & "Location_Code='" & Trim(txtLocationCode.Text) & "'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and Invoice_Type='EXP'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
                    Else
                        Call ToGetYearStartandYearEnd()
                        If UCase(txtInvoiceType.Text) = "TRF" And chkDeliverychallan.Checked = True Then
                            strUpdateCurrentNo = "Update saleConf set CURRENT_NO_TRF_SAMEGSTIN=" & Val(txtCurrentNo.Text) & " ,OpenningBal =" & CDbl(ctlOpBalence.Text)
                        Else
                            strUpdateCurrentNo = "Update saleConf set Current_No=" & Val(txtCurrentNo.Text) & " ,OpenningBal =" & CDbl(ctlOpBalence.Text)
                        End If
                        strUpdateCurrentNo = strUpdateCurrentNo & " Where Unit_Code='" & gstrUNITID & "' And "
                        strUpdateCurrentNo = strUpdateCurrentNo & "Location_Code='" & Trim(txtLocationCode.Text) & "'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and Invoice_Type<>'EXP'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
                        strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
                    End If
                Else
                    Call ToGetYearStartandYearEnd()
                    '**** To update some dat on the basis of invoice type
                    If UCase(txtInvoiceType.Text) = "TRF" And chkDeliverychallan.Checked = True Then
                        strUpdateCurrentNo = "Update saleConf set CURRENT_NO_TRF_SAMEGSTIN=" & Val(txtCurrentNo.Text) & " ,OpenningBal =" & Val(ctlOpBalence.Text)
                    Else
                        strUpdateCurrentNo = "Update saleConf set Current_No=" & Val(txtCurrentNo.Text) & " ,OpenningBal =" & Val(ctlOpBalence.Text)
                    End If

                    strUpdateCurrentNo = strUpdateCurrentNo & " Where Unit_Code='" & gstrUNITID & "' And Category='" & Trim(lblCategory.Text) & "' and "
                    If Me.ChkSameSeries.CheckState = 1 Then
                        strUpdateCurrentNo = strUpdateCurrentNo & " Single_series = 1 and "
                    Else
                        strUpdateCurrentNo = strUpdateCurrentNo & " Invoice_Type='" & Trim(txtInvoiceType.Text) & "' and "
                        strUpdateCurrentNo = strUpdateCurrentNo & " Single_series = 0 and "
                    End If
                    strUpdateCurrentNo = strUpdateCurrentNo & "Location_Code='" & Trim(txtLocationCode.Text) & "'"
                    strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
                    strUpdateCurrentNo = strUpdateCurrentNo & " and convert(char(10),Fin_end_date,103) = '" & strFinancialYearEnd & "'"
                End If
                '*To Update No of Copies & PrePrinted invoice by nisha on 10/07/02
                strUpdateNoOfCopies = ""
                Call ToGetYearStartandYearEnd()
                strUpdateNoOfCopies = "Update saleConf set PrePrinted_flag=" & IIf(chkPre.CheckState, 1, 0) & ", NoCopies = " & ctlNoOfCopies.Text
                strUpdateNoOfCopies = strUpdateNoOfCopies & " Where Unit_Code='" & gstrUNITID & "' And Category='" & Trim(lblCategory.Text) & "' and "
                strUpdateNoOfCopies = strUpdateNoOfCopies & " Invoice_Type='" & Trim(txtInvoiceType.Text) & "' and "
                strUpdateNoOfCopies = strUpdateNoOfCopies & "Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strUpdateNoOfCopies = strUpdateNoOfCopies & " and convert(char(10),Fin_start_date,103) = '" & strFinancialYearStart & "'"
                strUpdateNoOfCopies = strUpdateNoOfCopies & " and convert(char(10),Fin_end_date,103)  = '" & strFinancialYearEnd & "'"
                intmaxrows = spdCostCenter.MaxRows
                strInsert = ""
                For intLoopCounter = 1 To intmaxrows
                    Call InsertData(intLoopCounter)
                Next
                With mP_Connection
                    .BeginTrans()
                    .Execute("Set DateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strUpdateSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strUpdateCurrentNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strUpdateNoOfCopies, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ' cost center details
                    .Execute("DELETE from invcc_dtl Where Unit_Code='" & gstrUNITID & "' And location_code ='" & txtLocationCode.Text & "' and  Invoice_Type ='" & txtInvoiceType.Text & "' and Sub_Type ='" & txtInvSubType.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Trim(strInsert)) > 0 Then .Execute(strInsert, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .CommitTrans()
                End With
                gblnCancelUnload = False
                gblnFormAddEdit = False
                Me.CmdGrpSaleConf.Revert()
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Call EnableControls(False, Me)
                txtInvoiceType.Enabled = True : txtInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtInvSubType.Enabled = True : txtInvSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocationCode.Enabled = True
                CboFinancialYear.Enabled = True : CboFinancialYear.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdInvoiceType.Enabled = True : CmdInvSubType.Enabled = True
                With spdCostCenter
                    .Enabled = True
                    .BlockMode = True
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 3
                    .Lock = True
                    .BlockMode = False
                End With
                txtLocationCode.Focus()
            Case ""
                'If Cancel Button is Pressed
                Call frmMKTMST0005_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case "X"
                If e.ControlIndex = mlng_SAVEBEFOREEXIT Then GoTo SaveBeforeExit
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlNoOfCopies_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlNoOfCopies.KeyPress
        If e.KeyAscii = System.Windows.Forms.Keys.Return Then
            If ChkSameSeries.Enabled = True Then
                ChkSameSeries.Focus()
            Else
                spdCostCenter.Focus()
                spdCostCenter.Col = 3
                spdCostCenter.Row = 1
                spdCostCenter.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdCostCenter.Focus()
            End If
        End If
    End Sub
    Private Sub ctlNoOfCopies_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlNoOfCopies.Leave
        '****by nisha on 10/07/02
        If Me.ctlNoOfCopies.Enabled = True Then
            If Val(ctlNoOfCopies.Text) > 10 Then
                MsgBox("Can not Print more then 10 copies", MsgBoxStyle.Information, "empower")
                ctlNoOfCopies.Focus()
            ElseIf Val(ctlNoOfCopies.Text) < 1 Then
                MsgBox("Enter atleast 1 copy", MsgBoxStyle.Information, "empower")
                ctlNoOfCopies.Focus()
            End If
        End If
    End Sub
    Private Sub ctlOpBalence_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlOpBalence.KeyPress
        '*****************************************************
        'Created By     :   Kapil
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        ' chkPre.SetFocus
                        System.Windows.Forms.SendKeys.Send(vbTab)
                End Select
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtBond17_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtBond17.KeyPress
        '*****************************************************
        'Created By     :   Arshad Ali
        'Description    :   At Enter Key Press Set Focus to Next Control
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Trim(Me.CmdGrpSaleConf.lbMode)
                    Case "Edit"
                        ' chkPre.SetFocus
                        If ChkUpdateStockFlag.Enabled Then ChkUpdateStockFlag.Focus()
                End Select
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCurrentNo_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtCurrentNo.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Call txtCurrentNo_Validating(txtCurrentNo, New System.ComponentModel.CancelEventArgs(False))
                If chkPre.Enabled = True Then
                    chkPre.Focus()
                End If
        End Select
    End Sub
    Private Sub txtCurrentNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCurrentNo.Validating
        Dim Cancel As Boolean = e.Cancel
        If Val(txtCurrentNo.Text) > 9000000 Then
            Cancel = True
            Call MsgBox("No. Can Not Be Greater Then 9000000", MsgBoxStyle.Information, "empower")
            txtCurrentNo.Text = "9000000"
        End If
        e.Cancel = Cancel
    End Sub
 
    Private Sub chkDeliverychallan_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDeliverychallan.CheckedChanged
        selectDataFromSaleConf()
    End Sub
End Class