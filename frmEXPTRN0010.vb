Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmEXPTRN0010
	Inherits System.Windows.Forms.Form
	'===================================================================================
    '(c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    'File Name         :   FRMEXPTRN0010.frm
    'Function          :   Used to add sale details
    'Created By        :   Nisha & Kapil
    'Created On        :   15 May, 2001
    'Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 3
	'03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 7
	'09/10/2001  changed on version 8 for schedule Status
	'09/01/2002 changed fof Smiel Chennei to add CVD_PER,SVD_Per,Insurance
	'25/01/2002 changed for decimal 4 places on Chacked Out Form No = 4019
	'28/01/2002 changed for decimal 4 places on Chacked Out Form No = 4033
	'in ChangeCellTypeStaticText()
	'02/02/2002 Add Export Challan Entry
	'15/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4068
	'22/03/2002 INCREASED SIZE OF CONTAINER NO.
	'27/06/2002 DatePicker Added ,So that Dates of Export Invoice can be set.   -
	'                    - NITIN SOOD
	'changed by nisha on 21/03/2003 for financial rollover & temp no will start from 99000001
	'changed by nisha on 01/05/2003 for back date entry check
	'Changed by nisha 0n 25/07/203
	'Three new feilds added 1.Service type invoice Check Box
	'                        2.Bank id
	'                        3.Remarks
	'Changes done by Sourabh khatri on 19 jan 2004 for Add New Invoice sub Type (Sample Invoice)
	'Changes done by Jogender on 09 jun 2006 for (Sample Invoice)--Tax related saving against issue ID 17857
	'===================================================================================
	'Revision  By       : Ashutosh , Issue Id :18903
	'Revision On        : 02-11-2006
	'History            : Save Stock Location in Invoice (in Saleschallan_dtl.from_location).
	'                   : Stock check was not correct while saving invoice.
	'---------------------------------------------------------------------------------------
	'Revised By    :  Davinder Singh
	'Revision Date :  01 Dec 2006
	'Issue ID      :  19165
	'Purpose       :  To rondoff the saved data in SalesChallan_Dtl and Sales_dtl tables
	'                 according to parameters defined in sales_parameter
	'---------------------------------------------------------------------------------------
	'Revision  By       : Ashutosh , Issue Id :19385
	'Revision On        : 29-01-2007
	'History            : Wrong Schedule Checking in Export Documentation Invoice.(WMART  Kandla)
	'---------------------------------------------------------------------------------------
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 19992
	'Revision Date   : 29 June 2007
	'History         : Display Credit Term from Cust_Ord_Dtl and save into saleschallan_dtl
	'                  During Invoice Posting, fetch credit term from saleschallan_dtl for saving in ar_docmaster
    '-----------------------------------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    06/05/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------------------------------
    'Revised By      : Prashant Rajpal
    'Issue ID        : 10179325
    'Revision Date   : 19 Jan 2012
    '***************************************************************************************
    'Modified By Deepak on 31 Jan 2012 to support multiunit change management
    '**********************************************************************************************************************
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    19/03/2015
    'REVISION HISTORY  -    CHANGES FOR LOGIC CHANGE FOR GENERATING TEMPORARY INVOICE: DUE TO DOUBLE ENTRY DATA
    'ISSUE ID           -     10777177  
    '**********************************************************************************************************************

	Dim mintIndex As Short 'Declared To Hold The Form Count
	Dim mdblPrevQty() As Object 'to store prev quantity in edit mode
	Dim mdblToolCost() As Object 'to insert tool cost item wise
	Dim ArrExpDetails() As String
	Dim strExpDetails As String
	Dim strExpEditDetails As String
	Public mstrItemCode As String 'To Get The Value Of Item Code
	Dim mstrInvoiceType As String 'To Get The Value Of Invoice Type
	Dim mstrInvoiceSubType As String 'To Get The Value Of Invoice Sub Type
	Dim mstrAmmendmentNo As String 'To Get The Value Of Ammendment No.
	Dim mstrInvType As String 'To Get Value Of Inv Type From SalesChallan_Dtl
	Dim mstrInvSubType As String 'To Get Value Of Inv SubType From SalesChallan_Dtl
	Dim mstrUpdDispatchSql As String 'To Make Update Query For Dispatch_Qty From Daily/Monthly Mkt Schedule
	Dim mstrAmmNo As String
	Dim mstrRefNo As String
	Dim strupSalechallan As String
	Dim strupSaleDtl As String
	Dim strInvType As String
    Dim strInvSubType As String
    Public strRefAmm As String
    Dim mstrexportsotype As String = String.Empty
    Dim mblnWithPayExpInvoice As Boolean = False
    Private Sub chkServiceInvFormat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkServiceInvFormat.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtBankAc.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmbInvSubType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.SelectedIndexChanged
        On Error GoTo ErrHandler
        Call SelectInvTypeSubTypeFromSaleConf((CmbInvType.Text), (CmbInvSubType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmbInvSubType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvSubType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtLocationCode.Focus()
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
    Private Sub CmbInvSubType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.Leave
        '*******************************************************************************
        'Author             :   Sourabh
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh invoice Sub type
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If UCase(Trim(Me.CmbInvSubType.Text)) = "SAMPLE" Then
                Me.txtRefNo.Enabled = False : Me.txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.CmdRefNoHelp.Enabled = False : Me.CmdRefNoHelp.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Else
                Me.txtRefNo.Enabled = True : Me.txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.CmdRefNoHelp.Enabled = True : Me.CmdRefNoHelp.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub


    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.SelectedIndexChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh invoice Sub type
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmbInvType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        CmbInvSubType.Focus()
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
    Private Sub CmbInvType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.Leave
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To set controls enabled/disabled Condition According to
        '                       Invoice type Selected
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Select Case UCase(CmbInvType.Text)
                    Case "EXPORT INVOICE"
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                        txtAnnex.Enabled = False : txtAnnex.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdAnnexHelp.Enabled = False
                        txtExciseDuty.Enabled = False
                        txtExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtAddExciseDuty.Enabled = False
                        txtAddExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        ctlSVD.Enabled = False
                        ctlSVD.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSalesTax.Enabled = False
                        Me.ctlInsurance.Enabled = False
                        Me.ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        Me.txtFreight.Enabled = False
                        Me.txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        Me.txtSurcharge.Enabled = False
                        Me.txtSurcharge.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSalesTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        CmdSaleTaxType.Enabled = False
                End Select
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmbTransType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbTransType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtVehNo.Focus()
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

    Private Sub cmdAcCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAcCode.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From SaleTax Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strBankNo() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelpString = "Select Bnk_BankId,Bnk_accNo from Gen_bankMaster where unit_code='" & gstrUNITID & "' "
                strBankNo = ctlExportChallanEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelpString, "Bank Codes")
                If UBound(strBankNo) < 0 Then Exit Sub
                If strBankNo(0) = "0" Then
                    MsgBox("No Bank Code Available To Display.", MsgBoxStyle.Information, ResolveResString(100)) : txtBankAc.Text = "" : txtBankAc.Focus() : Exit Sub
                Else
                    If strBankNo(0).Trim <> "" Then
                        txtBankAc.Text = strBankNo(0)
                        lblAcCodeDes.Text = strBankNo(1)
                    End If
                End If
        End Select
        txtBankAc.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help For Invoice No.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Trim(txtLocationCode.Text) = "" Then
                    Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    txtLocationCode.Focus()
                    Exit Sub
                End If
                If Len(Trim(txtChallanNo.Text)) = 0 Then
                    strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", "Invoice_Date", "SalesChallan_Dtl ", "and Invoice_Type ='EXP' AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                Else
                    strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", "Invoice_Date", "SalesChallan_Dtl ", "and Invoice_Type ='EXP' AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                End If
        End Select
        txtChallanNo.Focus()
        If Val(txtChallanNo.Text) > 99000000 Then
            Cmditems.Enabled = True
        Else
            Cmditems.Enabled = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Customer code's Help
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Trim(mstrInvoiceType) = "EXP" Then
                    If Len(Trim(txtCustCode.Text)) = 0 Then
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", , "Customer Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                        End If
                    Else
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), txtCustCode.Text, "customer_code", "cust_name", "Customer_Mst", , "Customer Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                        End If
                    End If
                End If
        End Select
        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
        txtCustCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdGrpChEnt_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        'Function           :   Code for ADD/EDIT/UPDATE/CANCEL/CLOSE
        'Purpose       :  To rondoff the saved data in SalesChallan_Dtl and Sales_dtl tables
        '                 according to parameters defined in sales_parameter by Executing
        '                 stored procedure INVOICE_ROUNDOFF
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSalesChallan As String
        Dim strSalesDtl As String
        Dim Description As String
        Dim intLoopcount As Short
        Dim varQuantity As Object
        Dim varDrgNo As Object
        Dim varItemCode As Object
        Dim varRate As Object
        Dim varCustMtrl As Object
        Dim varPacking As Object
        Dim varOthers As Object
        Dim varFromBox As Object
        Dim VarToBox As Object
        Dim PresQty As Object
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim intLoop As Short
        Dim updatearr() As String
        Dim strsql As String
        Dim strMakeDate As String
        Dim strexportsotype As String
        Dim VarHSN As Object
        Dim VarCGSTType As Object
        Dim VarSGSTType As Object
        Dim VarUGSTType As Object
        Dim VarIGSTType As Object
        Dim VarCompensationtax As Object
        Dim intGSTTAXroundoff_decimal As Short
        Dim blnGSTTAXroundoff As Boolean
        Dim rsParameterData As ClsResultSetDB

        Dim strParamQuery = "select GSTTAX_ROUNDOFF,GSTTAX_ROUNDOFF_DECIMAL FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnGSTTAXroundoff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTTAXroundoff_decimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing


        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                strValues = ""
                strExpDetails = ""
                Call EnableControls(True, Me, True)
                Call SelectChallanNoFromSalesChallanDtl()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdChallanNo.Enabled = False : txtChallanNo.Enabled = False
                txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                lblLocCodeDes.Text = "" : lblCustCodeDes.Text = ""
                SpChEntry.Enabled = True
                CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0 : CmbTransType.SelectedIndex = 0
                With SpChEntry
                    .MaxRows = 1
                    .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .Lock = False : .BlockMode = False
                End With
                If Not (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    txtRefNo.Enabled = False
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdRefNoHelp.Enabled = False
                Else
                    txtRefNo.Enabled = True
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdRefNoHelp.Enabled = True
                End If
                CmbInvType.Visible = True : CmbInvSubType.Visible = True
                lblInvSubType.Visible = True : lblInvType.Visible = True
                lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat) ' FormatDateTime(GetServerDate(), DateFormat.GeneralDate)
                With dtpDateDesc
                    .Value = ConvertToDate(lblDateDes.Text)
                    .Visible = True
                End With
                Call SetMaxLengthInSpread()
                Call ChangeCellTypeStaticText()
                dtpDateDesc.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(False, Me)
                rsSalesChallandtl = New ClsResultSetDB
                rsSalesChallandtl.GetResult("SELECT Invoice_type from Saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_no = " & txtChallanNo.Text, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                rsSalesChallandtl.ResultSetClose()
                rsSalesChallandtl = Nothing

                txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                Me.txtRefNo.Enabled = False
                Me.txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Call SetMaxLengthInSpread()
                Call ChangeCellTypeStaticText()
                ReDim mdblPrevQty(SpChEntry.MaxRows - 1) ' To get value of Quantity in Arrey for updation in despatch
                For intLoop = 1 To SpChEntry.MaxRows
                    mdblPrevQty(intLoop - 1) = Nothing
                    Call SpChEntry.GetText(5, intLoop, mdblPrevQty(intLoop - 1))
                Next
                chkServiceInvFormat.Enabled = True : txtBankAc.Enabled = True : txtBankAc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdAcCode.Enabled = True
                txtRemarks.Enabled = True : txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmexport.Enabled = True : cmexport.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Not ValidatebeforeSave("ADD") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        lblDateDes.Text = VB6.Format(dtpDateDesc.Value, gstrDateFormat)
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        rsSaleConf = New ClsResultSetDB
                        rsSaleConf.GetResult("Select Invoice_Type,Sub_Type,Stock_Location from SaleConf where unit_code='" & gstrUNITID & "' and Description ='" & CmbInvType.Text & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        'strexportsotype = ""
                        'If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                        '    strexportsotype = Find_Value("Select exportsotype from cust_ord_hdr where UNIT_CODE='" + gstrUNITID + "' AND  account_code='" & Trim(txtCustCode.Text) & "' and cust_ref='" & txtRefNo.Text.Trim & "' and amendment_no='" & mstrAmmNo & "' and active_flag='a'")
                        'End If
                        strSalesChallan = ""
                        strSalesChallan = Trim(strSalesChallan) & "Insert into SalesChallan_dtl (UNIT_CODE,Location_Code,Doc_No,From_Location,Suffix,"
                        strSalesChallan = strSalesChallan & "Transport_Type,Vehicle_No,"
                        strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,Cust_Ref,"
                        strSalesChallan = strSalesChallan & "Amendment_No,Bill_Flag,Form3,Carriage_Name,"
                        strSalesChallan = strSalesChallan & "Year,Insurance,"
                        strSalesChallan = strSalesChallan & "Frieght_Tax,invoice_Type,Ref_Doc_No,Cust_Name,"
                        strSalesChallan = strSalesChallan & "Sub_Category,"
                        strSalesChallan = strSalesChallan & "Annex_no,invoice_Date,Ent_dt,"
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                            strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,print_flag,total_amount,ServiceInvoiceformatExport,CustBankID,Remarks, PrintExciseFormat, FreshCrRecd,ExportSotype ) Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                        Else
                            strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,print_flag,total_amount,ServiceInvoiceformatExport,CustBankID,Remarks, PrintExciseFormat, FreshCrRecd ) Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                        End If

                        strSalesChallan = strSalesChallan & "', " & Trim(txtChallanNo.Text) & ",'" & rsSaleConf.GetValue("Stock_Location") & "',''"
                        strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & (Trim(txtVehNo.Text)) & "' ,'"
                        strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                        strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "',0"
                        strSalesChallan = strSalesChallan & ",'','" & QuoteString(Trim(txtCarrServices.Text))
                        strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                        strSalesChallan = strSalesChallan & "IsNull(" & Val(ctlInsurance.Text) & ", 0)"
                        strSalesChallan = strSalesChallan & ",IsNull(" & Val(txtFreight.Text) & ", 0),'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                        strSalesChallan = strSalesChallan & Trim(txtAnnex.Text) & "','" & Trim(lblCustCodeDes.Text) & "',"
                        strSalesChallan = strSalesChallan & "'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "',"
                        strSalesChallan = strSalesChallan & "0,'" & getDateForDB(lblDateDes.Text) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',0," & CalculateTotalInvoiceAmount(0)
                        If chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked Then
                            strSalesChallan = strSalesChallan & ",1"
                        Else
                            strSalesChallan = strSalesChallan & ",0"
                        End If
                        rsSaleConf.ResultSetClose()
                        rsSaleConf = Nothing
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                            strSalesChallan = strSalesChallan & ",'" & Trim(txtBankAc.Text) & "','" & QuoteString(Trim(txtRemarks.Text)) & "', 0 ,0,'" & mstrexportsotype & "')"
                        Else
                            strSalesChallan = strSalesChallan & ",'" & Trim(txtBankAc.Text) & "','" & QuoteString(Trim(txtRemarks.Text)) & "', 0 ,0 )"
                        End If

                        If Len(Trim(strValues)) = 0 Then
                            MsgBox("Please Select Export Details", MsgBoxStyle.Information, ResolveResString(100))
                            cmexport.Focus()
                            Exit Sub
                        ElseIf Len(Trim(strValues)) = 36 Then
                            MsgBox("Please Select Export Details", MsgBoxStyle.Information, ResolveResString(100))
                            cmexport.Focus()
                            Exit Sub
                        Else
                            updatearr = Split(strValues, "§")
                            strsql = strsql & "update saleschallan_dtl set "
                            strsql = strsql & "Frieght_Amount=" & updatearr(16) & ","
                            strsql = strsql & "Currency_Code='" & updatearr(0) & "' ,"
                            strsql = strsql & "Nature_of_Contract='" & updatearr(7) & "' ,"
                            strsql = strsql & "OriginStatus='" & updatearr(1) & "' ,"
                            strsql = strsql & "Ctry_destination_goods='" & updatearr(2) & "' ,"
                            strsql = strsql & "Delivery_Terms='" & updatearr(11) & "' ,"
                            strsql = strsql & "Payment_Terms='" & Trim(lblCreditTerm.Text) & "' ,"
                            strsql = strsql & "Pre_carriage_by='" & updatearr(3) & "' ,"
                            strsql = strsql & "Receipt_Precarriage_at='" & updatearr(4) & "' ,"
                            strsql = strsql & "Vessel_flight_number= '' ,"
                            strsql = strsql & "Port_of_loading='" & updatearr(5) & "' ,"
                            strsql = strsql & "Port_of_discharge='" & updatearr(6) & "' ,"
                            strsql = strsql & "Final_destination='" & updatearr(8) & "' ,"
                            strsql = strsql & "Mode_of_Shipment='" & updatearr(9) & "' ,"
                            strsql = strsql & "DISPATCH_MODE='" & updatearr(10) & "' ,"
                            strsql = strsql & "Buyer_Description_of_Goods='" & updatearr(13) & "' ,"
                            strsql = strsql & "Invoice_Description_of_EPC='" & updatearr(14) & "' ,"
                            strsql = strsql & "Exchange_Rate='" & updatearr(15) & "',"
                            strsql = strsql & "Exchange_Date='" & getDateForDB(VB6.Format(updatearr(17), gstrDateFormat)) & "',"
                            strsql = strsql & "Other_ref ='" & QuoteString(updatearr(18)) & "',"
                            strsql = strsql & "buyer_id ='" & QuoteString(updatearr(19)) & "',"
                            strsql = strsql & "Prev_Yr_ExportSales =" & updatearr(20) & ","
                            strsql = strsql & "Permissible_Limit_SmpExport =" & updatearr(21) & ","
                            strsql = strsql & "varGeneralRemarks ='" & QuoteString(updatearr(22)) & "'"
                            strsql = strsql & ",total_amount =" & CalculateTotalInvoiceAmount(CDbl(updatearr(16)))
                        End If
                        strsql = strsql & " where unit_code='" & gstrUNITID & "' and "
                        strsql = strsql & " doc_no=" & txtChallanNo.Text & " and "
                        strsql = strsql & " suffix='' "
                        strSalesDtl = ""
                        For intLoopcount = 1 To SpChEntry.MaxRows
                            varItemCode = Nothing
                            varDrgNo = Nothing
                            varRate = Nothing
                            varCustMtrl = Nothing
                            varQuantity = Nothing
                            varPacking = Nothing
                            varOthers = Nothing
                            varFromBox = Nothing
                            VarToBox = Nothing
                            VarHSN = Nothing
                            VarCGSTType = Nothing
                            VarSGSTType = Nothing
                            VarUGSTType = Nothing
                            VarIGSTType = Nothing
                            VarCompensationtax = Nothing
                            Call SpChEntry.GetText(1, intLoopcount, varItemCode)
                            Call SpChEntry.GetText(2, intLoopcount, varDrgNo)
                            Call SpChEntry.GetText(3, intLoopcount, varRate)
                            Call SpChEntry.GetText(4, intLoopcount, varCustMtrl)
                            Call SpChEntry.GetText(5, intLoopcount, varQuantity)
                            Call SpChEntry.GetText(6, intLoopcount, varPacking)
                            Call SpChEntry.GetText(7, intLoopcount, varOthers)
                            Call SpChEntry.GetText(8, intLoopcount, varFromBox)
                            Call SpChEntry.GetText(9, intLoopcount, VarToBox)
                            Call SpChEntry.GetText(10, intLoopcount, VarHSN)
                            Call SpChEntry.GetText(11, intLoopcount, VarCGSTType)
                            Call SpChEntry.GetText(12, intLoopcount, VarSGSTType)
                            Call SpChEntry.GetText(13, intLoopcount, VarUGSTType)
                            Call SpChEntry.GetText(14, intLoopcount, VarIGSTType)
                            Call SpChEntry.GetText(15, intLoopcount, VarCompensationtax)
                            rsCustItemMst = New ClsResultSetDB
                            rsItemMst = New ClsResultSetDB
                            rsItemMst.GetResult("Select Description from Item_Mst where unit_code='" & gstrUNITID & "' and Item_Code ='" & Trim(varItemCode) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

                            rsCustItemMst.GetResult("Select Drg_desc from CustItem_Mst where unit_code='" & gstrUNITID & "' and Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & varDrgNo & "'and Item_code ='" & varItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            strSalesDtl = Trim(strSalesDtl) & "Insert into sales_Dtl(unit_code,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Packing,Others,Cust_Mtrl,Cust_Ref,Amendment_no,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,"
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Basic_Amount,Accessible_amount "
                            If gblnGSTUnit = True Then
                                strSalesDtl = strSalesDtl & " , HSNSACCODE,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT "
                            End If
                            strSalesDtl = strSalesDtl & " ) values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "',"

                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & ",'','" & Trim(varItemCode) & "'," & Val(varQuantity) & ","
                            strSalesDtl = strSalesDtl & Val(varFromBox) & "," & Val(VarToBox) & "," & Val(varRate) & ","
                            strSalesDtl = strSalesDtl & Val(varPacking) & "," & Val(varOthers) & "," & Val(varCustMtrl) & ",'" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "',"
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & ",'" & Trim(varDrgNo) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("dRG_DESC") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            rsItemMst.ResultSetClose()
                            rsItemMst = Nothing
                            rsCustItemMst.ResultSetClose()
                            rsCustItemMst = Nothing

                            If CmbInvType.Text = "NORMAL INVOICE" Then
                                strSalesDtl = strSalesDtl & mdblToolCost(intLoopcount - 1) & ",'',getdate(),'"
                            Else
                                strSalesDtl = strSalesDtl & "0,'',getdate(),'"
                            End If
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "'," & (Val(varRate) * Val(varQuantity)) & "," & (Val(varRate) * Val(varQuantity)) & ""
                            'GST CHANGES
                            If gblnGSTUnit = True Then
                                If blnGSTTAXroundoff = True Then
                                    strSalesDtl = strSalesDtl & ",'" & VarHSN & "','" & VarCGSTType & "','" & GetTaxRate(VarCGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "CGST", False)) & ",'" & VarSGSTType & "','" & GetTaxRate(VarSGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "SGST", False)) & ",'" & VarUGSTType & "','" & GetTaxRate(VarUGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='UTGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "UTGST", False)) & ",'" & VarIGSTType & "','" & GetTaxRate(VarIGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='IGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "IGST", False)) & ",'" & VarCompensationtax & "','" & GetTaxRate(VarCompensationtax, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='COMPENSATION CESS'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "GSTCC", False)) & ")"
                                Else
                                    strSalesDtl = strSalesDtl & ",'" & VarHSN & "','" & VarCGSTType & "','" & GetTaxRate(VarCGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "CGST", False), intGSTTAXroundoff_decimal) & ",'" & VarSGSTType & "','" & GetTaxRate(VarSGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "SGST", False), intGSTTAXroundoff_decimal) & ",'" & VarUGSTType & "','" & GetTaxRate(VarUGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='UTGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "UTGST", False), intGSTTAXroundoff_decimal) & ",'" & VarIGSTType & "','" & GetTaxRate(VarIGSTType, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='IGST'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "IGST", False), intGSTTAXroundoff_decimal) & ",'" & VarCompensationtax & "','" & GetTaxRate(VarCompensationtax, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='COMPENSATION CESS'") & "'," & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "GSTCC", False), intGSTTAXroundoff_decimal) & ")"
                                End If

                            Else
                                strSalesDtl = strSalesDtl & ")" & vbCrLf
                            End If
                            'GST CHANGES
                        Next

                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Not ValidatebeforeSave("EDIT") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        If Len(Trim(strValues)) = 0 Then
                            strValues = strExpDetails
                        End If
                        updatearr = Split(strValues, "§")
                        strSalesChallan = ""
                        strSalesChallan = "Update SalesChallan_Dtl Set "
                        strSalesChallan = strSalesChallan & "Insurance = " & Val(ctlInsurance.Text)
                        strSalesChallan = strSalesChallan & Val(txtSurcharge.Text) & ",Frieght_Tax=" & Val(txtFreight.Text)
                        strSalesChallan = strSalesChallan & ",Frieght_Amount=" & updatearr(16) & ","
                        strSalesChallan = strSalesChallan & "Currency_Code='" & updatearr(0) & "' ,"
                        strSalesChallan = strSalesChallan & "Nature_of_Contract='" & updatearr(7) & "' ,"
                        strSalesChallan = strSalesChallan & "OriginStatus='" & updatearr(1) & "' ,"
                        strSalesChallan = strSalesChallan & "Ctry_destination_goods='" & updatearr(2) & "' ,"
                        strSalesChallan = strSalesChallan & "Delivery_Terms='" & updatearr(11) & "' ,"
                        strSalesChallan = strSalesChallan & "Payment_Terms='" & Trim(lblCreditTerm.Text) & "' ,"
                        strSalesChallan = strSalesChallan & "Pre_carriage_by='" & updatearr(3) & "' ,"
                        strSalesChallan = strSalesChallan & "Receipt_Precarriage_at='" & updatearr(4) & "' ,"
                        strSalesChallan = strSalesChallan & "Vessel_flight_number= '' ,"
                        strSalesChallan = strSalesChallan & "Port_of_loading='" & updatearr(5) & "' ,"
                        strSalesChallan = strSalesChallan & "Port_of_discharge='" & updatearr(6) & "' ,"
                        strSalesChallan = strSalesChallan & "Final_destination='" & updatearr(8) & "' ,"
                        strSalesChallan = strSalesChallan & "Mode_of_Shipment='" & updatearr(9) & "' ,"
                        strSalesChallan = strSalesChallan & "DISPATCH_MODE='" & updatearr(10) & "' ,"
                        strSalesChallan = strSalesChallan & "Buyer_Description_of_Goods='" & updatearr(13) & "' ,"
                        strSalesChallan = strSalesChallan & "Invoice_Description_of_EPC='" & updatearr(14) & "' ,"
                        strSalesChallan = strSalesChallan & "Exchange_Date='" & getDateForDB(VB6.Format(updatearr(17), gstrDateFormat)) & "' ,"
                        strSalesChallan = strSalesChallan & "Exchange_Rate='" & Val(updatearr(15)) & "', "
                        strSalesChallan = strSalesChallan & "other_ref ='" & QuoteString(updatearr(18)) & "', "
                        strSalesChallan = strSalesChallan & "buyer_id ='" & QuoteString(updatearr(19)) & "', "
                        strSalesChallan = strSalesChallan & "Prev_Yr_ExportSales =" & updatearr(20) & ", "
                        strSalesChallan = strSalesChallan & "Permissible_Limit_SmpExport =" & updatearr(21) & ", "
                        strSalesChallan = strSalesChallan & "varGeneralRemarks ='" & QuoteString(updatearr(22)) & "' "
                        strSalesChallan = strSalesChallan & ",total_amount =" & CalculateTotalInvoiceAmount(CDbl(updatearr(16)))
                        If chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked Then
                            strSalesChallan = strSalesChallan & ",ServiceInvoiceformatExport = 1"
                        Else
                            strSalesChallan = strSalesChallan & ",ServiceInvoiceformatExport = 0"
                        End If
                        strSalesChallan = strSalesChallan & ",CustBankID = '" & Trim(txtBankAc.Text) & "'"
                        strSalesChallan = strSalesChallan & ",Remarks = '" & QuoteString(Trim(txtRemarks.Text)) & "'"
                    
                        strSalesChallan = strSalesChallan & " where unit_code='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                        strSalesChallan = strSalesChallan & " and Doc_No =" & Val(txtChallanNo.Text)

                        strSalesDtl = ""
                        For intLoopcount = 1 To SpChEntry.MaxRows
                            varQuantity = Nothing
                            varDrgNo = Nothing
                            varRate = Nothing
                            Call SpChEntry.GetText(5, intLoopcount, varQuantity)
                            Call SpChEntry.GetText(2, intLoopcount, varDrgNo)
                            Call SpChEntry.GetText(3, intLoopcount, varRate)

                            strSalesDtl = Trim(strSalesDtl) & "Update Sales_dtl set Sales_Quantity = " & Val(varQuantity) & ","
                            strSalesDtl = Trim(strSalesDtl) & "basic_amount=" & (Val(varQuantity) * Val(varRate)) & ","
                            strSalesDtl = Trim(strSalesDtl) & "Accessible_amount=" & (Val(varQuantity) * Val(varRate))
                            'GST CHANGES
                            If blnGSTTAXroundoff = True Then
                                strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "CGST", False)
                                strSalesDtl = Trim(strSalesDtl) & ",SGST_AMT=" & CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "SGST", False)
                                strSalesDtl = Trim(strSalesDtl) & ",UTGST_AMT=" & CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "UTGST", False)
                                strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "IGST", False)
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "CGST", False), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",SGST_AMT=" & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "SGST", False), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",UTGST_AMT=" & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "UTGST", False), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & System.Math.Round(CalculateGSTtaxes(intLoopcount, (Val(varRate) * Val(varQuantity)), "IGST", False), intGSTTAXroundoff_decimal)
                            End If

                            'GST CHANGES
                            strSalesDtl = Trim(strSalesDtl) & " where unit_code='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtl = Trim(strSalesDtl) & Trim(varDrgNo) & "'" & vbCrLf
                        Next
                End Select
                With mP_Connection
                    ResetDatabaseConnection()
                    .BeginTrans()
                    mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strSalesChallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        .Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If Len(Trim(mstrUpdDispatchSql)) > 0 And UCase(Trim(Me.CmbInvSubType.Text)) <> "SAMPLE" Then
                        .Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If RoundInvTables(CInt(txtChallanNo.Text), Trim(txtLocationCode.Text)) = False Then
                        .RollbackTrans()
                        Exit Sub
                    End If
                    .CommitTrans()
                End With
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Me.CmdGrpChEnt.Revert()
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                CmbInvType.Enabled = True : CmbInvSubType.Enabled = True
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
                With dtpDateDesc
                    lblDateDes.Text = VB6.Format(.Value, gstrDateFormat)
                    .Visible = False 'HIDE DTP - NITIN SOOD
                End With
                txtLocationCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmEXPTRN0010_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mstrUpdDispatchSql = ""
                    If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                        strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                    Else
                        strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                    End If
                    For intLoopcount = 1 To SpChEntry.MaxRows
                        varDrgNo = Nothing
                        PresQty = Nothing
                        Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                        Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) -  " & Val(PresQty)
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where unit_code='" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Status =1 " & vbCrLf

                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Update MonthlyMktSchedule set Despatch_qty ="
                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - " & Val(PresQty)
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'  and Status =1 " & vbCrLf
                    Next
                    Call DeleteRecords()
                    mP_Connection.BeginTrans()
                    mP_Connection.Execute(strupSaleDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If UCase(Trim(Me.CmbInvSubType.Text)) <> "SAMPLE" Then
                        mP_Connection.Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    mP_Connection.CommitTrans()
                    Call EnableControls(False, Me, True)
                    txtLocationCode.Enabled = True
                    txtLocationCode.BackColor = System.Drawing.Color.White
                    CmdLocCodeHelp.Enabled = True
                    txtChallanNo.Enabled = True
                    txtChallanNo.BackColor = System.Drawing.Color.White
                    CmdChallanNo.Enabled = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
        On Error GoTo ErrHandler
        GetTaxRate = 0
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then

            GetTaxRate = rsExistData.GetValue(Trim(pstrFieldName_WhichValueRequire))
        Else
            GetTaxRate = 0
        End If
        rsExistData.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function CalculateGSTtaxes(ByVal pintRowNo As Short, ByVal pdblAccessibleValue As Double, ByVal pTaxType As String, ByRef pblnEOU_FLAG As Boolean) As Double

        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsGetTaxRate As ClsResultSetDB
        Dim ldblTaxRate As Double
        Dim ldblTempTotalgsttax As Double

        On Error GoTo ErrHandler

        ldblTempTotalgsttax = 0
        Dim strCompCode As String
        Dim strInvoiceType As String
        Dim strsql As String

        If mblnWithPayExpInvoice = True Then
            With SpChEntry

                .Row = pintRowNo
                If UCase(pTaxType) = "CGST" Then
                    .Col = 11
                ElseIf UCase(pTaxType) = "SGST" Then
                    .Col = 12
                ElseIf UCase(pTaxType) = "UTGST" Then
                    .Col = 13
                ElseIf UCase(pTaxType) = "IGST" Then
                    .Col = 14
                Else
                    .Col = 15
                End If

                rsGetTaxRate = New ClsResultSetDB
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "'"
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblTaxRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblTaxRate = 0
                End If
                rsGetTaxRate.ResultSetClose()
                CalculateGSTtaxes = (pdblAccessibleValue * ldblTaxRate) / 100
            End With
        Else
            CalculateGSTtaxes = 0
        End If

        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Cmditems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmditems.Click
        'Function           :   Display Another Form for User To Select Item Code >From CustOrd_Dtl
        '                       And After Selecting Item Code Select Data From Sales_Dtl and Display
        '                       That Details In The Spread
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim rsSaleConf As ClsResultSetDB
        Dim rsSalesParameter As ClsResultSetDB
        Dim strSalesParameter As String
        Dim blnFGSRM As Boolean
        Dim strStockLocation As String
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        Select Case Me.CmdGrpChEnt.Mode

            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                rssalechallan = New ClsResultSetDB
                salechallan = ""
                salechallan = "Select Invoice_type,SUB_CATEGORY from saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_No = "
                salechallan = salechallan & Val(txtChallanNo.Text)
                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalechallan.GetNoRows > 0 Then
                    rssalechallan.MoveFirst()
                    strInvType = rssalechallan.GetValue("Invoice_type")
                    strInvSubType = rssalechallan.GetValue("sub_category")
                End If
                rssalechallan.ResultSetClose()
                rssalechallan = Nothing

                If (strInvType = "EXP") Then
                    strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE", " datediff(dd,'" & getDateForDB(lblDateDes.Text) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(lblDateDes.Text) & "')<=0")
                    If Len(Trim(strStockLocation)) > 0 Then
                        mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf")
                        Exit Sub
                    End If
                Else

                    mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") And (Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE") Then
                    If Len(Trim(txtRefNo.Text)) = 0 Then
                        Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtRefNo.Focus()
                        Exit Sub
                    End If
                End If
                rsSalesParameter = New ClsResultSetDB
                strSalesParameter = "Select bitSampleExportInvoiceforFGSRM from sales_parameter where unit_code='" & gstrUNITID & "'"
                rsSalesParameter.GetResult(strSalesParameter, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsSalesParameter.GetNoRows > 0 Then
                    rsSalesParameter.MoveFirst()
                    blnFGSRM = rsSalesParameter.GetValue("bitSampleExportInvoiceforFGSRM")
                End If

                rsSalesParameter.ResultSetClose()
                rsSalesParameter = Nothing

                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION", " datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0")
                    If Len(Trim(strStockLocation)) > 0 Then
                        If (Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE") Then
                            mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                        Else
                            If blnFGSRM Then
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst("SAMPLE INVOICE", "RAW MATERIAL & FINISHED GOODS", strStockLocation, Trim(Me.txtCustCode.Text))
                            Else
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst("SAMPLE INVOICE", "FINISHED GOODS", strStockLocation, Trim(Me.txtCustCode.Text))
                            End If
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf")
                        Exit Sub
                    End If
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                End If
        End Select
        If Len(mstrItemCode) > 0 Then
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    Call DisplayDetailsInSpread() 'Procedure Call To Select Data >From Sales_Dtl
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Call displayDeatilsfromCustOrdHdrandDtl()
                    Call DisplayCreditTerm()
            End Select
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                'If CDbl(Trim(txtChallanNo.Text)) > 99000000 Then 
                If CDbl(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then '' Changes by priti on 03 Feb 2026 to stop deleted lock invoice
                    Me.CmdGrpChEnt.Enabled(1) = True
                    Me.CmdGrpChEnt.Enabled(2) = True
                End If
            End If
            Me.CmdGrpChEnt.Focus()
        Else
            frmMKTTRN0021.Close()
        End If
        Call ChangeCellTypeStaticText()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From Location Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(Me.txtLocationCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.UNIT_CODE=l.UNIT_CODE AND s.Location_Code=l.Location_Code and datediff(dd,GETDATE(),S.fin_start_date)<=0  and datediff(dd,S.fin_end_date,GETDATE())<=0", , , , , , "s.unit_code")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.UNIT_CODE=l.UNIT_CODE and s.Location_Code=l.Location_Code and datediff(dd,GETDATE(),S.fin_start_date)<=0  and datediff(dd,S.fin_end_date,GETDATE())<=0", , , , , , "s.unit_code")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
        End Select
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdRefNoHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdRefNoHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Details Of Customer Order
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtCustCode.Focus()
            Exit Sub
        End If
        Dim intPos As Short
        strRefAmm = frmMKTTRN0020.SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text)) 'oldvb6 code commented by rajni
        If Len(strRefAmm) > 0 Then
            intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
            mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
            mstrAmmNo = Mid(strRefAmm, intPos + 2, ((Len(Trim(strRefAmm))) - intPos) - 2)
            txtRefNo.Text = Trim(mstrRefNo)
            If CmbInvType.Text.ToUpper = "EXPORT INVOICE" Then
                mstrexportsotype = Find_Value("SELECT EXPORTSOTYPE FROM CUST_ORD_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & txtCustCode.Text & "' AND cust_ref='" & mstrRefNo & "' and amendment_no='" & mstrAmmNo & "'")
                lblexportsodetails.Text = mstrexportsotype
            Else
                lblexportsodetails.Text = ""
            End If
            If txtCarrServices.Enabled Then txtCarrServices.Focus()
        Else
            If txtCarrServices.Enabled Then txtCarrServices.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From SaleTax Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtSaleTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "SaleTax_Code", "SaleTax_Type", "SaleTax_Mst")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), txtSaleTaxType.Text, "SaleTax_Code", "SaleTax_Type", "SaleTax_Mst")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
                    End If
                End If
        End Select
        txtSaleTaxType.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmexport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmexport.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Disply Export details Form
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim arrintExpDetails() As String
        Dim strMode As String
        frmMKTTRN0022.SetDocumentDate = CStr(dtpDateDesc.Value)
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            frmMKTTRN0022.SetCurrencyID = GetCurrencyINSO("ADD")
        Else
            frmMKTTRN0022.SetCurrencyID = GetCurrencyINSO("EDIT")
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strMode = "MODE_VIEW"
                If Len(Trim(strExpDetails)) Then
                    strExpDetails = frmMKTTRN0022.ShowValuestoString(strExpDetails, strMode)
                Else 'STREXPdETAIL IS lOCAL VARIABLE THEN TO ASSIGN VALUES OF STRVALUES
                    strExpDetails = strValues
                    strValues = ""
                    strExpDetails = frmMKTTRN0022.ShowValuestoString(strExpDetails, strMode)
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strMode = "MODE_EDIT"
                If Len(Trim(strExpDetails)) Then
                    strExpDetails = frmMKTTRN0022.ShowValuestoString(strExpDetails, strMode)
                Else 'STREXPdETAIL IS lOCAL VARIABLE THEN TO ASSIGN VALUES OF STRVALUES
                    strExpDetails = strValues
                    strValues = ""
                    strExpDetails = frmMKTTRN0022.ShowValuestoString(strExpDetails, strMode)
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strMode = "MODE_ADD"
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    If Len(Trim(txtRefNo.Text)) = 0 And (Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE") Then
                        Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtRefNo.Focus()
                        Exit Sub
                    Else
                        If Len(Trim(strValues)) = 0 Then
                            strExpDetails = frmMKTTRN0022.ShowValuestoString(strExpDetails, strMode)
                        Else
                            strExpDetails = frmMKTTRN0022.ShowValuestoString(strValues, strMode)
                        End If
                    End If
                End If
        End Select

        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To empower .hlp help
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0005.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlInsurance_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtFreight.Focus()
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlSVD_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlSVD.KeyPress
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ctlInsurance.Focus()
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmEXPTRN0010_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPTRN0010_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        Call EnableControls(False, Me, True)
                        CmbInvType.Enabled = True : CmbInvSubType.Enabled = True
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lblLocCodeDes.Text = ""
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False

                        Me.CmdGrpChEnt.Enabled(1) = False
                        Me.CmdGrpChEnt.Enabled(2) = False
                        Me.CmdGrpChEnt.Enabled(5) = False
                        CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        With Me.SpChEntry
                            .MaxRows = 1 : .set_RowHeight(1, 11)
                            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
                        End With
                        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                        dtpDateDesc.Visible = False
                        txtLocationCode.Focus()
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
    Private Sub frmEXPTRN0010_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To initilise Values on Form Load
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim blnyes As Boolean
        strValues = ""
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt)
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        CmdRefNoHelp.Image = My.Resources.ico111.ToBitmap
        Call EnableControls(False, Me, True)

        dtpDateDesc.Format = DateTimePickerFormat.Custom
        dtpDateDesc.CustomFormat = gstrDateFormat
        dtpDateDesc.Value = GetServerDate()
        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)  'FormatDateTime(GetServerDate(),)
        With dtpDateDesc
            .Value = ConvertToDate(lblDateDes.Text)
            .Visible = False
        End With
        Call AddTransPortTypeToCombo()
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
        mblnWithPayExpInvoice = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(WithPayExpInvoice,0) from sales_parameter where unit_code='" & gstrUNITID & "'"))

        Me.CmdGrpChEnt.Enabled(1) = False
        Me.CmdGrpChEnt.Enabled(2) = False
        Me.CmdGrpChEnt.Enabled(5) = False

        With Me.SpChEntry
            .Row = 0 : .Col = 1 : .Text = "Item Code" : .set_ColWidth(1, 20)
            .Row = 0 : .Col = 2 : .Text = "Drawing No." : .set_ColWidth(2, 20)
            .Row = 0 : .Col = 3 : .Text = "Rate" : .set_ColWidth(3, 11)
            .Row = 0 : .Col = 4 : .Text = "Cust Material"
            .Row = 0 : .Col = 5 : .Text = "Quantity"
            .Row = 0 : .Col = 6 : .Text = "Packing"
            .Row = 0 : .Col = 7 : .Text = "Others"
            .Row = 0 : .Col = 8 : .Text = "From Box"
            .Row = 0 : .Col = 9 : .Text = "To Box"
            .MaxCols = 15
        End With

        'GST DETAILS
        If gblnGSTUnit = True Then
            With SpChEntry
                '.MaxCols = EnumInv.IGSTTXRT_TYPE
                .Row = 0 : .Col = 10 : .Text = "HSN/SAC CODE"
                .Row = 0 : .Col = 11 : .Text = "CGST TAX"
                .Row = 0 : .Col = 12 : .Text = "SGST TAX"
                .Row = 0 : .Col = 13 : .Text = "UTGST TAX"
                .Row = 0 : .Col = 14 : .Text = "IGST TAX"
                .Row = 0 : .Col = 15 : .Text = "COMPENSATION CESS"
            End With
        End If
        'GST DETAILS
        Call SelectInvoiceTypeFromSaleConf()
        Call addRowAtEnterKeyPress(1)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub frmEXPTRN0010_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        If CmbInvType.Items.Count <= 0 Then
            MsgBox("No Data Defined for this financial Year For Export Invoice in Sales Conf.", MsgBoxStyle.Information, ResolveResString(100))
            Me.Close() : Exit Sub
        End If
        mdifrmMain.CheckFormName = mintIndex
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub frmEXPTRN0010_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmEXPTRN0010_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Do Coding Same As Close Button Click.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrpChEnt_ButtonClick(CmdGrpChEnt, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpChEnt.Focus()
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmEXPTRN0010_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmMKTTRN0020 = Nothing 'Assign form to nothing
        frmMKTTRN0021 = Nothing 'Assign form to nothing
        frmMKTTRN0022 = Nothing 'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Add Row At Enter Key Press Of Last Column Of Spread
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .ColsFrozen = 1 : .ColsFrozen = 2
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 11)
                .Col = 4 ''Cust Matt
                .ColHidden = True
                .Col = 6 ''Packing
                .ColHidden = True
                .Col = 7 ''Others
                .ColHidden = True
                If mblnWithPayExpInvoice = False Then
                    .Col = 10 ''HSN
                    .ColHidden = True
                    .Col = 11 ''CGST
                    .ColHidden = True
                    .Col = 12 ''SGST
                    .ColHidden = True
                    .Col = 13 ''UGST
                    .ColHidden = True
                    .Col = 14 ''IGST
                    .ColHidden = True
                    .Col = 15 ''Compensation
                    .ColHidden = True
                End If
                ''Addition Ends
            Next intRowHeight
            If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function SelectInvoiceTypeFromSaleConf() As Object
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Select Invoice Type,Invoice SubTypeDescription From SaleConf
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Description) from SaleConf where unit_code='" & gstrUNITID & "' and Invoice_Type in('EXP') and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            intRecCount = rsSaleConf.GetNoRows
            rsSaleConf.MoveFirst()
            For intLoopCounter = 0 To intRecCount - 1
                VB6.SetItemString(CmbInvType, intLoopCounter, rsSaleConf.GetValue("Description"))
                rsSaleConf.MoveNext()
            Next intLoopCounter
        End If
        rsSaleConf.ResultSetClose()

        rsSaleConf = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub SelectInvoiceSubTypeFromSaleConf(ByRef pstrInvType As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Select Invoice SubTypeDescription From SaleConf Acc. to Inv. Type
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf where unit_code='" & gstrUNITID & "' and  Description='" & Trim(pstrInvType) & "'and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            intRecCount = rsSaleConf.GetNoRows
            rsSaleConf.MoveFirst()
            CmbInvSubType.Items.Clear()
            For intLoopCounter = 0 To intRecCount - 1
                VB6.SetItemString(CmbInvSubType, intLoopCounter, rsSaleConf.GetValue("Sub_Type_Description"))
                rsSaleConf.MoveNext()
            Next intLoopCounter
            CmbInvSubType.SelectedIndex = 0
        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        'Return Value       :   NA
        'Function           :   To Select The Field Description In The Description Labels
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and unit_code='" & gstrUNITID & "'"
        rsDescription = New ClsResultSetDB
        rsDescription.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDescription.GetNoRows > 0 Then
            pContrName.Text = rsDescription.GetValue(Trim(pstrFieldName1))
        End If
        rsDescription.ResultSetClose()
        rsDescription = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub SpChEntry_EnterRow(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EnterRowEvent) Handles SpChEntry.EnterRow

        'Function           :   to check If Grid have items in it or not
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim StrItemCode As String
        StrItemCode = Replace(mstrItemCode, "'", "")
        If Len(Trim(StrItemCode)) = 0 Then
            MsgBox("First Select atleast one item First", MsgBoxStyle.OkOnly, ResolveResString(100))
            If Cmditems.Enabled = True Then
                Cmditems.Focus()
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub SpChEntry_KeyPressEvent1(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpChEntry.KeyPressEvent
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96, 45
                e.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAddExciseDuty_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtAddExciseDuty.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ctlSVD.Focus()
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAnnex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnnex.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If txtCarrServices.Enabled Then txtCarrServices.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtBankAc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBankAc.TextChanged
        On Error GoTo ErrHandler
        lblAcCodeDes.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtBankAc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBankAc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtRemarks.Focus()
                Call txtBankAc_Validating(txtBankAc, New System.ComponentModel.CancelEventArgs(False))
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtBankAc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtBankAc.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim clsBankMster As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtBankAc.Text) > 0 Then
            clsBankMster = New ClsResultSetDB
            clsBankMster.GetResult("Select Bnk_Bankid,Bnk_accNo from Gen_bankMaster where unit_code='" & gstrUNITID & "' and  bnk_Bankid ='" & Trim(txtBankAc.Text) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If clsBankMster.GetNoRows > 0 Then
                clsBankMster.MoveFirst()
                lblAcCodeDes.Text = clsBankMster.GetValue("Bnk_accNo")
                If Cmditems.Enabled Then Cmditems.Focus()
                clsBankMster.ResultSetClose()
                clsBankMster = Nothing
            Else
                MsgBox("Bank Code is not Valid.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                Cancel = True
                txtBankAc.Text = ""
                If txtBankAc.Enabled Then txtBankAc.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtCarrServices_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCarrServices.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        CmbTransType.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh Value of other controls when Challan No changes
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call RefreshForm("CHALLAN")
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.Enter
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show Data Selected.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Me.txtChallanNo.SelectionStart = 0
        Me.txtChallanNo.SelectionLength = Len(Me.txtChallanNo.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtChallanNo.Text) > 0 Then
                            Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.CmdGrpChEnt.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtCustCode.Focus()
                End Select
        End Select
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
            KeyAscii = 0
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtChallanNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtChallanNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdChallanNo.Enabled Then Call CmdChallanNo_Click(CmdChallanNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtChallanNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Check Validity Of Challan No. In SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtChallanNo.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "SalesChallan_Dtl") Then
                        If Len(txtLocationCode.Text) > 0 Then
                            If GetDataInViewMode() Then 'if record found
                                Cmditems.Enabled = True
                                cmexport.Enabled = True
                                Cmditems.Focus()
                            Else 'if no record found then display message
                                Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Cmditems.Enabled = False
                                txtLocationCode.Focus()
                            End If
                        Else 'if location code field is blank
                            Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtLocationCode.Focus()
                        End If
                    Else 'If Doc_No Is Invalid
                        Call ConfirmWindow(10404, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Text = ""
                    End If
                End If
        End Select
        If Val(txtChallanNo.Text) > 99000000 Then
            Cmditems.Enabled = True
        Else
            CmdGrpChEnt.Enabled(1) = False
            CmdGrpChEnt.Enabled(2) = False
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh the Values of Other Controls when Customer code changes.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            lblCustCodeDes.Text = ""
            txtRefNo.Text = ""
            SpChEntry.MaxRows = 0
            mstrItemCode = ""
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            txtCustCode.Focus()
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            CmdGrpChEnt.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (CmbInvType.Text = "NORMAL INVOICE") Or (CmbInvType.Text = "JOBWORK INVOICE") Then
                                txtRefNo.Focus()
                            Else
                                If txtCarrServices.Enabled Then txtCarrServices.Focus()
                            End If

                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdCustCodeHelp.Enabled Then Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Validate Customer Code Entered by User
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtCustCode.Text) > 0 Then
                    If Trim(mstrInvoiceType) = "EXP" Then
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst") Then
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                            If (CmbInvType.Text = "EXPORT INVOICE") And UCase(Trim(Me.CmbInvSubType.Text)) <> "SAMPLE" Then
                                txtRefNo.Focus()
                            Else
                                If txtCarrServices.Enabled Then txtCarrServices.Focus()
                            End If
                        Else
                            Cancel = True
                            Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtCustCode.Text = ""
                            txtCustCode.Focus()
                        End If
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtExciseDuty_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtExciseDuty.KeyPress
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtAddExciseDuty.Focus()
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtFreight_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtFreight.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If (CmbInvType.Text = "SAMPLE INVOICE") Or (CmbInvType.Text = "TRANSFER INVOICE") Or (CmbInvType.Text = "JOBWORK INVOICE") Then
                            txtSurcharge.Focus()
                        Else
                            txtSaleTaxType.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtSaleTaxType.Enabled Then
                            txtSaleTaxType.Focus()
                        Else
                            txtSurcharge.Focus()
                        End If
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Set The Values of Related control on Change of Location Code
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call RefreshForm("LOCATION")
            End Select
        End If
        txtCustCode.Text = ""
        lblCustCodeDes.Text = ""
        txtRefNo.Text = ""
        SpChEntry.MaxRows = 0
        mstrItemCode = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Enter
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to Show Selected Text
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Me.txtLocationCode.SelectionStart = 0
        Me.txtLocationCode.SelectionLength = Len(Me.txtLocationCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtLocationCode.Text) > 0 Then
                            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.CmdGrpChEnt.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtLocationCode.Text) > 0 Then
                            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtCustCode.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLocationCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectInvTypeSubTypeFromSaleConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Select Invoice Type,Sub Type From Sale Conf
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        strSaleConfSql = "Select Invoice_Type,Sub_Type from SaleConf where unit_code='" & gstrUNITID & "' and Description='" & Trim(pstrInvType) & "'"
        strSaleConfSql = strSaleConfSql & " and Sub_Type_Description='" & Trim(pstrInvSubtype) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
        End If
        rsSaleConf.ResultSetClose()

        rsSaleConf = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Check Validity Of Location Code In The Location_Mst
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SalesChallan_Dtl") Then
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Cancel = True
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf", " datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0") Then
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Cancel = True
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Reset the values of related control on change of Refrence No
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SpChEntry.MaxRows = 0
            mstrItemCode = ""
            txtRefNo.Focus()
            lblCreditTerm.Text = ""
            lblCreditTermDesc.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRefNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtRefNo_Validating(txtRefNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (CmbInvType.Text = "JOBWORK INVOICE") Then
                                txtAnnex.Focus()
                            Else
                                If txtCarrServices.Enabled Then txtCarrServices.Focus()
                            End If

                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRefNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtRefNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From Customer Master/Vendor Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdRefNoHelp.Enabled Then Call CmdRefNoHelp_Click(CmdRefNoHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtRefNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Validate Refrance No Entered by User in Text Box.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If Len(txtRefNo.Text) > 0 Then
                If CheckExistanceOfFieldData((txtRefNo.Text), "Cust_ref", "Cust_ord_Hdr") Then
                    If CmbInvType.Text = "EXPORT INVOICE" Then
                        If txtCarrServices.Enabled Then txtCarrServices.Focus()
                    Else
                        'txtAnnex.SetFocus
                    End If
                Else
                    Call ConfirmWindow(10436, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Cancel = True
                    txtRefNo.Text = ""
                    txtRefNo.Focus()
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        Cmditems.Focus()
                        Call txtBankAc_Validating(txtBankAc, New System.ComponentModel.CancelEventArgs(False))
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSalesTax_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtSalesTax.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtSurcharge.Focus()
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtSalesTax.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSaleTaxType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSaleTaxType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show Help on Sales Tax Type.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdSaleTaxType.Enabled Then Call CmdSaleTaxType_Click(CmdSaleTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSaleTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Validate Data of Sales Tax Type entered by user
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "SaleTax_Code", "SaleTax_Mst") Then
                txtSalesTax.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSaleTaxType.Text = ""
                txtSaleTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSurcharge_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtSurcharge.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case eventArgs.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                End Select
            Case 39, 34, 96
                eventArgs.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtVehNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVehNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                chkServiceInvFormat.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrFieldText - Field Text,pstrColumnName - Column Name
        '                       pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        'Return Value       :   NA
        'Function           :   To Check Validity Of Field Data Whethet it Exists In The
        '                       Database Or Not
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = strTableSql & " AND " & pstrCondition
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()

        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetDataInViewMode() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To display data in view mode from SalasChallan_Dtl,Sales_Dtl acc.to LacationCode & Challan_No.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        GetDataInViewMode = False
        Dim strGetData As String
        Dim rsGetData As ClsResultSetDB
        Dim rsBankMaster As ClsResultSetDB
        Dim strSalesChallanDtl As String

        strSalesChallanDtl = "SELECT Transport_type,Vehicle_No,Account_code,Cust_ref,salesTax_type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Tax, "
        strSalesChallanDtl = strSalesChallanDtl & "Amendment_No,ref_doc_no,"
        strSalesChallanDtl = strSalesChallanDtl & "Currency_Code,Originstatus,ctry_destination_goods,Pre_Carriage_by,"
        strSalesChallanDtl = strSalesChallanDtl & "Receipt_PreCarriage_at,Port_of_loading,Port_of_Discharge,"
        strSalesChallanDtl = strSalesChallanDtl & "nature_of_contract,Final_destination,Mode_of_shipment,Dispatch_Mode,"
        strSalesChallanDtl = strSalesChallanDtl & "Delivery_terms,Payment_terms,Buyer_Description_of_goods,Invoice_Description_of_EPC,"
        strSalesChallanDtl = strSalesChallanDtl & "Exchange_Rate,Frieght_amount,Exchange_Date,other_ref,buyer_id,ServiceInvoiceformatExport,CustBankID,Remarks,Prev_Yr_ExportSales,Permissible_Limit_SmpExport,varGeneralRemarks,exportsotype"
        strSalesChallanDtl = strSalesChallanDtl & " From Saleschallan_dtl where unit_code='" & gstrUNITID & "' and Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtCustCode.Text = rsGetData.GetValue("Account_code")
            txtRefNo.Text = rsGetData.GetValue("Cust_ref")
            txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
            ctlInsurance.Text = rsGetData.GetValue("Insurance")
            txtFreight.Text = rsGetData.GetValue("Frieght_tax")
            lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
            mstrAmmendmentNo = rsGetData.GetValue("Amendment_No")
            lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)
            mstrInvType = rsGetData.GetValue("Invoice_Type")
            mstrInvSubType = rsGetData.GetValue("Sub_Category")
            lblexportsodetails.Text = rsGetData.GetValue("exportsotype")
            Me.CmbInvType.Items.Clear()
            Call SelectInvoiceTypeFromSaleConf()
            Me.CmbInvSubType.Items.Clear()
            Call SelectInvoiceSubTypeFromSaleConf("EXPORT INVOICE")
            If Me.CmbInvType.Items.Count > -1 Then Me.CmbInvType.SelectedIndex = 0
            If mstrInvSubType = "S" Then
                If Me.CmbInvSubType.Items.Count > -1 Then
                    Me.CmbInvSubType.SelectedIndex = 1
                End If
            Else
                If Me.CmbInvSubType.Items.Count > -1 Then
                    Me.CmbInvSubType.SelectedIndex = 0
                End If
            End If

            If rsGetData.GetValue("ServiceInvoiceformatExport") = True Then
                chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            lblCreditTerm.Text = IIf(IsDBNull(rsGetData.GetValue("payment_terms")), "", rsGetData.GetValue("payment_terms"))
            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
            Else
                lblCreditTermDesc.Text = ""
            End If
            txtBankAc.Text = Trim(rsGetData.GetValue("CustBankID"))
            rsBankMaster = New ClsResultSetDB
            rsBankMaster.GetResult("Select bnk_accno from gen_bankMaster where unit_code='" & gstrUNITID & "' and bnk_Bankid = '" & Trim(txtBankAc.Text) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsBankMaster.GetNoRows > 0 Then
                lblAcCodeDes.Text = rsBankMaster.GetValue("bnk_accno")
            End If

            rsBankMaster.ResultSetClose()
            rsBankMaster = Nothing
            txtRemarks.Text = Trim(rsGetData.GetValue("Remarks"))
            CmbInvType.SelectedIndex = 0
            strExpDetails = ""
            strExpDetails = rsGetData.GetValue("Currency_Code") & "§" & rsGetData.GetValue("Originstatus") & "§" & rsGetData.GetValue("ctry_destination_goods") & "§" & rsGetData.GetValue("Pre_Carriage_by") & "§" & rsGetData.GetValue("Receipt_PreCarriage_at") & "§" & rsGetData.GetValue("Port_of_loading") & "§" & rsGetData.GetValue(" Port_of_Discharge") & "§" & rsGetData.GetValue("nature_of_contract") & "§" & rsGetData.GetValue("Final_destination") & "§" & rsGetData.GetValue("Mode_of_shipment") & "§" & rsGetData.GetValue("Dispatch_Mode") & "§" & rsGetData.GetValue("Delivery_terms") & "§" & rsGetData.GetValue("Payment_terms") & "§" & rsGetData.GetValue("Buyer_Description_of_goods") & "§" & rsGetData.GetValue("Invoice_Description_of_EPC") & "§" & rsGetData.GetValue("Exchange_Rate") & "§" & rsGetData.GetValue("Frieght_amount") & "§" & rsGetData.GetValue("Exchange_Date") & "§" & rsGetData.GetValue("other_ref") & "§" & rsGetData.GetValue("buyer_id") & "§" & rsGetData.GetValue("Prev_Yr_ExportSales") & "§" & rsGetData.GetValue("Permissible_Limit_SmpExport") & "§" & rsGetData.GetValue("varGeneralRemarks")

        Else
            GetDataInViewMode = False
        End If
        rsGetData.ResultSetClose()

        rsGetData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function DisplayDetailsInSpread() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intRecordCount As Short
        Dim strsaledtl As String
        Dim rsSalesDtl As ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsaledtl = ""
                strsaledtl = "Select Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,GL_code,SL_code,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_amount,pervalue,TotalExciseAmount,SupplementaryInvoiceFlag,To_Location,Discount_type,Discount_amt,Discount_perc,From_Location,Cust_ref,Amendment_No,SRVDINO,SRVLocation,USLOC,SchTime,BinQuantity,Packing_Type,ItemPacking_Amount,Item_remark,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS_TYPE from Sales_Dtl where unit_code='" & gstrUNITID & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") And Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE" Then
                    strsaledtl = ""
                    strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Others,tool_Cost,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  AS COMPENSATION_CESS_TYPE   from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and "
                    strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                    strsaledtl = strsaledtl & txtRefNo.Text & "' and Amendment_No = '" & mstrAmmNo & "'and "
                    strsaledtl = strsaledtl & " Active_flag ='A' and Cust_DrgNo in(" & mstrItemCode & ")"
                Else
                    strsaledtl = ""
                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst where unit_code='" & gstrUNITID & "' and "
                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                End If
        End Select
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopcount As Short
        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows
            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i
            Call addRowAtEnterKeyPress(intRecordCount - 1)
            rsSalesDtl.MoveFirst()
            rsSalesDtl.MoveFirst()
            For intLoopCounter = 1 To intRecordCount
                With Me.SpChEntry
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
                            .set_RowHeight(intLoopCounter, 11)
                            Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                            Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                            Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                            Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                            Call .SetText(5, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                            mdblPrevQty(intLoopCounter - 1) = Nothing
                            Call .GetText(5, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
                            Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                            Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Others"))
                            Call .SetText(8, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                            Call .SetText(9, intLoopCounter, rsSalesDtl.GetValue("To_Box"))

                            Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                            Call .SetText(11, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                            Call .SetText(12, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                            Call .SetText(13, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                            Call .SetText(14, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                            Call .SetText(15, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_TYPE"))
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            .Enabled = True
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .BlockMode = True : .Lock = False : .BlockMode = False
                            If (Trim(CmbInvType.Text) = "EXPORT INVOICE") And Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE" Then
                                Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
                                Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                                Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                                Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                                Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Others"))
                            Else
                                Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Standard_Rate"))
                            End If

                            Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                            Call .SetText(11, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                            Call .SetText(12, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                            Call .SetText(13, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                            Call .SetText(14, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                            Call .SetText(15, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_TYPE"))
                    End Select
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter
        End If
        rsSalesDtl.ResultSetClose()

        rsSalesDtl = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function ValidatebeforeSave(ByRef pstrMode As String) As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check the Blank Fields In The Form
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        Select Case UCase(Trim(pstrMode))
            Case "ADD"
                If (Len(Me.txtLocationCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtLocationCode
                    End If
                    ValidatebeforeSave = False
                End If
                If (Len(Me.txtCustCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                If Not DateIsAppropriate() Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Date specified either falls Before the LAST Invoice Date or is Greater than Todays Date."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                If SpChEntry.MaxRows = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Select Items"
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Cmditems
                    End If
                    ValidatebeforeSave = False
                End If


                If (Len(Me.txtAddExciseDuty.Text) = 0) Then
                    txtAddExciseDuty.Text = "0.00"
                End If

                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If

                If (Len(Me.txtSurcharge.Text) = 0) Then
                    txtSurcharge.Text = "0.00"
                End If

                If (Len(Me.ctlSVD.Text) = 0) Then
                    ctlSVD.Text = "0.00"
                End If

                If (Len(Me.ctlInsurance.Text) = 0) Then
                    ctlInsurance.Text = "0.00"
                End If
            Case "EDIT"
                '*****

                If (Len(Me.txtAddExciseDuty.Text) = 0) Then
                    txtAddExciseDuty.Text = "0.00"
                End If

                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If

                If (Len(Me.txtSurcharge.Text) = 0) Then
                    txtSurcharge.Text = "0.00"
                End If
        End Select
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function
    Private Sub ChangeCellTypeStaticText()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Change The Cell Type In Spread Control to Cell Type Static Text to
        '                       Make Cell Type UnEditable
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intcol As Short
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    Else
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If (UCase(strInvType) = "EXP") Then
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    Else
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
            End Select
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function QuantityCheck() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check Schedule Quantity From DailyMktSchedule/MonthlyMktSchedule
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        'Revision  By       : Ashutosh , Issue Id :19385
        'Revision On        : 29-01-2007
        'History            : Wrong Schedule Checking in Export Documentation Invoice.(WMART  Kandla)
        '---------------------------------------------------------------------------------------

        On Error GoTo ErrHandler
        QuantityCheck = False
        Dim strScheduleSql As String
        Dim rsMktSchedule As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim strQuantity As String
        Dim intRwCount As Short 'To Count No. Of Rows
        Dim intLoopcount As Short
        Dim varItemQty As Object 'To Get Quantity Acc. To Drawing No and Item Code
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim strItembal As String
        Dim PresQty As Object
        Dim intcol As Short
        Dim intFromBox As Short
        rsMktSchedule = New ClsResultSetDB
        For intRwCount = 1 To SpChEntry.MaxRows
            For intcol = 1 To SpChEntry.MaxCols
                SpChEntry.Col = intcol
                If (SpChEntry.Col = 5) Or (SpChEntry.Col = 3) Or (SpChEntry.Col = 9) Or (SpChEntry.Col = 8) Then
                    SpChEntry.Row = intRwCount
                    If (Val(Trim(SpChEntry.Text)) = 0) Then
                        QuantityCheck = True
                        Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        SpChEntry.Row = intRwCount : SpChEntry.Col = intcol : SpChEntry.Action = 0 : SpChEntry.Focus()
                        Exit Function
                    End If
                    If (SpChEntry.Col = 9) Then
                        SpChEntry.Row = intRwCount : SpChEntry.Col = 8 : intFromBox = Val(Trim(SpChEntry.Text))
                        SpChEntry.Row = intRwCount : SpChEntry.Col = 9
                        If Val(Trim(SpChEntry.Text)) < intFromBox Then
                            QuantityCheck = True
                            Call ConfirmWindow(10235, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 9 : SpChEntry.Action = 0 : SpChEntry.Focus()
                            Exit Function
                        End If
                    End If
                End If
            Next intcol
        Next intRwCount
        For intRwCount = 1 To SpChEntry.MaxRows
            varItemCode = Nothing
            varItemQty = Nothing
            Call SpChEntry.GetText(1, intRwCount, varItemCode)
            Call SpChEntry.GetText(5, intRwCount, varItemQty)
            If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount) = False Then
                QuantityCheck = True
                Exit Function
            End If
        Next
        Dim strMakeDate As String
        If ((Trim(CmbInvType.Text) = "EXPORT INVOICE") And Trim(UCase(Me.CmbInvSubType.Text)) <> "SAMPLE") Then
            strScheduleSql = ""
            strScheduleSql = "Select Quantity=Schedule_Quantity-isnull(Despatch_Qty,0),Cust_DrgNo,Item_Code from DailyMktSchedule where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and Status =1"

            rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsMktSchedule.GetNoRows > 0 Then 'If Record Found
                rsMktSchedule.ResultSetClose()
                rsMktSchedule = Nothing
                For intRwCount = 1 To Me.SpChEntry.MaxRows
                    varItemCode = Nothing
                    Call Me.SpChEntry.GetText(2, intRwCount, varItemCode)
                    strScheduleSql = ""
                    strScheduleSql = "Select Quantity=Schedule_Quantity-isnull(Despatch_Qty,0),Cust_DrgNo,Item_Code from DailyMktSchedule where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                    strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo in('" & Trim(varItemCode) & "') and Status =1"
                    rsMktSchedule = New ClsResultSetDB
                    rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsMktSchedule.GetNoRows > 0 Then 'If Record Found
                        strQuantity = rsMktSchedule.GetValue("Quantity")
                    Else
                        strQuantity = CStr(0)
                        MsgBox("No Daily Schedule Defined For Selected Item - " & varItemCode & " .Define Schedule First")
                        QuantityCheck = True
                        rsMktSchedule.ResultSetClose()
                        rsMktSchedule = Nothing
                        Exit Function
                    End If
                    Select Case CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                            strQuantity = Val(rsMktSchedule.GetValue("Quantity")) + mdblPrevQty(intRwCount - 1)
                    End Select
                    rsMktSchedule.ResultSetClose()
                    rsMktSchedule = Nothing
                    If Val(varItemQty) > Val(strQuantity) Then
                        QuantityCheck = True
                        MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity & " For Item " & varItemCode)
                        With Me.SpChEntry
                            .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                        Exit Function
                    Else
                        QuantityCheck = False
                    End If
                Next intRwCount
                mstrUpdDispatchSql = ""
                For intLoopcount = 1 To SpChEntry.MaxRows
                    varDrgNo = Nothing
                    PresQty = Nothing
                    Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                    Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                    mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                    mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & ")"
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Status =1" & vbCrLf
                Next
            Else
                If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                Else
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                End If

                For intRwCount = 1 To Me.SpChEntry.MaxRows

                    varItemCode = Nothing
                    Call Me.SpChEntry.GetText(2, intRwCount, varItemCode)

                    strScheduleSql = "Select Quantity=Schedule_Qty-isnull(Despatch_Qty,0), isnull(Despatch_Qty,0) as Despatch_Qty from MonthlyMktSchedule where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                    strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo in('" & Trim(varItemCode) & "') and status =1"
                    'rsMktSchedule.ResultSetClose()
                    rsMktSchedule = New ClsResultSetDB
                    rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsMktSchedule.GetNoRows > 0 Then
                        strQuantity = rsMktSchedule.GetValue("Quantity")
                    Else
                        strQuantity = CStr(0)
                        MsgBox("No Schedule Defined For These Selected Items,Define Schedule First")
                        QuantityCheck = True
                        rsMktSchedule.ResultSetClose()
                        rsMktSchedule = Nothing
                        Exit Function
                    End If

                    Select Case CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                            strQuantity = Val(rsMktSchedule.GetValue("Quantity")) + mdblPrevQty(intRwCount - 1)
                    End Select
                    rsMktSchedule.ResultSetClose()
                    rsMktSchedule = Nothing
                    If Val(varItemQty) > Val(strQuantity) Then
                        QuantityCheck = True
                        MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity & " for Item " & varItemCode)
                        With Me.SpChEntry
                            .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                        Exit Function
                    Else
                        QuantityCheck = False
                    End If
                Next intRwCount
                mstrUpdDispatchSql = ""
                For intLoopcount = 1 To SpChEntry.MaxRows
                    varDrgNo = Nothing
                    PresQty = Nothing
                    Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                    Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                    mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                    mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & ")"
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Where unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                    mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and status =1" & vbCrLf
                Next
            End If
        End If
        Dim strItCode As String 'To Make Item Code String
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            varItemCode = Nothing
            Call Me.SpChEntry.GetText(1, intRwCount, varItemCode)
            strItCode = strItCode & "'" & Trim(varItemCode) & "',"
        Next intRwCount
        strItCode = Mid(strItCode, 1, Len(strItCode) - 1)
        rsSaleConf = New ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If mstrInvType = "" Then mstrInvType = "EXP"
                If mstrInvSubType = "" Then mstrInvSubType = IIf(UCase(CmbInvSubType.Text) = "SAMPLE", "S", "E")
                rsSaleConf.GetResult("select Stock_Location From saleconf where unit_code='" & gstrUNITID & "' and invoice_type ='" & Trim(mstrInvType) & "' and sub_type ='" & Trim(mstrInvSubType) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(lblDateDes.Text) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(lblDateDes.Text) & "')<=0", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSaleConf.GetResult("select Stock_Location From saleconf where unit_code='" & gstrUNITID & "' and Description ='" & Trim(CmbInvType.Text) & "' and sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        End Select

        If Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0 Then
            MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.OkOnly, ResolveResString(100))
            QuantityCheck = True
            Exit Function
        End If
        For intRwCount = 1 To Me.SpChEntry.MaxRows

            varItemCode = Nothing
            varItemQty = Nothing
            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
            Call Me.SpChEntry.GetText(1, intRwCount, varItemCode)
            strItembal = "Select Cur_Bal From ItemBal_Mst where unit_code='" & gstrUNITID & "' and Location_Code ='" & Trim(rsSaleConf.GetValue("Stock_Location")) & "' and item_Code ='" & Trim(varItemCode) & "' "
            rsMktSchedule = New ClsResultSetDB
            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsMktSchedule.GetNoRows > 0 Then
                strQuantity = rsMktSchedule.GetValue("Cur_Bal")
            Else
                strQuantity = CStr(0)
            End If
            rsMktSchedule.ResultSetClose()
            rsMktSchedule = Nothing
            If Val(varItemQty) > Val(strQuantity) Then
                QuantityCheck = True
                If CDbl(strQuantity) = 0 Then
                    MsgBox("No Balance Available for Item [" & varItemCode & "]", MsgBoxStyle.OkOnly, ResolveResString(100))
                Else
                    MsgBox("Available Balance for Item [" & varItemCode & "] is " & strQuantity & " at location  " & rsSaleConf.GetValue("Stock_Location"), MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
                With Me.SpChEntry
                    .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End With
                Exit Function
            Else
                QuantityCheck = False
            End If
        Next intRwCount

        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        If UCase(Trim(mstrInvoiceType)) = "JOB" Then
            If BomCheck() = False Then
                QuantityCheck = True
                Exit Function
            Else
                QuantityCheck = False
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshForm(ByRef pstrType As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh All The Fields
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case UCase(pstrType)
            Case "LOCATION"
                txtLocationCode.Text = "" : lblLocCodeDes.Text = ""
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtExciseDuty.Text = ""
                txtAddExciseDuty.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = ""
                txtSalesTax.Text = ""
                txtSurcharge.Text = ""
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtExciseDuty.Text = ""
                txtAddExciseDuty.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = ""
                txtSalesTax.Text = ""
                txtSurcharge.Text = ""
                If CmbInvType.Items.Count > 0 Then
                    CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0
                End If
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
        End Select

        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddTransPortTypeToCombo()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Transport Type in Combo
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        VB6.SetItemString(CmbTransType, 0, "R - Road") 'Road
        VB6.SetItemString(CmbTransType, 1, "L - Rail") 'Rail
        VB6.SetItemString(CmbTransType, 2, "S - Sea") 'Sea
        VB6.SetItemString(CmbTransType, 3, "A - Air") 'Air
        VB6.SetItemString(CmbTransType, 4, "H - Hand") 'Hand
        VB6.SetItemString(CmbTransType, 5, "C - Courier") 'Courier
        CmbTransType.SelectedIndex = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSalesChallanDtl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Select Max.  Challan No. From SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '***************************************************************************
        'Revised by         :   Prashant rajpal
        'Revised Date       :   19/03/2015
        'revised issue id   :   10777177
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As ClsResultSetDB
        Dim strUpdateSQL As String
        '10777177 
        strChallanNo = "Select Current_No From  DocumentType_Mst (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & _
                        " Doc_Type = 9999  AND fin_start_date <= CONVERT(DateTime,'" & dtpDateDesc.Value & "',103) " & _
                        " And Fin_End_date >= Convert(datetime,'" & dtpDateDesc.Value & "',103)"

        'strChallanNo = "Select max(Doc_No) as Doc_No from SalesChallan_Dtl where unit_code='" & gstrUNITID & "' and Doc_No>" & 99000000
        rsChallanNo = New ClsResultSetDB
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = (rsChallanNo.GetValue("Current_No") + 1).ToString

            strUpdateSQL = "UPDATE DocumentType_Mst with (ROWLOCK) Set Current_No = " & CLng(strChallanNo) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strUpdateSQL = strUpdateSQL + " Doc_Type = 9999 AND fin_start_date <= CONVERT(DateTime,'" & dtpDateDesc.Value & "',103) "
            strUpdateSQL = strUpdateSQL + " And Fin_End_date >= Convert(datetime,'" & dtpDateDesc.Value & "',103) "
            mP_Connection.Execute(strUpdateSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            txtChallanNo.Text = strChallanNo
        Else
            MsgBox("Temporary Invoice No. Series Not Define. Invoice Entry Can Not Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
            txtChallanNo.Text = ""
        End If
        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        '10777177 
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub displayDeatilsfromCustOrdHdrandDtl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Sales order details on Selection of Customer Refrance
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strCustOrdHdr As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strCustOrdHdr = "Select max(Order_date),Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax ,SalesTax_Type from Cust_ord_hdr"
                strCustOrdHdr = strCustOrdHdr & " Where unit_code='" & gstrUNITID & "' and Account_code='" & txtCustCode.Text & "' and Cust_Ref ='"
                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
                strCustOrdHdr = strCustOrdHdr & " group by Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax,SalesTax_Type"
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                txtExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Excise_Duty"))
                txtAddExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Extra_Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Extra_Excise_Duty"))
                txtSalesTax.Text = IIf(rsCustOrdHdr.GetValue("Sales_tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Sales_tax"))
                txtSurcharge.Text = IIf(rsCustOrdHdr.GetValue("Surcharge_Sales_Tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Surcharge_Sales_Tax"))
                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                rsCustOrdHdr.ResultSetClose()
                rsCustOrdHdr = Nothing
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

        End Select
        Call DisplayDetailsInSpread()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SetMaxLengthInSpread()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Set Max Length Of Columns Of Spread
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intRow As Short
        With Me.SpChEntry
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 1 : .TypeMaxEditLen = 16
                .Col = 2 : .TypeMaxEditLen = 30
                .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.9999")
                .Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999.99")
                .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.9999")
                .Col = 7 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 8 : .TypeMaxEditLen = 4
                .Col = 9 : .TypeMaxEditLen = 4
            Next intRow
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function DeleteRecords() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Delete Records
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        DeleteRecords = False
        strupSalechallan = "Delete SalesChallan_Dtl where unit_code='" & gstrUNITID & "' and Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallan = strupSalechallan & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"

        strupSaleDtl = "Delete Sales_Dtl where unit_code='" & gstrUNITID & "' and Doc_No =" & Trim(txtChallanNo.Text)
        strupSaleDtl = strupSaleDtl & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"

        DeleteRecords = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As Object, ByRef strQuantity As Object, ByRef intRow As Short) As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strItem - Item code
        '                       strQuantity - Quantity of ITem
        '                       introw -Current Row count in Spread
        'Return Value       :   Boolean YES OR No
        'Function           :   To check if decimal allowed flag is yes or No
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
        strMeasure = strMeasure & " where a.unit_code=b.unit_code and b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "' and a.unit_code='" & gstrUNITID & "'"
        rsMeasure = New ClsResultSetDB
        rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
            rsMeasure.ResultSetClose()
            rsMeasure = Nothing

            If System.Math.Round(Val(strQuantity), 0) - Val(strQuantity) <> 0 Then
                Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                CheckMeasurmentUnit = False
                Call SpChEntry.SetText(5, intRow, CShort(strQuantity))
                SpChEntry.Col = 5
                SpChEntry.Row = SpChEntry.ActiveRow
                SpChEntry.Focus()
                Exit Function
            Else
                CheckMeasurmentUnit = True
            End If
        Else
            rsMeasure.ResultSetClose()
            rsMeasure = Nothing
            CheckMeasurmentUnit = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function BomCheck() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check Bom details & Qty required in Case of Jobwork Challan
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intSpreadRow As Short 'Max count of spread row
        Dim intSpCurrentRow As Short 'Currunt count of spread in loop
        Dim intBomMaxItem As Short 'Max Parent row in Bom_Mst for finished Item
        Dim intCurrentItem As Short 'Current count of row in Parent loop
        Dim intBomMaxRaw As Short 'Max Count of Child Items in Bom_Mst
        Dim intCurrentRaw As Short 'Current raw count in Bom_Mst for finished row
        Dim inti As Short 'To Change Array Size
        Dim intTotalReqQty As Double 'Req_Qty + Waste_Qty in Bom_Mst
        Dim VarFinishedItem As Object 'to get finished Item code from Spread
        Dim VarFinishedQty As Object 'To get Qty of Finished Item from Spread
        Dim strCustAnnexDtl As String
        Dim strBomMst As String
        Dim strBomMstRaw As String
        Dim strBomItem As String
        Dim arrItem() As String
        Dim arrQty() As Double
        Dim strParent As String
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsBomMst As ClsResultSetDB
        Dim rsBomMstRaw As ClsResultSetDB
        rsBomMst = New ClsResultSetDB
        rsCustAnnexDtl = New ClsResultSetDB
        rsBomMstRaw = New ClsResultSetDB
        BomCheck = False
        intSpreadRow = SpChEntry.MaxRows
        inti = 0
        If SpChEntry.MaxRows >= 1 Then
            'Loop for Spread
            For intSpCurrentRow = 1 To intSpreadRow
                VarFinishedItem = Nothing
                VarFinishedQty = Nothing
                With SpChEntry
                    Call .GetText(1, intSpCurrentRow, VarFinishedItem)
                    Call .GetText(5, intSpCurrentRow, VarFinishedQty)
                End With
                'String for Parent Item in Bom_Mst
                strBomMst = "Select distinct(Item_Code),"
                strBomMst = strBomMst & " Bom_level from Bom_Mst where unit_code='" & gstrUNITID & "' and Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                strParent = ""
                For intCurrentItem = 1 To intBomMaxItem
                    strBomItem = rsBomMst.GetValue("Item_code")
                    If Len(Trim(strParent)) > 0 Then
                        strParent = Trim(strParent) & "," & Chr(34) & strBomItem & Chr(34)
                    Else
                        strParent = Chr(34) & strBomItem & Chr(34)
                    End If
                    rsBomMst.MoveNext()
                Next
                rsBomMst.MoveFirst()
                'Loop for Parent Items
                For intCurrentItem = 1 To intBomMaxItem
                    strBomItem = ""
                    strBomItem = rsBomMst.GetValue("Item_code")
                    strParent = Replace(strParent, Chr(34) & strBomItem & Chr(34), Chr(34) & "Found" & Chr(34))
                    strCustAnnexDtl = "Select Item_Code,Balance_qty from CustAnnex_hdr where unit_code='" & gstrUNITID & "' and Customer_code ='"
                    strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' and ref57f4_no ='"
                    strCustAnnexDtl = strCustAnnexDtl & Trim(txtAnnex.Text) & "' and getdate() <= "
                    strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                    strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                    rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsCustAnnexDtl.GetNoRows >= 1 Then
                        rsCustAnnexDtl.MoveFirst()
                        ReDim Preserve arrItem(inti)
                        ReDim Preserve arrQty(inti)

                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                        intTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                        If arrQty(inti) < intTotalReqQty * VarFinishedQty Then
                            MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, ResolveResString(100))  'converted code
                            SpChEntry.Row = intSpCurrentRow
                            SpChEntry.Col = 5
                            SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            BomCheck = False
                            Exit Function
                        End If
                    Else
                        'String for Child Items in Bom_Mst for a Parent Item
                        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
                        strBomMstRaw = strBomMstRaw & " As TotalReqQty from Bom_Mst where unit_code='" & gstrUNITID & "' and "
                        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
                        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
                        strBomMstRaw = strBomMstRaw & VarFinishedItem & "'"
                        rsBomMstRaw = New ClsResultSetDB
                        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        intBomMaxRaw = rsBomMstRaw.GetNoRows
                        rsBomMstRaw.MoveFirst()
                        'Loop for Child Items
                        For intCurrentRaw = 1 To intBomMaxRaw
                            strBomItem = ""
                            strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                            intTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")
                            strCustAnnexDtl = "Select Item_Code,Balance_qty from CustAnnex_hdr where unit_code='" & gstrUNITID & "' and Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' and ref57f4_no ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtAnnex.Text) & "'  and getdate() <= "
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                            rsCustAnnexDtl.ResultSetClose()
                            rsCustAnnexDtl = Nothing
                            rsCustAnnexDtl = New ClsResultSetDB
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then
                                rsCustAnnexDtl.MoveFirst()
                                ReDim Preserve arrItem(inti)
                                ReDim Preserve arrQty(inti)
                                arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                If arrQty(inti) < intTotalReqQty * VarFinishedQty Then
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, ResolveResString(100))
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                End If
                            Else
                                If InStr(1, strParent, Chr(34) & strBomItem & Chr(34), CompareMethod.Text) = 0 Then
                                    MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, ResolveResString(100))
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                End If
                            End If
                            rsBomMstRaw.MoveNext()
                        Next  'Child Item Loop
                    End If
                    rsBomMst.MoveNext()
                    inti = inti + 1
                Next  'Parent Item Loop
                intSpCurrentRow = intSpCurrentRow + 1
            Next  'Spread Item Loop
        End If
        rsBomMst.ResultSetClose()
        rsBomMst = Nothing
        rsCustAnnexDtl.ResultSetClose()
        rsCustAnnexDtl = Nothing
        rsBomMstRaw.ResultSetClose()
        rsBomMstRaw = Nothing
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrItemCode - Item Code to be Calculated from BOM
        '                       pstrfinished - Finished Product code For which invoice has to be done
        'Return Value       :   Quantity
        'Function           :   To Used in Jobwork invoice while Bom consideration
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB

        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where unit_code='" & gstrUNITID & "' and finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
        rsParentQty = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String, ByRef pstrCondition As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrInvType - Invoice type
        '                       pstrInvSubtype -invoice Sub type
        '                       pstrFeild Feild - to Be Selected
        '                       pstrCondition - Condition in Query
        'Return Value       :   Stock Location as String
        'Function           :   To Check Stock Location Acc to Selected invoice type & sub type.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where unit_code='" & gstrUNITID & "' and Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' and " & pstrCondition, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where unit_code='" & gstrUNITID & "' and Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' and " & pstrCondition, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        rsSalesConf.ResultSetClose()
        rsSalesConf = Nothing
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Sub EDitExpDetails()
        On Error GoTo ErrHandler
        strExpEditDetails = ""
        strExpEditDetails = ArrExpDetails(0) & "§" & ArrExpDetails(1) & "§" & ArrExpDetails(2) & "§" & ArrExpDetails(3) & "§" & ArrExpDetails(4) & "§" & ArrExpDetails(5) & "§" & ArrExpDetails(6) & "§" & ArrExpDetails(7) & "§" & ArrExpDetails(8) & "§" & ArrExpDetails(9) & "§" & ArrExpDetails(10) & "§" & ArrExpDetails(11) & "§" & ArrExpDetails(12) & "§" & ArrExpDetails(13) & "§" & ArrExpDetails(14) & "§" & ArrExpDetails(15) & "§" & ArrExpDetails(16) & "§" & ArrExpDetails(17)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function DateIsAppropriate() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Checks for Specified Date is within LIMITs From SalesChallan_DTL
        'Comments           :   NA
        'Creation Date      :   27/06/2002
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim MaxInvoiceDate As DateTime 'Get Max Date of Last Invoice made
        Dim CurrentDate As DateTime
        MaxInvoiceDate = CDate(SelectDataFromTable("INVOICE_DATE", "SalesChallan_Dtl", " WHERE UNIT_CODE='" & gstrUNITID & "' AND BILL_FLAG = 1 and invoice_type = 'EXP' ORDER BY INVOICE_DATE "))
        CurrentDate = GetServerDate()
        If (CurrentDate >= dtpDateDesc.Value) And (dtpDateDesc.Value >= MaxInvoiceDate) Then
            DateIsAppropriate = True
        Else
            DateIsAppropriate = False
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Get Data from BackEnd
        'Comments           :   NA
        'Creation Date      :   27/06/2002
        '*******************************************************************************
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = CStr(GetServerDate())
            End If
        Else
            SelectDataFromTable = CStr(GetServerDate())
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function GetCurrencyINSO(ByVal pstrMode As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Get Data from BackEnd
        'Comments           :   NA
        'Creation Date      :   27/06/2002
        '*******************************************************************************
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        If Trim(pstrMode) = "ADD" Then
            StrSQLQuery = "SELECT Currency_code FROM Cust_Ord_Hdr WHERE UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & Trim(txtCustCode.Text) & "' AND Cust_Ref='" & Trim(txtRefNo.Text) & "' AND Po_Type='E'"
            GetDataFromTable = New ClsResultSetDB
            If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
                If GetDataFromTable.GetNoRows > 0 Then
                    GetCurrencyINSO = GetDataFromTable.GetValue("Currency_Code")
                Else
                    GetCurrencyINSO = ""
                End If
            Else
                GetCurrencyINSO = ""
            End If
        Else
            StrSQLQuery = "SELECT Currency_code FROM SalesChallan_dtl WHERE unit_code='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "' AND Doc_No=" & Trim(txtChallanNo.Text)
            GetDataFromTable = New ClsResultSetDB
            If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
                If GetDataFromTable.GetNoRows > 0 Then
                    GetCurrencyINSO = GetDataFromTable.GetValue("Currency_Code")
                Else
                    GetCurrencyINSO = ""
                End If
            Else
                GetCurrencyINSO = ""
            End If
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GetCurrencyINSO = ""
    End Function

    Private Function CalculateTotalInvoiceAmount(ByVal dblFreight As Double) As Double
        ''Changes done By Ashutosh on 12-01-2007, Issue Id:19339, Include freight tax in invoice calculation.
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateTotalInvoiceAmount
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim lintLoopCounter As Short
        Dim ldblRate As Double
        Dim ldblQuantity As Double
        Dim ldblFinalAmount As Double
        Dim ldblTotalGSTTAXTES As Double = 0
        Dim blnGSTTAXroundoff As Boolean
        Dim intGSTTAXroundoff_decimal As Short
        Dim rsParameterData As ClsResultSetDB

        Dim strParamQuery = "select GSTTAX_ROUNDOFF,GSTTAX_ROUNDOFF_DECIMAL FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnGSTTAXroundoff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTTAXroundoff_decimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing


        CalculateTotalInvoiceAmount = 0
        ldblQuantity = 0
        ldblRate = 0
        ldblFinalAmount = 0
        With SpChEntry
            For lintLoopCounter = 1 To .MaxRows
                .Row = lintLoopCounter
                .Col = 3
                ldblRate = Val(.Text)
                .Col = 5
                ldblQuantity = Val(.Text)
                ldblFinalAmount = ldblFinalAmount + Val(CStr(ldblQuantity * ldblRate))

                If blnGSTTAXroundoff = True Then
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "CGST", False)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "SGST", False)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "UTGST", False)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "IGST", False)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "GSTCC", False)
                Else
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "CGST", False), intGSTTAXroundoff_decimal)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "SGST", False), intGSTTAXroundoff_decimal)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "UTGST", False), intGSTTAXroundoff_decimal)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "IGST", False), intGSTTAXroundoff_decimal)
                    ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, Val(CStr(ldblQuantity * ldblRate)), "GSTCC", False), intGSTTAXroundoff_decimal)
                End If

            Next
        End With
        CalculateTotalInvoiceAmount = ldblFinalAmount + dblFreight + ldblTotalGSTTAXTES
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        CalculateTotalInvoiceAmount = 0
    End Function
    Private Function RoundInvTables(ByVal Doc_No As Integer, ByVal Loc_Code As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Author        :  Davinder Singh
        'Creation Date :  01 Dec 2006
        'Return        :  Boolean
        'Issue ID      :  19165
        'Purpose       :  To rondoff the saved data in SalesChallan_Dtl and Sales_dtl tables
        '                 according to parameters defined in sales_parameter
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim cmd As ADODB.Command
        cmd = New ADODB.Command
        With cmd
            .ActiveConnection = mP_Connection
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "INVOICE_ROUNDOFF"
            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10))
            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput))
            .Parameters.Append(.CreateParameter("@LOCATION_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10))
            .Parameters.Append(.CreateParameter("@ERROR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
            .Parameters(0).Value = gstrUNITID
            .Parameters(1).Value = Doc_No
            .Parameters(2).Value = Loc_Code
            .Parameters(3).Value = ""

            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            If Len(.Parameters(3).Value) > 0 Then
                Call MsgBox(.Parameters(3).Value, MsgBoxStyle.Information, ResolveResString(100))
                RoundInvTables = False
                cmd = Nothing
                Exit Function
            Else
                RoundInvTables = True
            End If
            cmd = Nothing
        End With
        Exit Function
ErrHandler:
        cmd = Nothing
        RoundInvTables = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Sub DisplayCreditTerm()
        '---------------------------------------------------------------------------------------
        'Author        :  Manoj Kr. Vaish
        'Creation Date :  29 June 2007
        'Return        :  NILL
        'Issue ID      :  19992
        'Purpose       :  To display the credit term for the selected Ref no. from Cust_ord_hdr
        '---------------------------------------------------------------------------------------

        Dim rsCredit As ClsResultSetDB
        Dim strsql As String

        rsCredit = New ClsResultSetDB

        strsql = "select term_payment from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
        strsql = strsql & mstrRefNo & "' and Amendment_No ='" & mstrAmmNo & "' and active_flag = 'A'"
        rsCredit.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCredit.GetNoRows > 0 Then
            lblCreditTerm.Text = IIf(IsDBNull(rsCredit.GetValue("term_payment")), "", rsCredit.GetValue("term_payment"))
            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
            Else
                lblCreditTermDesc.Text = ""
            End If
        Else
            lblCreditTerm.Text = ""
            lblCreditTermDesc.Text = ""
        End If
        rsCredit.ResultSetClose()
        rsCredit = Nothing
    End Sub

    Private Sub dtpDateDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDateDesc.KeyDown
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        If e.KeyCode = System.Windows.Forms.Keys.Return And e.Shift = 0 Then
            CmbInvType.Focus()
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub


End Class